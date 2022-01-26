#ifndef WINWMI_H
#define WINWMI_H

// ===========================================================================
//       A simple one header solution to interacting with Windows WMI in C++
//
//               Copyright (C) by Konstantin 'cpz' L.
//                  https://github.com/cpz/winwmi
//
// ===========================================================================
/*
MIT License

Copyright (c) 2021 Konstantin 'cpz' L. (https://github.com/cpz/winwmi)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

#define _WIN32_DCOM
#include <comdef.h>
#include <string>
#include <Wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")

enum class WmiType : int
{
    kBool,
    kBstr,
    kUint8,
    kUint32
};

enum class WmiError : int
{
    kNone,
    kFailedCoInitialize,
    kFailedCoInitializeSecurity,
    kFailedCoCreateInstance,
    kFailedConnectServer,
    kFailedCoSetProxyBlanket,
    kFailedGetObject,
    kFailedGetMethod,
    kFailedSpawnInstance,
    kFailedPutVariable,
    kFailedExecMethod,
    kWrongDataType,
    kEmptyClassName,
};

class WinWmi
{
public:
    WinWmi(const std::wstring_view namespace_name,
        const std::optional<std::wstring_view> class_name = std::nullopt,
        const std::optional<std::wstring_view> method_name = std::nullopt)
    {
        method_name_ = method_name.has_value() ? SysAllocString(method_name.value().data()) : nullptr;

        if (class_name.has_value())
        {
            class_name_ = SysAllocString(class_name.value().data());
            class_name_string_ = class_name.value();
        }

        auto wmi_cleanup = [&](const std::optional<WmiError> error)
        {
            last_error_ = error.value_or(WmiError::kNone);
            wbem_service_->Release();
            wbem_locator_->Release();
            CoUninitialize();
        };

        result_ = CoInitializeEx(nullptr, COINIT_MULTITHREADED);

        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedCoInitialize;
            return;
        }

        result_ = CoInitializeSecurity(
            nullptr,
            -1,
            nullptr,
            nullptr,
            RPC_C_AUTHN_LEVEL_DEFAULT,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,
            EOAC_NONE,
            nullptr
        );

        if (FAILED(result_) && result_ != RPC_E_TOO_LATE)
        {
            last_error_ = WmiError::kFailedCoInitializeSecurity;
            CoUninitialize();
            return;
        }

        result_ = CoCreateInstance(CLSID_WbemLocator, nullptr,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator, reinterpret_cast<LPVOID*>(&wbem_locator_));

        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedCoCreateInstance;
            CoUninitialize();
            return;
        }
        result_ = wbem_locator_->ConnectServer(
            _bstr_t(namespace_name.data()),
            nullptr, nullptr, nullptr, 0, nullptr, nullptr, &wbem_service_
        );

        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedConnectServer;
            wbem_locator_->Release();
            CoUninitialize();
            return;
        }

        result_ = CoSetProxyBlanket(
            wbem_service_,
            RPC_C_AUTHN_WINNT,
            RPC_C_AUTHZ_NONE,
            nullptr,
            RPC_C_AUTHN_LEVEL_CALL,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,
            EOAC_NONE
        );

        if (FAILED(result_))
        {
            wmi_cleanup(WmiError::kFailedCoSetProxyBlanket);
            return;
        }

        if (SysStringLen(class_name_) != 0 && SysStringLen(method_name_) != 0)
        {
            result_ = wbem_service_->GetObjectW(class_name_, 0, nullptr, &class_, nullptr);
            if (FAILED(result_))
            {
                wmi_cleanup(WmiError::kFailedGetObject);
                return;
            }

            result_ = class_->GetMethod(method_name_, 0, &param_, nullptr);
            if (FAILED(result_))
            {
                wmi_cleanup(WmiError::kFailedGetMethod);
                return;
            }

            result_ = param_->SpawnInstance(0, &class_instance_);
            if (FAILED(result_))
            {
                wmi_cleanup(WmiError::kFailedSpawnInstance);
                return;
            }
        }
    }

    ~WinWmi()
    {
        if (SysStringLen(class_name_) != 0)
            SysFreeString(class_name_);

        if (SysStringLen(method_name_) != 0)
            SysFreeString(method_name_);

        if (class_)
            class_->Release();

        if (class_instance_)
            class_instance_->Release();

        if (param_)
            param_->Release();

        if (wbem_locator_)
            wbem_locator_->Release();

        if (wbem_service_)
            wbem_service_->Release();

        CoUninitialize();
    }

    [[nodiscard]] std::optional<std::vector<std::wstring>>
        get_all(const std::wstring_view variable,
            const std::optional<std::wstring_view> class_name = std::nullopt)
    {
        std::vector<std::wstring> objects = {};

        if (!is_class_name_valid(class_name))
        {
            last_error_ = WmiError::kEmptyClassName;
            return std::nullopt;
        }

        auto query = create_query(class_name);
        auto [hr, enumerator] = exec_query(query);
        if (FAILED(hr))
            return std::nullopt;

        IWbemClassObject* class_object = nullptr;
        ULONG u_return = 0;

        while (enumerator)
        {
            hr = enumerator->Next(WBEM_INFINITE, 1,
                &class_object, &u_return);

            if (u_return == 0)
                break;

            VARIANT vt_prop = {};
            hr = class_object->Get(variable.data(), 0, &vt_prop, nullptr, nullptr);
            if (hr != S_OK)
                break;

            if (SysStringLen(vt_prop.bstrVal) != 0)
                objects.emplace_back(vt_prop.bstrVal);

            VariantClear(&vt_prop);
            class_object->Release();
        }

        return objects;
    }

    [[nodiscard]] std::optional<std::wstring>
        get(const std::wstring_view variable,
            const std::optional<std::wstring_view> class_name = std::nullopt)
    {
        std::optional<std::wstring> value = {};
        if (!is_class_name_valid(class_name))
        {
            last_error_ = WmiError::kEmptyClassName;
            return std::nullopt;
        }

        auto query = create_query(class_name);
        auto [hr, enumerator] = exec_query(query);
        if (FAILED(hr))
            return std::nullopt;

        IWbemClassObject* class_object = nullptr;
        ULONG u_return = 0;

        while (enumerator)
        {
            hr = enumerator->Next(WBEM_INFINITE, 1,
                &class_object, &u_return);

            if (u_return == 0)
                break;

            VARIANT vt_prop = {};
            hr = class_object->Get(variable.data(), 0, &vt_prop, nullptr, nullptr);
            if (hr != S_OK)
                break;

            if (SysStringLen(vt_prop.bstrVal) == 0)
                value = std::nullopt;
            else
                value = vt_prop.bstrVal;

            VariantClear(&vt_prop);
            class_object->Release();
        }

        return value;
    }

    template <typename T>
    bool get(const std::wstring_view variable,
        const WmiType type,
        T& value,
        const std::optional<std::wstring_view> class_name = std::nullopt)
    {
        if (is_class_name_valid(class_name))
        {
            last_error_ = WmiError::kEmptyClassName;
            return false;
        }

        auto query = create_query(class_name);
        auto [hr, enumerator] = exec_query(query);
        if (FAILED(hr))
            return false;

        IWbemClassObject* class_object = nullptr;
        ULONG u_return = 0;

        while (enumerator)
        {
            enumerator->Next(WBEM_INFINITE, 1,
                &class_object, &u_return);

            if (u_return == 0)
                break;

            VARIANT vt_prop = {};
            hr = class_object->Get(variable.data(), 0, &vt_prop, nullptr, nullptr);
            if (hr != S_OK)
                break;

            switch (type)
            {
            case WmiType::kBool:
                value = vt_prop.boolVal;
                break;

            case WmiType::kUint8:
                value = vt_prop.uintVal;
                break;

            case WmiType::kUint32:
                value = vt_prop.uintVal;
                break;

            default:
                last_error_ = WmiError::kWrongDataType;
                value = std::nullopt;
                break;
            }

            VariantClear(&vt_prop);
            class_object->Release();
        }

        return true;
    }

    bool set(const std::wstring_view variable,
        const std::wstring_view value,
        const std::optional<std::wstring_view> class_name = std::nullopt)
    {
        if (is_class_name_valid(class_name))
        {
            last_error_ = WmiError::kEmptyClassName;
            return false;
        }

        BSTR exec_class = nullptr;
        if (class_name.has_value())
            exec_class = SysAllocString(class_name.value().data());
        else
            exec_class = class_name_;

        VARIANT var_cmd{};
        var_cmd.vt = VT_BSTR;
        var_cmd.bstrVal = _bstr_t(value.data());

        result_ = class_instance_->Put(variable.data(), 0, &var_cmd, 0);
        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedPutVariable;
            return false;
        }

        IWbemClassObject* out_params = nullptr;
        result_ = wbem_service_->ExecMethod(exec_class,
            method_name_,
            0,
            nullptr,
            class_instance_,
            &out_params,
            nullptr);

        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedExecMethod;
            return false;
        }

        SysFreeString(exec_class);
        VariantClear(&var_cmd);
        if (out_params)
            out_params->Release();

        return true;
    }

    template <typename T>
    bool set(const std::wstring_view variable,
        WmiType type,
        T value,
        const std::optional<std::wstring_view> class_name = std::nullopt)
    {
        if (is_class_name_valid(class_name))
        {
            last_error_ = WmiError::kEmptyClassName;
            return false;
        }

        BSTR exec_class = nullptr;
        if (class_name.has_value())
            exec_class = SysAllocString(class_name.value().data());
        else
            exec_class = class_name_;

        VARIANT var_cmd{};

        switch (type)
        {
        case WmiType::kBool:
            var_cmd.vt = VT_BOOL;
            var_cmd.boolVal = value;
            break;

        case WmiType::kUint8:
            var_cmd.vt = VT_UI1;
            var_cmd.uintVal = value;
            break;

        case WmiType::kUint32:
            var_cmd.vt = VT_UI4;
            var_cmd.uintVal = value;
            break;

        default:
            last_error_ = WmiError::kWrongDataType;
            return false;
        }

        result_ = class_instance_->Put(variable.data(), 0, &var_cmd, 0);
        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedPutVariable;
            return false;
        }

        IWbemClassObject* out_params = nullptr;
        result_ = wbem_service_->ExecMethod(exec_class,
            method_name_,
            0,
            nullptr,
            class_instance_,
            &out_params,
            nullptr);

        if (FAILED(result_))
        {
            last_error_ = WmiError::kFailedExecMethod;
            return false;
        }

        SysFreeString(exec_class);
        VariantClear(&var_cmd);
        if (out_params)
            out_params->Release();

        return true;
    }

    [[nodiscard]] WmiError GetLastError() const { return last_error_; }

private:
    inline std::wstring create_query(std::optional<std::wstring_view> class_name)
    {
        std::wstring query = L"SELECT * FROM ";
        if (class_name.has_value())
            query += class_name.value().data();
        else
            query += class_name_;

        return query;
    }

    inline bool is_class_name_valid(std::optional<std::wstring_view> class_name)
    {
        return class_name.has_value() && SysStringLen(class_name_) != 0;
    }

    inline std::tuple<HRESULT, IEnumWbemClassObject*> exec_query(const std::wstring& query)
    {
        IEnumWbemClassObject* enumerator = nullptr;
        auto hr = wbem_service_->ExecQuery(
            _bstr_t("WQL"),
            _bstr_t(query.c_str()),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            nullptr,
            &enumerator
        );

        return { hr, enumerator };
    }

private:
    WmiError last_error_ = WmiError::kNone;
    HRESULT result_ = 0;

    IWbemServices* wbem_service_ = nullptr;
    IWbemLocator* wbem_locator_ = nullptr;

    IWbemClassObject* class_ = nullptr;
    IWbemClassObject* param_ = nullptr;
    IWbemClassObject* class_instance_ = nullptr;

    BSTR method_name_ = nullptr;
    BSTR class_name_ = nullptr;
    std::wstring class_name_string_;
};

#endif
