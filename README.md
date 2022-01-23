# 🤡  WinWMI

[![Sponsor](https://img.shields.io/badge/💜-sponsor-blueviolet)](https://github.com/sponsors/cpz)

A simple one header solution to interacting with [Windows WMI](https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page "Windows WMI") in C++

## Usage

Just `#include "winwmi.hpp"` in your code!

### Initialize

To initialize your winwmi you basically should do:

```C++
const auto win_wmi = std::make_unique<WinWmi>(
        L"root\\CIMV2", // WMI Namespace
        L"Win32_Process" // Optional WMI Class
        // Optional WMI Method to set data (ex. L"Set")
    );

if (const auto error = win_wmi->GetLastError(); error != WmiError::kNone)
    throw std::exception(fmt::format("WinWmi Failed with code: {}\n", static_cast<int>(error)).c_str());
```

### Retrieving information from class

```C++
if (auto res = win_wmi->get_all(L"Caption"); res.has_value())
{
    for (const auto& element : res.value())
    {
        fmt::print(L"[WnWMI] Caption: {}\n", element);
    }
} 
```

### WinWmi Errors

WinWmi has got his own error handling which can be accessed via 

```C++
win_wmi->GetLastError()
```

If you want to see what code does mean then you can look in to WmiError enum. Their names should describe what exactly happened and what went wrong.

## License

```
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
```