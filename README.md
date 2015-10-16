#WinHTTP Constants for Scripting

**Microsoft Windows HTTP Services ([WinHTTP](https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106.aspx))** provides developers with an HTTP client application programming interface (API) to send requests through the HTTP protocol to other HTTP servers.

The WinHttpRequest object reference on MSDN, although is aimed at scripting developers (JScript, VBScript, Classic ASP, VB6, Visual Basic for Applications...) **doesn't include any indications about the values of constants** and enumerations. The only available reference files are written in C, and the constants in there are not compatible with the ones in ActiveX objects.

Because of this fact, there have been many frustrated developers since 2001.

A long time ago I compiled the correct values and definitions of those constants and now I'm sharing them here for you to use.

There are two versions of these constants:

- **WinHTTP_Constants.js**: for JavaScript developers. The constants are defined as global  variables. There are 4 classes defined with all the constants available.
- **WinHTTP_Constants.vbs**: for VBScript developers and Visual Basic developers. In this case the constants are defined as individual global const vars.

Both include simple examples of using WinHttp.

Hope this helps!

Manually compiled by **Jose M. Alarcon** - www.jasoft.org

```
  ___ __ _ _ __ ___  _ __  _   _ ___ _ __ _____   ___ __   ___ ___ 
 / __/ _` | '_ ` _ \| '_ \| | | / __| '_ ` _ \ \ / | '_ \ / _ / __|
| (_| (_| | | | | | | |_) | |_| \__ | | | | | \ V /| |_) |  __\__ \
 \___\__,_|_| |_| |_| .__/ \__,_|___|_| |_| |_|\_/ | .__(_\___|___/
                    |_|                            |_|             
```

**[www.campusmvp.es](http://www.campusmvp.es)**