
'  ___ __ _ _ __ ___  _ __  _   _ ___ _ __ _____   ___ __   ___ ___ 
' / __/ _` | '_ ` _ \| '_ \| | | / __| '_ ` _ \ \ / | '_ \ / _ / __|
'| (_| (_| | | | | | | |_) | |_| \__ | | | | | \ V /| |_) |  __\__ \
' \___\__,_|_| |_| |_| .__/ \__,_|___|_| |_| |_|\_/ | .__(_\___|___/
                    '|_|                            |_|             

'					www.campusmvp.es
					
'WINHTTP ACTIVEX OBJECT CONSTANTS DEFINITION FOR SCRIPTING
'
'Global variables that define all the constants available for scripting the native WinHTTPObject with VBScript

'This information is not easy to get since Microsoft doesn't provide any similar file or document detailing these values.
'Just add this file to your Windows Scripting Host or Classic ASP project to be able to easily use all the features of WinHTTP.

'NOTE= This file must be encoded in ANSI or in UTF-8 without BOM, because WSH gets an error if the BOM is included in UTF-8.

'Find all the information about WinHTTP at: https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106.aspx

'Manually compiled by Jose M. Alarcon - www.jasoft.org


Const WinHttpRequestOption_UserAgentString = 0
Const WinHttpRequestOption_URL = 1
Const WinHttpRequestOption_URLCodePage = 2
Const WinHttpRequestOption_EscapePercentInURL = 3
Const WinHttpRequestOption_SslErrorIgnoreFlags = 4
Const WinHttpRequestOption_SelectCertificate = 5
Const WinHttpRequestOption_EnableRedirects = 6
Const WinHttpRequestOption_UrlEscapeDisable = 7
Const WinHttpRequestOption_UrlEscapeDisableQuery = 8
Const WinHttpRequestOption_SecureProtocols = 9
Const WinHttpRequestOption_EnableTracing = 10
Const WinHttpRequestOption_RevertImpersonationOverSsl = 11
Const WinHttpRequestOption_EnableHttpsToHttpRedirects = 12
Const WinHttpRequestOption_EnablePassportAuthentication = 13
Const WinHttpRequestOption_MaxAutomaticRedirects = 14
Const WinHttpRequestOption_MaxResponseHeaderSize = 15
Const WinHttpRequestOption_MaxResponseDrainSize = 16
Const WinHttpRequestOption_EnableHttp1_1 = 17
Const WinHttpRequestOption_EnableCertificateRevocationCheck= 18

Const WinHttpRequestSecureProtocols_SecureProtocol_ALL = 168
Const WinHttpRequestSecureProtocols_SecureProtocol_SSL2 = 8
Const WinHttpRequestSecureProtocols_SecureProtocol_SSL3 = 32
Const WinHttpRequestSecureProtocols_SecureProtocol_TLS1 = 128


Const WinHttpRequestSslErrorFlags_SslErrorFlag_CertCNInvalid = 4096
Const WinHttpRequestSslErrorFlags_SslErrorFlag_CertDateInvalid = 8192
Const WinHttpRequestSslErrorFlags_SslErrorFlag_CertWrongUsage = 512
Const WinHttpRequestSslErrorFlags_SslErrorFlag_Ignore_All = 13056
Const WinHttpRequestSslErrorFlags_SslErrorFlag_UnknownCA = 256

Const WinHttpRequestAutoLogonPolicy_Always = 0
Const WinHttpRequestAutoLogonPolicy_Never = 2
Const WinHttpRequestAutoLogonPolicy_OnlyIfBypassProxy = 1

'Example:

'	Dim obj 
'	Set obj = WScript.CreateObject("WinHttp.WinHttpRequest.5.1")
'	obj.open "GET", "http://www.google.com", false
	'Enable redirection from HTTPS to HTTP (is Off by default)
'	obj.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = true
'	obj.send()
