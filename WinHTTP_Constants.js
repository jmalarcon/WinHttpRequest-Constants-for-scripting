/*
  ___ __ _ _ __ ___  _ __  _   _ ___ _ __ _____   ___ __   ___ ___ 
 / __/ _` | '_ ` _ \| '_ \| | | / __| '_ ` _ \ \ / | '_ \ / _ / __|
| (_| (_| | | | | | | |_) | |_| \__ | | | | | \ V /| |_) |  __\__ \
 \___\__,_|_| |_| |_| .__/ \__,_|___|_| |_| |_|\_/ | .__(_\___|___/
                    |_|                            |_|             

					www.campusmvp.es
					
WINHTTP ACTIVEX OBJECT CONSTANTS DEFINITION FOR SCRIPTING

Global variables that define 4 classes with all the constants available for scripting the native WinHTTPObject with JavaScript

This information is not easy to get since Microsoft doesn't provide any similar file or document detailing these values.
Just add this file to your Windows Scripting Host or Classic ASP project to be able to easily use all the features of WinHTTP.

NOTE: This file must be encoded in ANSI or in UTF-8 without BOM, because WSH gets an error if the BOM is included in UTF-8.

Find all the information about WinHTTP at: https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106.aspx

Manually compiled by Jose M. Alarcon - www.jasoft.org
*/
var WinHttpRequestOption = {
	UserAgentString : 0,
	URL : 1,
	URLCodePage : 2,
	EscapePercentInURL : 3,
	SslErrorIgnoreFlags : 4,
	SelectCertificate : 5,
	EnableRedirects : 6,
	UrlEscapeDisable : 7,
	UrlEscapeDisableQuery : 8,
	SecureProtocols : 9,
	EnableTracing : 10,
	RevertImpersonationOverSsl : 11,
	EnableHttpsToHttpRedirects : 12,
	EnablePassportAuthentication : 13,
	MaxAutomaticRedirects : 14,
	MaxResponseHeaderSize : 15,
	MaxResponseDrainSize : 16,
	EnableHttp1_1 : 17,
	EnableCertificateRevocationCheck: 18
};

var WinHttpRequestSecureProtocols = {
	SecureProtocol_ALL : 168,
	SecureProtocol_SSL2 : 8,
	SecureProtocol_SSL3 : 32,
	SecureProtocol_TLS1 : 128
};

var WinHttpRequestSslErrorFlags = {
	SslErrorFlag_CertCNInvalid : 4096,
	SslErrorFlag_CertDateInvalid : 8192,
	SslErrorFlag_CertWrongUsage : 512,
	SslErrorFlag_Ignore_All : 13056,
	SslErrorFlag_UnknownCA : 256
};

var WinHttpRequestAutoLogonPolicy = {
	AutoLogonPolicy_Always : 0,
	AutoLogonPolicy_Never : 2,
	AutoLogonPolicy_OnlyIfBypassProxy : 1
};

/*
Example:

	var obj = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
	obj.open("GET", "http://www.google.com", false);
	//Enable redirection from HTTPS to HTTP (is Off by default)
	obj.Option(WinHttpRequestOption.EnableHttpsToHttpRedirects) = true;
	obj.send();

*/