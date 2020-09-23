Attribute VB_Name = "ICQBAS"
'=============================================================================================================
' ICQ API Functions TRANSLATED FROM C++ , for more information visit www.mirabilis.com/api
' The functions use ICQMAPI.DLL that should be in your windows\system folder !
' Michael Belenky 2001 (c)
'=============================================================================================================

Declare Function SetLicenseKey Lib "icqmapi.dll" Alias "ICQAPICall_SetLicenseKey" (ByVal pszName As String, ByVal pszPassword As String, ByVal pszLicense As String) As Boolean
Declare Function SetOwnerState Lib "icqmapi.dll" Alias "ICQAPICall_SetOwnerState" (ByVal iState As Long) As Boolean

Global Const BICQAPI_USER_STATE_ONLINE = 0
Global Const BICQAPI_USER_STATE_CHAT = 1
Global Const BICQAPI_USER_STATE_AWAY = 2
Global Const BICQAPI_USER_STATE_NA = 3
Global Const BICQAPI_USER_STATE_OCCUPIED = 4
Global Const BICQAPI_USER_STATE_DND = 5
Global Const BICQAPI_USER_STATE_INVISIBLE = 6
Global Const BICQAPI_USER_STATE_OFFLINE = 7


'BOOL WINAPI ICQAPICall_SendFile(int iUIN, char *pszFileNames);
'(?)
Declare Function SendFile Lib "icqmapi.dll" Alias "ICQAPICall_SendFile" (iUin As Long, pszFileNames As String) As Boolean


Declare Function SendMessage Lib "icqmapi.dll" Alias _
     "ICQAPICall_SendMessage" (ByVal iUin As Long, _
                               ByVal pszMessage As String) As Boolean
                              
Declare Function SendExternal Lib "icqmapi.dll" Alias _
     "ICQAPICall_SendExternal" (ByVal iUin As Long, _
                                ByVal pszExternal As String, _
                                ByVal pszMessage As String, _
                                ByVal bAutoSend As Long) As Boolean

Declare Function SendURL Lib "icqmapi.dll" Alias _
     "ICQAPICall_SendURL" (ByVal iUin As Long, _
                           ByVal pszURL As String) As Boolean



