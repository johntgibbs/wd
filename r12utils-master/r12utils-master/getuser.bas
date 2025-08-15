Attribute VB_Name = "GetUser"
Option Explicit
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nsize As Long) As Long
