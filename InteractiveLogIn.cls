VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InteractiveLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================================================
'
'   see notes in:
'   http://msdn.microsoft.com/library/default.asp?url=/library/en-us/security/accclsrv_0jle.asp
'
'==============================================================================================

Private Const PROVIDER_DEFAULT          As Long = 0
Private Const PROVIDER_WINNT35          As Long = 1
Private Const LOGON_INTERACTIVE         As Long = 2
Private Const LOGON_BATCH               As Long = 4
Private Const LOGON_SERVICE             As Long = 5

Private Const MESSAGE_FROM_SYSTEM       As Long = &H1000
Private Const MESSAGE_IGNORE_INSERTS    As Long = &H200
Private Const MESSAGE_FROM_API_CALL     As Long = MESSAGE_FROM_SYSTEM Or MESSAGE_IGNORE_INSERTS

Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private mlngTokenHandle                 As Long

Public Function LogInAs(ByVal strUserName As String, _
                        ByVal strPassword As String, _
                        ByVal strDomain As String) As Boolean

    Dim strErrorText        As String
    Dim lngErrorNumber      As Long
    If Not (mlngTokenHandle = 0) Then LogOut
    If RevertToSelf Then
        If LogonUser(strUserName, strDomain, strPassword, LOGON_INTERACTIVE, PROVIDER_DEFAULT, mlngTokenHandle) Then
            LogInAs = ImpersonateLoggedOnUser(mlngTokenHandle)
        End If
    End If
    If Not LogInAs Then
        lngErrorNumber = Err.LastDllError
        strErrorText = ApiErrorText(lngErrorNumber)
        Err.Raise lngErrorNumber, App.EXEName & "." & TypeName(Me) & ".LogInAs()", strErrorText
    End If
End Function

Public Sub LogOut()
    Dim strErrorText        As String
    Dim lngErrorNumber      As Long
    CloseHandle mlngTokenHandle
    If RevertToSelf = 0 Then
        lngErrorNumber = Err.LastDllError
        strErrorText = ApiErrorText(lngErrorNumber)
        Err.Raise lngErrorNumber, App.EXEName & "." & TypeName(Me) & ".LogOut()", strErrorText
    End If
End Sub

Private Function ApiErrorText(ByVal lngErrorNumber As Long) As String
    Dim strApiErrorText     As String
    Dim lngApiErrorTextLen  As Long
    strApiErrorText = Space(256)
    lngApiErrorTextLen = FormatMessage(MESSAGE_FROM_API_CALL, 0&, lngErrorNumber, 0&, strApiErrorText, 256, 0)
    ApiErrorText = Trim$(Replace(Left$(strApiErrorText, lngApiErrorTextLen), vbCrLf, " "))
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    LogOut
End Sub
