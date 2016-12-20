Attribute VB_Name = "functions"

Public Declare Function GetMenuItemCount& Lib "user32" (ByVal hMenu As Long)
Public Declare Function GetSystemMenu& Lib "user32" (ByVal Hwnd As Long, ByVal bRevert As Long)
Public Declare Function RemoveMenu& Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long)
Public Declare Function DrawMenuBar& Lib "user32" (ByVal Hwnd As Long)

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function BringWindowToTop Lib "user32.dll" (ByVal Hwnd As Long) As Long

Const LVM_FIRST = &H1000
Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Const LVS_EX_FULLROWSELECT = &H20

Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000
Public CurrentUserName As String
Private Declare Function GetUserName Lib "advapi32.DLL" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Public IsAdmin As Boolean
Public IsfrmMainLoaded As Boolean
Public Function Checknombers(Numbers As String) As Boolean
    Dim No() As String
    No = Split(Numbers, ",", , vbTextCompare)
    For Each y In No()
        If Len(y) <> 10 Then
            Checknombers = False
            Exit Function
        Else
            If VBA.Left(y, 4) = "0599" Then
                Checknombers = True
            Else
                Checknombers = False
            End If
        End If
        If IsNumeric(y) Then
            Checknombers = True
        Else
            Checknombers = 0
        End If
        ' now check the spaces
        For i = 1 To Len(y)
            If Mid(y, i, 1) = (" ") Then Checknombers = 0: Exit Function
        Next i
    Next

End Function
Public Function Checknomber(Number As String) As Boolean
    Dim y
    y = Number
    If Len(y) <> 10 Then
        Checknomber = False
        Exit Function
    Else
        If VBA.Left(y, 4) = "0599" Then
            Checknomber = True
        Else
            Checknomber = False
        End If
    End If
    If IsNumeric(y) Then
        Checknomber = True
    Else
        Checknomber = 0
    End If
    ' now check the spaces
    For i = 1 To Len(y)
        If Mid(y, i, 1) = (" ") Then Checknomber = 0: Exit Function
    Next i

End Function

Public Sub RemoveSysMenuX(frm As Form)
   Dim hMenu As Long
   Dim lngMnuCount As Long

   hMenu = GetSystemMenu(frm.Hwnd, 0)
  
   lngMnuCount = GetMenuItemCount(hMenu)
   RemoveMenu hMenu, lngMnuCount - 1, MF_REMOVE Or MF_BYPOSITION
   RemoveMenu hMenu, lngMnuCount - 2, MF_REMOVE Or MF_BYPOSITION
   DrawMenuBar frm.Hwnd
End Sub
Public Sub LVFullRowSelect(lstvw As ListView)
   Dim rs As Long
   Dim R As Long
   rs = SendMessageLong(lstvw.Hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
   rs = rs Or LVS_EX_FULLROWSELECT
   R = SendMessageLong(lstvw.Hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rs)
End Sub

Function ReadINI(Section$, KeyName$, Filename$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   ReadINI = Left(RetStr, GetPrivateProfileString(Section$, ByVal KeyName$, "", RetStr, Len(RetStr), Filename$))
End Function

Public Sub WriteINI(Filename As String, Section As String, key As String, Text As String)
    WritePrivateProfileString Section, key, Text, Filename
End Sub
Sub StartWithWindows(Optional start As Boolean = True)
 Const WinInfo = "Software\Microsoft\Windows\CurrentVersion"
 Const WinInfoNT = "Software\Microsoft\Windows NT\CurrentVersion"
    If start = True Then
        If GetWindows Then
             SaveString HKEY_LOCAL_MACHINE, WinInfo & "\Run", "VLS", App.Path & "\" & App.EXEName & ".exe /min"
        Else
             SaveString HKEY_LOCAL_MACHINE, WinInfoNT & "\Run", "VLS", App.Path & "\" & App.EXEName & ".exe /min"
        End If
    Else
        If GetWindows Then
             SaveString HKEY_LOCAL_MACHINE, WinInfo & "\Run", "VLS", ""
        Else
             SaveString HKEY_LOCAL_MACHINE, WinInfoNT & "\Run", "VLS", ""
        End If
    End If

End Sub
Function GetWindows() As Boolean 'Get is NT or not
Dim A, vKey As String
 Const WinInfoNT = "Software\Microsoft\Windows NT\CurrentVersion"

A = GetString(HKEY_LOCAL_MACHINE, WinInfoNT, vKey)
If vKey <> "" Then
   GetWindows = False 'is Windows NT
Else
   GetWindows = True 'is Windows
End If
End Function
Function GetPathTemp() As String
    Dim strTemp As String, strUserName As String
    'Create a buffer
    strTemp = String(100, Chr$(0))
    'Get the temporary path
    GetTempPath 100, strTemp
    'strip the rest of the buffer
    strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
    GetPathTemp = strTemp
End Function

Function GetCurrentUserName() As String
     'Create a buffer
     Dim strUserName As String
    strUserName = String(100, Chr$(0))
    'Get the username
    GetUserName strUserName, 100
    'strip the rest of the buffer
    GetCurrentUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
    
End Function
Function ComputerName() As String
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    ComputerName = strString
End Function
Function MaxDayInMonth(TheMonth As Integer, Optional TheYear As Integer) As Integer
On Error GoTo er:
    Select Case TheMonth
        Case 4, 6, 9, 11
            MaxDayInMonth = 30
        Case 1, 3, 5, 7, 8, 10, 12
            MaxDayInMonth = 31
        Case 2
            If TheYear = 0 Then TheYear = Year(Date)
            If TheYear Mod 4 = 0 Then
                MaxDayInMonth = 29
            Else
                MaxDayInMonth = 28
            End If
    End Select
    Exit Function
er:
    MsgBox Err.Description
End Function
Sub AboutSystem()
    frmAbout.Show 1
    'MsgBox " „  »—„Ã… «·»—‰«„Ã ⁄‰ ÿ—Ìﬁ:" & vbNewLine & "   „.«‘—› ﬂ„«· «·ﬁ’«’ " & vbNewLine & "AshrafNet4u@HotMail.Com" & vbNewLine & "Aqssass@Ccast.Edu.Ps", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
End Sub
Sub ShowHelp()
    MsgBox " Õ  «·«‰‘«¡", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
End Sub

