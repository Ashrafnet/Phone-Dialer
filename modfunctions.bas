Attribute VB_Name = "modfunctions"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public db As New ADODB.Connection

'waits a number of seconds
Public Sub Wait(sngSeconds As Single)

    Dim sngEndTime As Single
    sngEndTime = Timer + sngSeconds

    While Timer < sngEndTime
        DoEvents
    Wend

End Sub


Function ConnectToDB(DataName As String) As Boolean
    On Error GoTo FindErr
    ConnectToDB = True
    Dim strQ As String ' query string
    
    DBPath = GetSetting(App.EXEName, "DB", "Path")
    If DBPath = "" Then
        DBPath = App.Path & "\" & DataName
    Else
        DBPath = Trim$(DBPath)
    End If
    strQ = "Provider=Microsoft.Jet.OLEDB.4.0;password= ;User ID=Admin;Data source=" & DBPath
    db.Open strQ           ' Establish Connection with DB

    Exit Function
FindErr:
    MsgBox "���� �� ����� ������ ���� �������� ��� ����� ������ ������ ��� ���� :  ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    
    ' If the database isn't found, use the FindDB function to find it.
    If Err.Number = -2147467259 Then
        On Error Resume Next
        Connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;password=;User ID=Admin;Data source=" & FindDB("db.mdb")
        Resume Next
    End If

    ConnectToDB = False
End Function
Private Function FindDB(dbName As String) As String
    On Error GoTo ErrHandler

    ' Configure cmdDialog in case the database can't be found.
    FindDB = modDialFolder.OpenDialog(mdiMain.Hwnd, "��� ����� �������� (*.�mdb)", "*.mdb")
    If FindDB = "" Then End
    SaveSetting App.EXEName, "DB", "Path", FindDB
    DBPath = FindDB
    Exit Function
ErrHandler:
    MsgBox Err.Description
    'Resume
End Function

Sub LogCall(phoneNo As String, UserName As String, status As Integer, WaveFiles As String, ContactID As Long)
On Error GoTo er:
    Dim rs As New Recordset
    
    Dim Thedate As Date
    Thedate = Date
    GetCurrentUserName
    rs.Open "insert into log (thephone,thename,thedate,thetime,status,log_user,compname,wavsend,uid) values('" & phoneNo & "','" & UserName & "','" & Thedate & "','" & time$ & "','" & status & "','" & GetCurrentUserName & "','" & ComputerName & "','" & WaveFiles & "','" & ContactID & "')", db
    Exit Sub
er:
    MsgBox Err.Description
End Sub
Public Function TapiStateStr(State As TxTapiState) As String
  Dim S As String
  Select Case State
    Case tsIdle:               S = "��� �����"
    Case tsOffering:           S = "�����"
    Case tsAccepted:           S = "�� ��� ����� ������"
    Case tsDialTone:           S = "���� �����"
    Case tsDialing:            S = "����.."
    Case tsRingback:           S = "���"
    Case tsBusy:               S = "�����"
    Case tsSpecialInfo:        S = "������ ����"
    Case tsConnected:          S = "�� ����� �����"
    Case tsProceeding:         S = "��� ���� ��������"
    Case tsOnHold:             S = "�����"
    Case tsConferenced:        S = "�� ��� �����"
    Case tsOnHoldPendConf:     S = "tsOnHoldPendConf"
    Case tsOnHoldPendTransfer: S = "tsOnHoldPendTransfer"
    Case tsDisconnected:       S = "�� ��� �������"
    Case tsUnknown:            S = "��� �����"
  End Select
  TapiStateStr = "���� �������: " & S '& " (" & CStr(Apax1.TapiState) & ")"
End Function

