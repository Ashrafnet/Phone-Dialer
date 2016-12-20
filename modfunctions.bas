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
    MsgBox "íÈÏæ Çä ÇáãáÝ ÇáÇæáí áÏÚã ÇáÈíÇäÇÊ ÛíÑ ãæÌæÏ ÇáÑÌÇÁ ÇíÌÇÏå ßãÇ Óíáí :  ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    
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
    FindDB = modDialFolder.OpenDialog(mdiMain.Hwnd, "ãáÝ ÞÇÚÏÉ ÇáÈíÇäÇÊ (*.Émdb)", "*.mdb")
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
    Case tsIdle:               S = "ÛíÑ ãÔÛæá"
    Case tsOffering:           S = "ÊÞÏíã"
    Case tsAccepted:           S = "Êã ÑÝÚ ÓãÇÚÉ ÇáåÇÊÝ"
    Case tsDialTone:           S = "äÛãÉ ÇÊÕÇá"
    Case tsDialing:            S = "íÊÕá.."
    Case tsRingback:           S = "íÑä"
    Case tsBusy:               S = "ãÔÛæá"
    Case tsSpecialInfo:        S = "ãÚáæãÉ ÎÇÕÉ"
    Case tsConnected:          S = "Êã ÇäÔÇÁ ÇÊÕÇá"
    Case tsProceeding:         S = "íÊã ÇáÇä ÇáãÚÇáÌÉ"
    Case tsOnHold:             S = "ãÍÌæÒ"
    Case tsConferenced:        S = "Ëã ÚÞÏ ãÄÊãÑ"
    Case tsOnHoldPendConf:     S = "tsOnHoldPendConf"
    Case tsOnHoldPendTransfer: S = "tsOnHoldPendTransfer"
    Case tsDisconnected:       S = "Êã ÞØÚ ÇáÇÊÕÇá"
    Case tsUnknown:            S = "ÛíÑ ãÚÑæÝ"
  End Select
  TapiStateStr = "ÍÇáÉ ÇáÇÊÕÇá: " & S '& " (" & CStr(Apax1.TapiState) & ")"
End Function

