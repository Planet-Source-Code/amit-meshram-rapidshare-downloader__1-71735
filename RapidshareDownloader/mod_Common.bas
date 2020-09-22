Attribute VB_Name = "mod_Common"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public FileSize As Long
Dim SecondServer As String
Dim SharedFileNameURL  As String
Dim SharedFileName As String
Public Bool1 As Boolean
Public Bool2 As Boolean

Dim SharedGetFileExt As String

Dim ThirdServer As String
Dim counter As Integer

Dim tmp1, tmp2, tmp3, tmp4, tmp5 As String
Dim tmp6, tmp7, tmp8, tmp9, tmp10 As String

Dim tmp11, tmp12, tmp13 As String
Dim tmpsize1, tmpsize2, tmpsize3 As String

Sub GetInfo1(Inet1 As Inet, URL As String)
    Dim Res As String
    Dim Res1 As String
    
    Res = Inet1.OpenURL(URL)
    
    'For File Size
    If InStr(Res, "/files/") Then
        strpos3 = InStr(Res, "/files/")
        tmpsize1 = Mid(Res, InStr(1, Res, "|") + 1)
        tmpsize2 = Left(tmpsize1, InStr(1, tmpsize1, "KB") + 1)
        tmpsize3 = Replace(tmpsize2, "KB", "")
        Debug.Print Trim(tmpsize3)
    End If
    
    If InStr(Res, "<form action=") Then
        strpos1 = InStr(Res, "<form action=")
        tmp1 = Mid(Res, InStr(1, Res, "<form action=") + 1)
        tmp2 = Mid(tmp1, 14, Len(Res))
        tmp3 = Mid(tmp2, 1, InStr(1, tmp2, Chr(&H22)) - 1)
        
        'Second Server Name
        tmp4 = Mid(tmp3, 8, InStr(1, tmp3, "/files") - 8)
        SecondServer = tmp4
        'Debug.Print tmp4
        
        'Original file URL
        SharedFileNameURL = Trim(tmp3)
        'Debug.Print Trim(tmp3)
        
        'For Posting Value from /files....
        tmp5 = Mid(tmp3, InStr(1, tmp3, ".com") + 4)
        'Debug.Print tmp5
        
        'Zip/Rar file Name
        tmp4 = Mid(tmp3, InStr(40, tmp3, "/") + 1)
        SharedFileName = Trim(tmp4)
        'Debug.Print SharedFileName
        
        'Save File Name
        SharedGetFileExt = Left(tmp4, Len(tmp4) + 3)
        'Debug.Print SharedGetFileExt
        
        Do While Inet1.StillExecuting
            DoEvents
        Loop
    End If
    Call GetInfo2(Inet1)
End Sub

Sub GetInfo2(Inet2 As Inet)
    Dim Res1
    Dim Cnt As Integer
    
    Inet2.Execute "http://" & SecondServer & tmp5, "POST", "dl.start=Free", "Content-Type: application/x-www-form-urlencoded"
    Do While Inet2.StillExecuting
        DoEvents
    Loop
    Res1 = Inet2.GetChunk(8024, icString)
    
    If InStr(Res1, "<form name=") Then
        strpos2 = InStr(Res1, "<form name=")
        tmp6 = Mid(Res1, InStr(1, Res1, "<form name=") + 1)
        tmp7 = Mid(tmp6, 25, Len(Res1))
        tmp8 = Mid(tmp7, 1, InStr(1, tmp7, Chr(&H22)) - 1)
        
        'Third Server Founded
        ThirdServer = tmp8
        Debug.Print ThirdServer
        frmMain.Timer1.Enabled = True
    End If
End Sub

Sub DownloadCreate()
    Dim FileNumber As Integer
    Dim FileData() As Byte
    Dim FileSize  As Long
    Dim FileRemaining As Long
    
    frmMain.Inet3.Execute ThirdServer, "POST", "mirror=on&x=44&y=34", "Content-Type: application/x-www-form-urlencoded"
    
    Do While frmMain.Inet3.StillExecuting
        DoEvents
    Loop
    
    'FileSize = tmpsize3 * 1024
    FileSize = frmMain.Inet3.GetHeader("Content-Length")
        sz = FileSize / 1000
        frmMain.lblSize.Caption = sz & " KB"
    FileRemaining = FileSize
    FileSize_Current = 0
    
    Debug.Print FileSize
    
    FileNumber = FreeFile
        
    Open App.Path & "/" & SharedGetFileExt For Binary Access Write As #FileNumber
    
    Do Until FileRemaining = 0
        If frmMain.Tag = "Cancel" Then
            frmMain.Inet3.Cancel
            Exit Sub
        End If
        
        If FileRemaining > 1024 Then
            FileData = frmMain.Inet3.GetChunk(1024, icByteArray)
            FileRemaining = FileRemaining - 1024
        Else
            FileData = frmMain.Inet3.GetChunk(FileRemaining, icByteArray)
            FileRemaining = 0
        End If
        
        FileSize_Current = FileSize - FileRemaining
        PBValue = CInt((100 / FileSize) * FileSize_Current)
        frmMain.lblSaved.Caption = FileSize_Current & " bits"
        frmMain.lblRemaining.Caption = FileSize - FileSize_Current & " bits"
        frmMain.lblPercentage.Caption = "% " & PBValue
        frmMain.PB1.Value = PBValue
        
        Put #FileNumber, , FileData
    Loop
    Close #FileNumber
End Sub

Function GetStatus(st As Integer, Inet2 As Inet)
    Select Case st
        Case icError
            GetStatus = Left$(Inet2.ResponseInfo, _
            Len(Inet2.ResponseInfo) - 2)
        Case icResolvingHost, icRequesting, icRequestSent
            GetStatus = "Searching... "
        Case icHostResolved
            GetStatus = "Found." & vName
        Case icReceivingResponse, icResponseReceived
            GetStatus = "Receiving data "
        Case icResponseCompleted
            GetStatus = "Connected"
        Case icConnecting, icConnected
            GetStatus = "Connecting..."
        Case icDisconnecting
            GetStatus = "Disconnecting..."
        Case icDisconnected
            GetStatus = "Disconnected"
        Case Else
    End Select
End Function

