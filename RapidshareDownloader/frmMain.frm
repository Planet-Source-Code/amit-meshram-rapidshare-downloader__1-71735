VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rapidshare Downloader..."
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   8355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Opt60 
      Caption         =   "Check Download in 60 sec"
      Height          =   315
      Left            =   5760
      TabIndex        =   15
      Top             =   1740
      Width           =   2265
   End
   Begin VB.OptionButton Opt30 
      Caption         =   "Check Download in 30 sec"
      Height          =   315
      Left            =   3420
      TabIndex        =   14
      Top             =   1740
      Value           =   -1  'True
      Width           =   2235
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1410
      Top             =   540
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   6780
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7380
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel Download"
      Height          =   405
      Left            =   4830
      TabIndex        =   12
      Top             =   1260
      Width           =   1515
   End
   Begin VB.CommandButton CmdDownload 
      Caption         =   "Start Download"
      Height          =   405
      Left            =   3390
      TabIndex        =   11
      Top             =   1260
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   540
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StPanel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   2145
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9551
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:30 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "http://rapidshare.com/files/180564572/7zip_www.sxforum.org.rar"
      Top             =   120
      Width           =   6105
   End
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   7950
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblWait 
      Caption         =   "Wait : "
      Height          =   285
      Left            =   150
      TabIndex        =   13
      Top             =   1740
      Width           =   3135
   End
   Begin VB.Label lblPercentage 
      Caption         =   "Persent % Completed..."
      Height          =   285
      Left            =   3390
      TabIndex        =   10
      Top             =   900
      Width           =   2925
   End
   Begin VB.Label lblRemaining 
      Caption         =   "In Bits"
      Height          =   285
      Left            =   2190
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "File Remaining"
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblSaved 
      Caption         =   "In Bits"
      Height          =   285
      Left            =   2190
      TabIndex        =   6
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "File Saved"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label lblSize 
      Caption         =   "In Bits"
      Height          =   285
      Left            =   2190
      TabIndex        =   3
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "File Size"
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Rapidshare URL"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnt30s As Integer

Private Sub CmdCancel_Click()
    If frmMain.Tag = "Cancel" Then
        Inet1.Cancel
        Inet2.Cancel
        Inet3.Cancel
    End If
    Timer1.Enabled = False
End Sub

Private Sub CmdDownload_Click()
    Call GetInfo1(Inet1, Trim(txtURL.Text))
    Timer1.Enabled = True
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Inet3_StateChanged(ByVal State As Integer)
    StPanel.Panels(1).Text = GetStatus(State, Inet3)
End Sub

Private Sub Timer1_Timer()
    If Opt30.Value = True Then
        Cnt30s = Cnt30s + 1
        lblWait.Caption = "Please Wait for : " & Cnt30s

        If Cnt30s > 30 Then
            Call DownloadCreate
            Timer1.Enabled = False
        End If
    End If
    
    If Opt60.Value = True Then
        Cnt30s = Cnt30s + 1
        lblWait.Caption = "Please Wait for : " & Cnt30s
        
        If Cnt30s > 60 Then
            Call DownloadCreate
            Timer1.Enabled = False
        End If
    End If
End Sub
