VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demon Soft. - Port Scanner"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "PortScann.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock myTCPclient4 
      Left            =   5280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock myTCPclient3 
      Left            =   4800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock myTCPclient2 
      Left            =   4320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock myTCPclient1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   5535
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Current port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label CurrntPort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   5535
      Begin VB.ComboBox ConectionsEdit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "PortScann.frx":08CA
         Left            =   1440
         List            =   "PortScann.frx":08CC
         TabIndex        =   19
         Text            =   "5"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton ConnectBTN 
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton StopScan 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox RemoteIP_1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "127"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox RemoteIP_2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox RemoteIP_3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox PortStart 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox PortEnd 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "65534"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox RemoteIP_4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TimeOutInterval 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "500"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "T. Out:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Ports:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Conectons:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   120
   End
   Begin VB.TextBox Status 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3480
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock myTCPclient 
      Left            =   3360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line4 
      X1              =   2880
      X2              =   2880
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Demon Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port1 As Long
Dim Port2 As Long
Dim IPnum As String
Dim ScanBroi As Long
Dim ScannStr As Boolean
Dim AboutStr As String
Dim Temp As Long
Dim ScanSess As Byte
Dim Stat As Byte
Dim ShowPort As Long
Dim ConnectNow As Byte
Dim ProgBValue As Long
Dim ProgBMax As Long

Private Sub Command1_Click()
MsgBox (AboutStr)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Hand
Temp = 0

Rem About string
AboutStr = "Created by " & Chr(10) & Chr(13) _
& "Ivan Blagoev and Filip Bogdanov Andonov" _
& Chr(10) & Chr(13) & "'Demon Software'."

Temp = 0
ScanBroi = 0
ScannStr = False

Rem ComboBox Values
Me.ConectionsEdit.AddItem ("1")
Me.ConectionsEdit.AddItem ("2")
Me.ConectionsEdit.AddItem ("3")
Me.ConectionsEdit.AddItem ("4")
Me.ConectionsEdit.AddItem ("5")
Exit Sub

Err_Hand:
MsgBox Err.Description
Unload Me
End Sub


Private Sub StopScan_Click()
Rem Cancel Scanning
If ScannStr = True Then
Me.TimeOut.Enabled = False
myTCPclient.Close
MsgBox ("Abort scanning.")
Me.ConnectBTN.Enabled = True
Me.RemoteIP_1.Enabled = True
Me.RemoteIP_2.Enabled = True
Me.RemoteIP_3.Enabled = True
Me.RemoteIP_4.Enabled = True
Me.TimeOutInterval.Enabled = True
Me.PortStart.Enabled = True
Me.PortEnd.Enabled = True
Me.ConectionsEdit.Enabled = True
ScannStr = False
Me.ProgressBar1.Visible = False
End If
End Sub

Rem Scann Button
Private Sub ConnectBTN_Click()
On Error GoTo Err_Hand

Rem Verify values IP and Ports
On Error GoTo Err_IP
If Me.RemoteIP_1.Text < 0 Or Me.RemoteIP_1.Text > 255 Or Me.RemoteIP_1.Text = "" Then GoTo Err_IP
If Me.RemoteIP_2.Text < 0 Or Me.RemoteIP_2.Text > 255 Or Me.RemoteIP_2.Text = "" Then GoTo Err_IP
If Me.RemoteIP_3.Text < 0 Or Me.RemoteIP_3.Text > 255 Or Me.RemoteIP_3.Text = "" Then GoTo Err_IP
If Me.RemoteIP_4.Text < 0 Or Me.RemoteIP_4.Text > 255 Or Me.RemoteIP_4.Text = "" Then GoTo Err_IP

On Error GoTo Err_Hand
Rem Verfy Port Values
If Val(Me.PortStart.Text) < 1 Or Val(Me.PortStart.Text) > 65534 Then GoTo No_Port
If Val(Me.PortEnd.Text) < 1 Or Val(Me.PortEnd.Text) > 65534 Then GoTo No_Port
If Me.PortStart.Text > Me.PortEnd.Text Then GoTo No_Port

Rem Time Out Value
If Val(Me.TimeOutInterval.Text) < 10 Or Val(Me.TimeOutInterval.Text) > 10000 Or Me.TimeOutInterval.Text = "" Then
MsgBox ("Time out interval is: 10 to 10000")
Exit Sub
End If

Rem Verify ComboBox Value
If Val(Me.ConectionsEdit.Text) < 1 Or Val(Me.ConectionsEdit.Text) > 5 Then
MsgBox ("No valid caonections entry.")
Me.ConectionsEdit.Text = 5
Exit Sub
End If

Rem Value entry
Port1 = Me.PortStart.Text
Port2 = Me.PortEnd.Text
IPnum = Me.RemoteIP_1.Text & "." & Me.RemoteIP_2.Text & "." _
& Me.RemoteIP_3.Text & "." & Me.RemoteIP_4.Text
ScanSess = Me.ConectionsEdit.Text

Rem Change Menu and run scanning
Me.Status.Text = ""
Me.CurrntPort.Caption = ""
ScanBroi = Port1
Me.StopScan.SetFocus
Me.ConnectBTN.Enabled = False
Me.RemoteIP_1.Enabled = False
Me.RemoteIP_2.Enabled = False
Me.RemoteIP_3.Enabled = False
Me.RemoteIP_4.Enabled = False
Me.TimeOutInterval.Enabled = False
Me.PortStart.Enabled = False
Me.PortEnd.Enabled = False
Me.ConectionsEdit.Enabled = False
ScannStr = True
Me.TimeOut.Interval = Me.TimeOutInterval.Text

Me.ProgressBar1.Value = 0
ProgBMax = Port2 - Port1
ProgBValue = 0
Me.ProgressBar1.Max = ProgBMax
Me.ProgressBar1.Visible = True

Call SannProc

Exit Sub

Rem Errors in procedure
No_Port:
MsgBox ("No valid port entry.")
Exit Sub

Err_IP:
MsgBox ("No valid IP address.")
Exit Sub

Err_Hand:
MsgBox Err.Description & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rem Çàòâàðÿíå íà âðúçêàòà êúì ñúðâúðà
Me.myTCPclient.Close
End Sub

Private Sub ShowStatus()
On Error GoTo Err_Hand
Dim Str2 As String

Select Case Stat
Case 0
            Exit Sub
Case 1
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " - Open" & vbCrLf
Case 2
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " - Listening" & vbCrLf
Case 3
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " - Connection pending" & vbCrLf
Case 4
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " -  Resolving host" & vbCrLf
Case 5
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " -  Host resolved" & vbCrLf
Case 6
            Exit Sub
Case 7
            Rem Íàìèðàíå íà òî÷íèÿò Socket
            If ConnectNow = 5 Then
                Me.myTCPclient4.GetData Str2, , 150
            ElseIf ConnectNow = 4 Then
                Me.myTCPclient3.GetData Str2, , 150
            ElseIf ConnectNow = 3 Then
                Me.myTCPclient2.GetData Str2, , 150
            ElseIf ConnectNow = 2 Then
                Me.myTCPclient1.GetData Str2, , 150
            Else
                Me.myTCPclient.GetData Str2, , 150
            End If
            Rem Àêî ñòðèíãúò å ïðàçåí
            If Str2 = "" Then
                Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " -  Connected" & vbCrLf
                Exit Sub
            End If
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " -  Connected: " & Str2 & vbCrLf
Case 8
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " -  Peer is closing the connection" & vbCrLf
Case 9
            Me.Status.Text = Me.Status.Text & vbCrLf & ShowPort & " -  Error" & vbCrLf
End Select
Str2 = ""
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub Err_IP()
Rem IP message
On Error GoTo Err_Hand
MsgBox ("No valid IP: 0 - 255")
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub TimeOut_Timer()
On Error GoTo Err_Hand
Me.TimeOut.Enabled = False

Rem Get Connection status
Rem If connections > 1
Select Case ScanSess
Case 1
        ConnectNow = 1
        Stat = myTCPclient.State
        ShowPort = myTCPclient.RemotePort
        Call ShowStatus
Case 2
        ConnectNow = 1
        Stat = myTCPclient.State
        ShowPort = myTCPclient.RemotePort
        Call ShowStatus
        
        ConnectNow = 2
        Stat = myTCPclient1.State
        ShowPort = myTCPclient1.RemotePort
        Call ShowStatus
Case 3
        ConnectNow = 1
        Stat = myTCPclient.State
        ShowPort = myTCPclient.RemotePort
        Call ShowStatus
        
        ConnectNow = 2
        Stat = myTCPclient1.State
        ShowPort = myTCPclient1.RemotePort
        Call ShowStatus
        
        ConnectNow = 3
        Stat = myTCPclient2.State
        ShowPort = myTCPclient2.RemotePort
        Call ShowStatus
Case 4
        ConnectNow = 1
        Stat = myTCPclient.State
        ShowPort = myTCPclient.RemotePort
        Call ShowStatus
        
        ConnectNow = 2
        Stat = myTCPclient1.State
        ShowPort = myTCPclient1.RemotePort
        Call ShowStatus
        
        ConnectNow = 3
        Stat = myTCPclient2.State
        ShowPort = myTCPclient2.RemotePort
        Call ShowStatus
        
        ConnectNow = 4
        Stat = myTCPclient3.State
        ShowPort = myTCPclient3.RemotePort
        Call ShowStatus
Case 5
        ConnectNow = 1
        Stat = myTCPclient.State
        ShowPort = myTCPclient.RemotePort
        Call ShowStatus
        
        ConnectNow = 2
        Stat = myTCPclient1.State
        ShowPort = myTCPclient1.RemotePort
        Call ShowStatus
        
        ConnectNow = 3
        Stat = myTCPclient2.State
        ShowPort = myTCPclient2.RemotePort
        Call ShowStatus
        
        ConnectNow = 4
        Stat = myTCPclient3.State
        ShowPort = myTCPclient3.RemotePort
        Call ShowStatus
        
        ConnectNow = 5
        Stat = myTCPclient4.State
        ShowPort = myTCPclient4.RemotePort
        Call ShowStatus
End Select

Rem ÏðîãðåñÁàð
ProgBValue = ProgBValue + ScanSess
If ProgBMax < ProgBValue Then
Me.ProgressBar1.Value = ProgBMax
GoTo No_ProgB
End If
Me.ProgressBar1.Value = ProgBValue
No_ProgB:

Rem Netxt port
ScanBroi = ScanBroi + ScanSess

Rem If all ports scannet then EXIT
If ScanBroi > Port2 Then
EndScan:
Me.TimeOut.Enabled = False
myTCPclient.Close
MsgBox ("End scanning.")
Me.ConnectBTN.Enabled = True
Me.RemoteIP_1.Enabled = True
Me.RemoteIP_2.Enabled = True
Me.RemoteIP_3.Enabled = True
Me.RemoteIP_4.Enabled = True
Me.TimeOutInterval.Enabled = True
Me.PortStart.Enabled = True
Me.PortEnd.Enabled = True
Me.ConectionsEdit.Enabled = True
Me.ProgressBar1.Visible = False
ScannStr = False
Exit Sub
End If

Call SannProc
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub SannProc()
On Error GoTo Err_Hand

Rem View current port
Me.CurrntPort.Caption = ScanBroi

Rem If connections > 1
Select Case ScanSess
Case 1
        Rem Close Connections
        If (myTCPclient.State <> sckClosed) Then myTCPclient.Close

        Rem Connect IP and Port
        Me.myTCPclient.RemoteHost = IPnum
        Me.myTCPclient.RemotePort = ScanBroi
        Me.myTCPclient.Connect
        Stat = myTCPclient.State
Case 2
        Rem Close Connections
        If (myTCPclient.State <> sckClosed) Then myTCPclient.Close
        If (myTCPclient1.State <> sckClosed) Then myTCPclient1.Close

        Rem Connect IP and Port
        Me.myTCPclient.RemoteHost = IPnum
        Me.myTCPclient.RemotePort = ScanBroi
        Me.myTCPclient.Connect
        
        Me.myTCPclient1.RemoteHost = IPnum
        Me.myTCPclient1.RemotePort = ScanBroi + 1
        Me.myTCPclient1.Connect
Case 3
        Rem Close Connections
        If (myTCPclient.State <> sckClosed) Then myTCPclient.Close
        If (myTCPclient1.State <> sckClosed) Then myTCPclient1.Close
        If (myTCPclient2.State <> sckClosed) Then myTCPclient2.Close

        Rem Connect IP and Port
        Me.myTCPclient.RemoteHost = IPnum
        Me.myTCPclient.RemotePort = ScanBroi
        Me.myTCPclient.Connect
        
        Me.myTCPclient1.RemoteHost = IPnum
        Me.myTCPclient1.RemotePort = ScanBroi + 1
        Me.myTCPclient1.Connect
        
        Me.myTCPclient2.RemoteHost = IPnum
        Me.myTCPclient2.RemotePort = ScanBroi + 2
        Me.myTCPclient2.Connect
Case 4
        Rem Close Connections
        If (myTCPclient.State <> sckClosed) Then myTCPclient.Close
        If (myTCPclient1.State <> sckClosed) Then myTCPclient1.Close
        If (myTCPclient2.State <> sckClosed) Then myTCPclient2.Close
        If (myTCPclient3.State <> sckClosed) Then myTCPclient3.Close

        Rem Connect IP and Port
        Me.myTCPclient.RemoteHost = IPnum
        Me.myTCPclient.RemotePort = ScanBroi
        Me.myTCPclient.Connect
        
        Me.myTCPclient1.RemoteHost = IPnum
        Me.myTCPclient1.RemotePort = ScanBroi + 1
        Me.myTCPclient1.Connect
        
        Me.myTCPclient2.RemoteHost = IPnum
        Me.myTCPclient2.RemotePort = ScanBroi + 2
        Me.myTCPclient2.Connect
        
        Me.myTCPclient3.RemoteHost = IPnum
        Me.myTCPclient3.RemotePort = ScanBroi + 3
        Me.myTCPclient3.Connect
Case 5
        Rem Close Connections
        If (myTCPclient.State <> sckClosed) Then myTCPclient.Close
        If (myTCPclient1.State <> sckClosed) Then myTCPclient1.Close
        If (myTCPclient2.State <> sckClosed) Then myTCPclient2.Close
        If (myTCPclient3.State <> sckClosed) Then myTCPclient3.Close
        If (myTCPclient4.State <> sckClosed) Then myTCPclient4.Close

        Rem Connect IP and Port
        Me.myTCPclient.RemoteHost = IPnum
        Me.myTCPclient.RemotePort = ScanBroi
        Me.myTCPclient.Connect
        
        Me.myTCPclient1.RemoteHost = IPnum
        Me.myTCPclient1.RemotePort = ScanBroi + 1
        Me.myTCPclient1.Connect
        
        Me.myTCPclient2.RemoteHost = IPnum
        Me.myTCPclient2.RemotePort = ScanBroi + 2
        Me.myTCPclient2.Connect
        
        Me.myTCPclient3.RemoteHost = IPnum
        Me.myTCPclient3.RemotePort = ScanBroi + 3
        Me.myTCPclient3.Connect
        
        Me.myTCPclient4.RemoteHost = IPnum
        Me.myTCPclient4.RemotePort = ScanBroi + 4
        Me.myTCPclient4.Connect
End Select


Rem Enable Timer
Me.TimeOut.Enabled = True
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub
