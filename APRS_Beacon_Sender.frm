VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form APRS_Beacon_Sender 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APRS Beacon Sender by 9A4AM - V1.0"
   ClientHeight    =   9090
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "APRS_Beacon_Sender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   5640
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Interval        =   40000
      Left            =   6360
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   5760
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   5160
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Manual send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8160
      Width           =   7455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5520
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H0080FF80&
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4440
      Width           =   7500
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisconnect 
      BackColor       =   &H000080FF&
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtHost 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   2
      Text            =   "finland.aprs2.net"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   1
      Text            =   "14580"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox chkEcho 
      BackColor       =   &H0080FF80&
      Caption         =   "Server echo"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   5040
      Top             =   1440
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FF80&
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   29
      Top             =   2760
      Width           =   75
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   28
      Top             =   2760
      Width           =   120
   End
   Begin VB.Label lblAutostart 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "Timer next Beacon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblPacket 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   25
      Top             =   2760
      Width           =   75
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "9A4AM@2021"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   8520
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   23
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "Symbol:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblSymbol 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lblComment 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Comment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblSSID 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lblCall 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "SSID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "Call:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblPort 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblServer 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "APRS server response"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   4080
      Width           =   3015
   End
   Begin Project1.TelnetTTYClient ttcControl 
      Left            =   6240
      Top             =   1320
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Host:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Menu Config 
      Caption         =   "Config"
   End
   Begin VB.Menu Service 
      Caption         =   "Service"
   End
   Begin VB.Menu Author_prg 
      Caption         =   "Author"
      Index           =   1
   End
End
Attribute VB_Name = "APRS_Beacon_Sender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fso As New FileSystemObject
Dim Data As String
Dim DataLines() As String
Dim Index As Integer
Dim Temp As Variant
Dim Press As Variant
Dim Priv As String
Dim Faran As Integer
Dim Pressure As Integer
Dim FileToread As TextStream
Dim FileToWrite As TextStream
Dim CurrentLine As String
Dim FinalFaran As String
Dim FinalPress As String
Private Const APPCAPTION As String = "TelnetTTY"
Private Const LFCR = vbLf & vbCr
Dim Call_cfg As String
Dim SSID_cfg As String
Dim Passw_cfg As String
Dim Server_cfg As String
Dim Port_cfg As String
Dim Symbol_cfg As String
Dim Long_cfg As String
Dim Lat_cfg As String
Dim Comment_cfg As String
Dim Time_cfg As Integer
Dim StartUp_cfg As Boolean
Dim Symbol_APRS As String
Dim Packet As Integer
Dim Minus As String
Dim iCount As Integer

Private lngPort As Long

Private Sub DisconnectUI()
    cmdOk.Caption = "Connect"
    cmdDisconnect.Enabled = False
    
End Sub

Private Sub Autor_Click()
frmAutor.Show
End Sub

Private Sub Author_prg_Click(Index As Integer)
frmAutor.Show
End Sub

Private Sub cmdDisconnect_Click()
    ttcControl.Disconnect
    DisconnectUI
End Sub

Private Sub cmdExit_Click()
    ttcControl.Disconnect
    DisconnectUI
    Unload frmConfig
    Unload frmAutor
    Unload frmPassword
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If cmdOk.Caption = "Ok" Then
        ttcControl.Echo chkEcho.Value = vbChecked
       
      
    Else
        ttcControl.TermType = "NVT|TTY"
        lngPort = CLng(Port_cfg)
        ttcControl.Connect Server_cfg, lngPort
    End If
End Sub





Private Sub Command1_Click()
 APRS_Beacon_Sender.ttcControl.SendData Call_cfg & Minus & SSID_cfg & ">APU25N,TCPIP*:@090247z" & Lat_cfg & "/" & Long_cfg & Symbol_APRS & Comment_cfg
       APRS_Beacon_Sender.ttcControl.SendData vbCrLf
      ' MsgBox Call_cfg & Minus & SSID_cfg & ">APU25N,TCPIP*:@090247z" & Lat_cfg & "/" & Long_cfg & Symbol_APRS & Comment_cfg
   Packet = Packet + 1
   lblPacket = Packet
    Timer4.Enabled = True
 
End Sub




Private Sub Config_Click()
Call cmdDisconnect_Click
frmConfig.Show
Me.Hide
End Sub

Private Sub Form_Activate()
Select Case Symbol_cfg
    Case "Antenna"
    Symbol_APRS = "r"
    Case "Ballon"
    Symbol_APRS = "O"
    Case "Home"
    Symbol_APRS = "-"
    Case "WX Station"
    Symbol_APRS = "_"
    Case "Dish antenna"
    Symbol_APRS = "`"
End Select
lblServer = Server_cfg
lblPort = Port_cfg
lblCall = Call_cfg
lblSSID = SSID_cfg
lblComment = Comment_cfg
lblSymbol = Symbol_cfg
lblPacket = iCount

End Sub

Private Sub Form_Load()
On Error GoTo Handler
cmdDisconnect.Enabled = False
cmdOk.Enabled = False
Command1.Enabled = False
    Server_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Server")
    Port_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Port")
    Time_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Time")
    Symbol_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Symbol")
    Comment_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Comment")
    Long_cfg = ReadIniValue(App.Path & "\Config.ini", "Location", "Longitude")
    Lat_cfg = ReadIniValue(App.Path & "\Config.ini", "Location", "Latitude")
    Call_cfg = ReadIniValue(App.Path & "\Config.ini", "Personal_Data", "Call")
    SSID_cfg = ReadIniValue(App.Path & "\Config.ini", "Personal_Data", "SSID")
    Passw_cfg = ReadIniValue(App.Path & "\Config.ini", "Personal_Data", "Password")
    StartUp_cfg = ReadIniValue(App.Path & "\Config.ini", "App", "Start")
    
    lblStatus.BackColor = vbRed
    lblStatus = "Disconnected"
   
    Config.Enabled = False
    If SSID_cfg = "" Then
Minus = ""
Else
Minus = "-"
End If
   If StartUp_cfg = True Then
    
lblAutostart.BackColor = vbGreen
  lblAutostart = "Autostart App with Windows is ENABLED"
    



               
Else

   lblAutostart.BackColor = vbYellow
   lblAutostart = "Autostart App with Windows is DISABLED"
    End If
    Exit Sub
Handler:
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False

MsgBox "Missed APRS data, please enter your data!", vbInformation
APRS_Beacon_Sender.Hide
frmConfig.Show
Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
 ttcControl.Disconnect
 DisconnectUI
    Unload frmConfig
    Unload frmAutor
    Unload frmPassword
    Unload Me
End Sub







Private Sub Service_Click()
frmPassword.Show
End Sub

Private Sub Timer2_Timer()
Call cmdOk_Click
Timer2.Enabled = False


End Sub

Private Sub Timer3_Timer()
Call Command1_Click
Timer3.Enabled = False
Config.Enabled = True

End Sub

Private Sub Timer4_Timer()
       lblStatus.BackColor = vbYellow
    lblStatus = "Disconnected till next Beacon"
    Call cmdDisconnect_Click
    Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
lblStat = Time_cfg
End Sub

Private Sub ttcControl_Connect()
    cmdOk.Caption = "Ok"
    'cmdDisconnect.Enabled = True
    lblStatus.BackColor = vbGreen
    lblStatus = "Connected"
    ttcControl.Echo chkEcho.Value = vbChecked
Sleep 1000
 APRS_Beacon_Sender.ttcControl.SendData "user " & Call_cfg & Minus & SSID_cfg & " pass " & Passw_cfg & " vers APRS_Beacon filter m/1"
APRS_Beacon_Sender.ttcControl.SendData vbCrLf
Sleep 1000
 APRS_Beacon_Sender.ttcControl.SendData "user " & Call_cfg & Minus & SSID_cfg & " pass " & Passw_cfg & " vers APRS_Beacon filter m/1"
APRS_Beacon_Sender.ttcControl.SendData vbCrLf


End Sub

Private Sub ttcControl_DataArrival()
    Dim strdata As String
    
    strdata = Replace$(ttcControl.GetData(), vbCrLf, vbFormFeed)
    strdata = Replace$(strdata, LFCR, vbFormFeed)
    strdata = Replace$(strdata, vbLf, vbFormFeed)
    strdata = Replace$(strdata, vbFormFeed, vbCrLf)
    With txtLog
        If Len(.Text) > 10000 Then
            .Text = Right$(.Text, 10000)
        End If
        .Text = .Text & strdata
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub ttcControl_Disconnect()
    MsgBox "Server connection lost.", vbOKOnly Or vbExclamation, APPCAPTION
    ttcControl.Disconnect
    DisconnectUI
    lblStatus.BackColor = vbRed
    lblStatus = "Disconnected"
End Sub

Private Sub ttcControl_Error(ByVal Number As Long, ByVal Description As String)
    MsgBox "Error &H" & Hex$(Number) & " " & Description, _
           vbOKOnly Or vbExclamation, APPCAPTION
End Sub

Private Sub txtHost_Validate(Cancel As Boolean)
    If Len(Server_cfg) = 0 Then
        MsgBox "Port must be numeric, in the range 1 to 65535.", _
               vbOKOnly Or vbInformation, _
               APPCAPTION
        Cancel = True
    End If
End Sub

Private Sub txtPort_Validate(Cancel As Boolean)
    Cancel = True
    If IsNumeric(txtPort.Text) Then
        lngPort = CLng(txtPort.Text)
        If 0 < lngPort And lngPort < 65536 Then
            Cancel = False
        End If
    End If
    If Cancel Then
        MsgBox "Port must be numeric, in the range 1 to 65535.", _
               vbOKOnly Or vbInformation, _
               APPCAPTION
    End If
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
   
    If iCount = Time_cfg Then
        iCount = 0
        lblPacket = 0
        'YOUR CODE HERE
      
        'APRS_Beacon_Sender.ttcControl.SendData "9A4AM-2>APU25N,TCPIP*:@090247z4528.49N/01645.84Er.../..." & "Test-Home"
        'APRS_Beacon_Sender.ttcControl.SendData vbCrLf
       'APRS_Beacon_Sender.ttcControl.SendData "9A4AM-5>APU25N,TCPIP*:@090247z4528.49N/01645.84Er" & Comment_cfg
       'APRS_Beacon_Sender.ttcControl.SendData vbCrLf
       'APRS_Beacon_Sender.ttcControl.SendData Call_cfg & Minus & SSID_cfg & ">APU25N,TCPIP*:@090247z" & Lat_cfg & "/" & Long_cfg & Symbol_APRS & Comment_cfg
       'APRS_Beacon_Sender.ttcControl.SendData vbCrLf
       'Timer2.Enabled = True
       'Timer3.Enabled = True
       Call Update
       Timer4.Enabled = True
       txtLog.Text = ""
       
       'Packet = Packet + 1
       'lblPacket = Packet
    Else
 
        iCount = iCount + 1
        lblPacket = iCount
       
    End If
End Sub
Public Sub Update()
 Server_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Server")
    Port_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Port")
    Time_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Time")
    Symbol_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Symbol")
    Comment_cfg = ReadIniValue(App.Path & "\Config.ini", "APRS_Data", "Comment")
    Long_cfg = ReadIniValue(App.Path & "\Config.ini", "Location", "Longitude")
    Lat_cfg = ReadIniValue(App.Path & "\Config.ini", "Location", "Latitude")
    Call_cfg = ReadIniValue(App.Path & "\Config.ini", "Personal_Data", "Call")
    SSID_cfg = ReadIniValue(App.Path & "\Config.ini", "Personal_Data", "SSID")
    Passw_cfg = ReadIniValue(App.Path & "\Config.ini", "Personal_Data", "Password")
    StartUp_cfg = ReadIniValue(App.Path & "\Config.ini", "App", "Start")
     txtLog.Text = ""
Select Case Symbol_cfg
    Case "Antenna"
    Symbol_APRS = "r"
    Case "Ballon"
    Symbol_APRS = "O"
    Case "Home"
    Symbol_APRS = "-"
    Case "WX Station"
    Symbol_APRS = "_"
    Case "Dish antenna"
    Symbol_APRS = "`"
End Select
If SSID_cfg = "" Then
Minus = ""
Else
Minus = "-"
End If

lblServer = Server_cfg
lblPort = Port_cfg
lblCall = Call_cfg
lblSSID = SSID_cfg
lblComment = Comment_cfg
lblSymbol = Symbol_cfg
 DisconnectUI
    lblStatus.BackColor = vbRed
    lblStatus = "Disconnected"
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
If StartUp_cfg = True Then
    
 lblAutostart.BackColor = vbGreen
  lblAutostart = "Autostart App with Windows is ENABLED"
    



               
Else
 
   lblAutostart.BackColor = vbYellow
   lblAutostart = "Autostart App with Windows is DISABLED"
    End If
End Sub


