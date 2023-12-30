VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00FF8080&
   Caption         =   "Config APRS data"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStartUp 
      BackColor       =   &H00FF8080&
      Caption         =   "Start when PC startup"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   4320
      Width           =   3375
   End
   Begin VB.ComboBox cboSymbol 
      Height          =   420
      ItemData        =   "Config.frx":0000
      Left            =   6360
      List            =   "Config.frx":0013
      TabIndex        =   21
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtComment 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox txtLat 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtLong 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H0080FF80&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Save 
      BackColor       =   &H008080FF&
      Caption         =   "SAVE DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtAPRS_time 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtSSID 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtCall 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cboServer 
      Height          =   420
      ItemData        =   "Config.frx":0048
      Left            =   6360
      List            =   "Config.frx":005B
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      Caption         =   "Use this format"
      Height          =   375
      Left            =   6720
      TabIndex        =   26
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      Caption         =   "01645.83E"
      Height          =   375
      Left            =   7080
      TabIndex        =   25
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF8080&
      Caption         =   "4528.49N"
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF8080&
      Caption         =   "Symbol:"
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "APRS comment:"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "Latitude:"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "Longitude:"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "min"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Time:"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Port:"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Server:"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Password:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "SSID:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Call:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub Cancel_Click()
APRS_Beacon_Sender.Update
frmConfig.Hide
APRS_Beacon_Sender.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
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
    
    cboServer.Text = Server_cfg
    txtPort.Text = Port_cfg
    txtAPRS_time.Text = Time_cfg
    cboSymbol.Text = Symbol_cfg
    txtComment.Text = Comment_cfg
    txtLong.Text = Long_cfg
    txtLat.Text = Lat_cfg
    txtCall.Text = Call_cfg
    txtSSID.Text = SSID_cfg
    txtPassword.Text = Passw_cfg
    chkStartUp.Value = Val(StartUp_cfg)
    chkStartUp.Enabled = False    'for now -- not work
    

    
    
    
End Sub

Private Sub Save_Click()
WriteIniValue App.Path & "\Config.ini", "APRS_Data", "Server", cboServer.Text
WriteIniValue App.Path & "\Config.ini", "APRS_Data", "Port", txtPort.Text
WriteIniValue App.Path & "\Config.ini", "APRS_Data", "Time", txtAPRS_time.Text
WriteIniValue App.Path & "\Config.ini", "APRS_Data", "Symbol", cboSymbol.Text
WriteIniValue App.Path & "\Config.ini", "APRS_Data", "Comment", txtComment.Text
WriteIniValue App.Path & "\Config.ini", "Location", "Longitude", txtLong.Text
WriteIniValue App.Path & "\Config.ini", "Location", "Latitude", txtLat.Text
WriteIniValue App.Path & "\Config.ini", "Personal_Data", "Call", txtCall.Text
WriteIniValue App.Path & "\Config.ini", "Personal_Data", "SSID", txtSSID.Text
WriteIniValue App.Path & "\Config.ini", "Personal_Data", "Password", txtPassword.Text
WriteIniValue App.Path & "\Config.ini", "App", "Start", chkStartUp.Value
APRS_Beacon_Sender.Update
frmConfig.Hide
APRS_Beacon_Sender.Show

End Sub



Private Sub txtCall_LostFocus()
 txtCall.Text = UCase(txtCall.Text)
End Sub
