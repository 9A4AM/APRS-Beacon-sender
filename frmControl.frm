VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WX_APRS"
   ClientHeight    =   1800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   225
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox chkEcho 
      Caption         =   "Server echo"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   600
      TabIndex        =   1
      Text            =   "14580"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtHost 
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Text            =   "finland.aprs2.net"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin TelnetTTY.TelnetTTYClient ttcControl 
      Left            =   3000
      Top             =   1200
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const APPCAPTION As String = "TelnetTTY"
Private Const LFCR = vbLf & vbCr

Private lngPort As Long

Private Sub DisconnectUI()
    cmdOk.Caption = "Connect"
    cmdDisconnect.Enabled = False
    frmTerminal.Hide
    frmTerminal.txtLog.Text = ""
    Me.Show
End Sub

Private Sub cmdDisconnect_Click()
    ttcControl.Disconnect
    DisconnectUI
End Sub

Private Sub cmdExit_Click()
    ttcControl.Disconnect
    Unload frmTerminal
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If cmdOk.Caption = "Ok" Then
        ttcControl.Echo chkEcho.Value = vbChecked
        Me.Hide
        frmTerminal.Show vbModeless, Me
        frmTerminal.txtInput.SetFocus
    Else
        ttcControl.TermType = "NVT|TTY"
        lngPort = CLng(txtPort.Text)
        ttcControl.Connect txtHost.Text, lngPort
    End If
End Sub



Private Sub Command1_Click()
frmTerminal.Show vbModeless, Me
    frmTerminal.txtInput.SetFocus
End Sub

Private Sub ttcControl_Connect()
    cmdOk.Caption = "Ok"
    cmdDisconnect.Enabled = True
    ttcControl.Echo chkEcho.Value = vbChecked
    Me.Hide
    frmTerminal.Show vbModeless, Me
    frmTerminal.txtInput.SetFocus
End Sub

Private Sub ttcControl_DataArrival()
    Dim strData As String
    
    strData = Replace$(ttcControl.GetData(), vbCrLf, vbFormFeed)
    strData = Replace$(strData, LFCR, vbFormFeed)
    strData = Replace$(strData, vbLf, vbFormFeed)
    strData = Replace$(strData, vbFormFeed, vbCrLf)
    With frmTerminal.txtLog
        If Len(.Text) > 10000 Then
            .Text = Right$(.Text, 10000)
        End If
        .Text = .Text & strData
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub ttcControl_Disconnect()
    MsgBox "Server connection lost.", vbOKOnly Or vbExclamation, APPCAPTION
    ttcControl.Disconnect
    DisconnectUI
End Sub

Private Sub ttcControl_Error(ByVal Number As Long, ByVal Description As String)
    MsgBox "Error &H" & Hex$(Number) & " " & Description, _
           vbOKOnly Or vbExclamation, APPCAPTION
End Sub

Private Sub txtHost_Validate(Cancel As Boolean)
    If Len(txtHost.Text) = 0 Then
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
