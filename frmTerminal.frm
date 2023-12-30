VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTerminal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WX_APRS"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTerminal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   5880
      Top             =   5400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Login"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send to APRS"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read BME280"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   7335
   End
   Begin VB.TextBox txtLog 
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   270
      Width           =   7500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".........1.........2.........3.........4.........5.........6.........7.........8"
      Height          =   255
      Left            =   8
      TabIndex        =   3
      Top             =   120
      Width           =   7480
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data As String
Dim DataLines() As String
Dim Index As Integer
Dim Temp As Variant
Dim Press As Variant
Dim Priv As String
Dim Faran As Integer
Dim Pressure As Integer
Dim Fso As New FileSystemObject
Dim FileToread As TextStream
Dim FileToWrite As TextStream
Dim CurrentLine As String
Dim FinalFaran As String
Dim FinalPress As String


Private Sub cmdControl_Click()
    frmControl.Show
    Me.Hide
End Sub

Private Sub Command1_Click()


Data = Inet1.OpenURL("http://192.168.1.230:8085")
DataLines = Split(Data, vbCrLf)
For Index = LBound(DataLines) To UBound(DataLines)
  Priv = DataLines(Index)
Next
Temp = Mid(Priv, 190, 55)
Press = Mid(Priv, 277, 7)
Pressure = CDbl(Val((Press)))
'Faran = (CInt(Temp) * 1.8) + 32
Faran = (CDbl(Val((Temp))) * 1.8) + 32
MsgBox Faran
MsgBox Temp
MsgBox Press

        Set FileToWrite = Fso.OpenTextFile(App.Path & "\Log.txt", ForAppending, True)
        FileToWrite.WriteLine (Date & ";" & Time & ";" & CDbl(Val((Temp)))) & ";" & Pressure
        FileToWrite.Close
End Sub

Private Sub Command2_Click()
 If Faran >= 0 Then
        FinalFaran = "t0" & Faran
        Else
        FinalFaran = "t" & Faran
        End If
        If Press >= 10000 Then
        FinalPress = "b" & Pressure
        Else
        FinalPress = "b0" & Pressure
        End If
        frmControl.ttcControl.SendData "9A4AM-2>APRS,TCPIP*:@090247z4528.49N/01645.84E_.../..." & FinalFaran & "Test-WX-station/1.0"
        frmControl.ttcControl.SendData vbCrLf
        Set FileToWrite = Fso.OpenTextFile(App.Path & "Log.txt", ForAppending, True)
        FileToWrite.WriteLine (Date & ";" & Time & ";" & CDbl(Val((Temp))))
        FileToWrite.Close
End Sub

Private Sub Command3_Click()
frmControl.ttcControl.SendData "user 9A4AM-2 pass 13282 vers WX_Station 0.1 filter m/1"
        frmControl.ttcControl.SendData vbCrLf
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Static iCount As Integer
    If iCount = 10 Then
        iCount = 0
        'YOUR CODE HERE
        Data = Inet1.OpenURL("http://192.168.1.230:8085")
DataLines = Split(Data, vbCrLf)
For Index = LBound(DataLines) To UBound(DataLines)
  Priv = DataLines(Index)
Next
Temp = Mid(Priv, 190, 5)
Press = Mid(Priv, 348, 7)
Pressure = CDbl(Val((Press)))
'Faran = (CInt(Temp) * 1.8) + 32
Faran = (CDbl(Val((Temp))) * 1.8) + 32
        If Faran >= 0 Then
        FinalFaran = "t0" & Faran
        Else
        FinalFaran = "t" & Faran
        End If
        If Press >= 10000 Then
        FinalPress = "b" & Pressure
        Else
        FinalPress = "b0" & Pressure
        End If
        frmControl.ttcControl.SendData "9A4AM-2>APRS,TCPIP*:@090247z4528.49N/01645.84E_.../..." & FinalFaran & "Test-WX-station/1.0"
        frmControl.ttcControl.SendData vbCrLf
        Set FileToWrite = Fso.OpenTextFile(App.Path & "Log.txt", ForAppending, True)
        FileToWrite.WriteLine (Date & ";" & Time & ";" & CDbl(Val((Temp))))
        FileToWrite.Close
    Else
        iCount = iCount + 1
    End If

End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
        frmControl.ttcControl.SendData txtInput.Text
        frmControl.ttcControl.SendData vbCrLf
        KeyAscii = 0
        txtLog.Text = txtLog.Text & txtInput.Text & vbCrLf
        txtLog.SelStart = Len(txtLog.Text)
        txtInput.Text = ""
    End If
End Sub
