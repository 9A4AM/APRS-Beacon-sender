VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl TelnetTTYClient 
   CanGetFocus     =   0   'False
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "TelnetTTYClient.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   435
   ToolboxBitmap   =   "TelnetTTYClient.ctx":0B82
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock wscTelnet 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "TelnetTTYClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------
'TelnetTTYClient 1.1
'-------------------
'
'A private UserControl that wraps the Microsoft Winsock
'control and provides some limited Telnet Protocol
'functionality on top of it.
'
'
'COPYRIGHT
'
'Copyright © 2005, 2007, 2011, 2012 Robert D. Riemersma, Jr.
'All Rights Reserved.
'
'Permission granted for unlimited use.  Derivative works
'are encouraged.  No support provided, no guarantees offered,
'no liabilities accepted.  Offered as-is, use at your own risk.
'
'
'PROPERTIES
'
'  LocalPort    Long (R/W)
'
'               Similar to Winsock control.
'
'  RemoteHost   String (R/W)
'
'               Similar to Winsock control.
'
'  RemoteHostIP String (RO)
'
'               Similar to Winsock control.
'
'  RemotePort   Long (R/W)
'
'               Similar to Winsock control.
'
'  TermType     String (R/W)
'
'               One or more Telnet terminal type names,
'               separated by | characters.
'
'               Examples:
'
'                 TTY
'                 VT52|TTY|UNKNOWN
'
'               If not specified (prior to Connect() call),
'               TelnetTTYControl will use the Telnet terminal
'               type UNKNOWN per Telnet RFC 1091.
'
'               Once a conection has been established and
'               terminal type has been negotiated, returns the
'               negotiated terminal type.
'
'
'METHODS
'
'  Connect      (Optional ByVal RemoteHost As String,
'                Optional ByVal RemotePort As Long)
'
'               Similar to Winsock control Connect().
'
'  Disconnect   ()
'
'               Similar to Winsock control Close().
'
'  Echo         (ByVal Yes As Boolean)
'
'               Initiates Telnet option dialog to request
'               remote-echo on (Yes=True) or off.  Call after
'               the Connect event has been raised.
'
'               NOTE: Windows Telnet Service will refuse
'               no-echo requests.
'
'  GetData      () As String
'
'               Similar to Winsock control GetData() but
'               fetches received data from the server *without*
'               embedded Telnet Protocol commands which are
'               extracted from the stream.
'
'               Only returns whole String containing the
'               entire contents of the buffer.
'
'  SendData     (ByVal Data As String)
'
'               Similar to Winsock control SendData() but only
'               accepts a String to output.
'
'
'EVENTS
'
'  Connect      ()
'
'               Similar to Winsock control Connect()
'
'  DataArrival  ()
'
'               Similar to Winsock control DataArrival() but
'               without any "bytesTotal" parameter.
'
'  Disconnect   ()
'
'               Similar to Winsock control Close()
'
'  Error        (ByVal Number As Long, ByVal Description As String)
'
'               Similar to Winsock control Error() but a simpler
'               parameter list.  Can return Telnet-related
'               errors *and* passes through Winsock control
'               errors and some socket errors that occur as well.
'
'
'ENUMS
'
'  TTC_ERRORS   TelnetTTYControl-specific error numbers and a
'               socket error value.  See below for details.

Option Explicit

Private Const SOCKET_ERROR = -1
Private Const SOL_SOCKET = 65535 'Options for socket level.
Private Const SO_OOBINLINE = &H100& 'Leave received OOB data in line.

Private Declare Function setsockopt Lib "ws2_32" ( _
    ByVal s As Long, _
    ByVal level As Long, _
    ByVal optname As Long, _
    ByRef optval As Any, _
    ByVal optlen As Long) As Long

Public Enum TTC_ERRORS
    TTC_HOST_WILLECHO = &H80040401
    TTC_HOST_WONTECHO = &H80040402
    TTC_REFUSED_DOECHO = &H80040411
    TTC_REFUSED_DONTECHO = &H80040412
    TTC_SOCKET_ERROR = &H80040421
End Enum

Public Event Connect()

Public Event DataArrival()

Public Event Disconnect()

Public Event Error(ByVal Number As Long, ByVal Description As String)

Private telIAC As String
Private telSB As String
Private telSE As String

Private telDO As String
Private telDONT As String
Private telWILL As String
Private telWONT As String

Private telECHO As String
Private telDO_ECHO As String
Private telDONT_ECHO As String
Private telWILL_ECHO As String
Private telWONT_ECHO As String

Private telTERMTYPE As String
Private telTTIS As String
Private telDO_TERMTYPE As String
Private telWILL_TERMTYPE As String
Private telWONT_TERMTYPE As String
Private telSB_TERMTYPE As String

Private Enum ClientInitiatedCommandStateEnum
    cicIdle
    cicDoEchoBegin
    cicDontEchoBegin
End Enum

Private ClientInitiatedCommandState As ClientInitiatedCommandStateEnum
Private colSendCommandQ As New Collection
Private strRawBuf As String 'Raw data received from Winsock control.
Private strRcvData As String 'Data (with Telnet commands filtered out).
Private varTTypes As Variant 'Would be a dynamic String array, but IsEmpty() is useful.
Private intTType As Integer 'Index into array in varTTypes.
Private strTTypeNegotiated As String 'Negotiated TType after Connect().

Private Sub UserControl_Initialize()
    telIAC = Chr$(255)
    
    telDO = Chr$(253)
    telDONT = Chr$(254)
    telWILL = Chr$(251)
    telWONT = Chr$(252)
    telSB = Chr$(250)
    telSE = Chr$(240)
    
    telECHO = Chr$(1)
    telDO_ECHO = telIAC & telDO & telECHO
    telDONT_ECHO = telIAC & telDONT & telECHO
    telWILL_ECHO = telIAC & telWILL & telECHO
    telWONT_ECHO = telIAC & telWONT & telECHO
    
    telTERMTYPE = Chr$(24)
    telTTIS = Chr$(0)
    telDO_TERMTYPE = telIAC & telDO & telTERMTYPE
    telWILL_TERMTYPE = telIAC & telWILL & telTERMTYPE
    telWONT_TERMTYPE = telIAC & telWONT & telTERMTYPE
    telSB_TERMTYPE = telIAC & telSB & telTERMTYPE
    
    wscTelnet.Protocol = sckTCPProtocol
    wscTelnet.LocalPort = 0
    Me.LocalPort = 0
    
    strRcvData = ""
    ClientInitiatedCommandState = cicIdle
End Sub

Private Sub SendCommand()
    Dim strCommand As String
    
    If ClientInitiatedCommandState = cicIdle Then
        'Not awaiting reply, send next one queued.
        If colSendCommandQ.Count > 0 Then
            strCommand = colSendCommandQ.Item(1)
            colSendCommandQ.Remove 1
            wscTelnet.SendData strCommand
            Select Case Left$(strCommand, 3)
                Case telDO_ECHO
                    ClientInitiatedCommandState = cicDoEchoBegin
                Case telDONT_ECHO
                    ClientInitiatedCommandState = cicDontEchoBegin
            End Select
        End If
    End If
End Sub

Private Sub QueueCommand(ByVal Command As String)
    colSendCommandQ.Add Command
    SendCommand
End Sub

Private Sub SendTermType(ByVal TType As String)
    wscTelnet.SendData telSB_TERMTYPE
    wscTelnet.SendData telTTIS
    wscTelnet.SendData TType
    wscTelnet.SendData telIAC
    wscTelnet.SendData telSE
    strTTypeNegotiated = TType
End Sub

Private Function ProcessReceivedCommand(ByVal intCommandStart As Integer) As Integer
    'Returns length of command processed.
    Dim strCommand As String
    Dim strBounce As String
    
    ProcessReceivedCommand = 3 'Most commands are 3 bytes.
    strCommand = Mid$(strRawBuf, intCommandStart, 3)
    Select Case strCommand
        Case telWILL_ECHO
            Select Case ClientInitiatedCommandState
                Case cicDoEchoBegin
                    'Good response.
                Case cicDontEchoBegin
                    RaiseEvent Error(TTC_REFUSED_DONTECHO, _
                                     "DONT ECHO refused by host")
                Case Else
                    RaiseEvent Error(TTC_HOST_WILLECHO, _
                                     "Host WILL ECHO")
            End Select
            ClientInitiatedCommandState = cicIdle
        Case telWONT_ECHO
            Select Case ClientInitiatedCommandState
                Case cicDontEchoBegin
                    'Good response.
                Case cicDoEchoBegin
                    RaiseEvent Error(TTC_REFUSED_DOECHO, _
                                     "DO ECHO refused by host")
                Case Else
                    RaiseEvent Error(TTC_HOST_WONTECHO, _
                                     "Host WONT ECHO")
            End Select
            ClientInitiatedCommandState = cicIdle
        Case telDO_TERMTYPE
            If Not IsEmpty(varTTypes) Then
                intTType = -1
                wscTelnet.SendData telWILL_TERMTYPE
            Else
                wscTelnet.SendData telWONT_TERMTYPE
            End If
        Case telSB_TERMTYPE 'SEND request.
            If Len(strRawBuf) < intCommandStart + 5 Then
                ProcessReceivedCommand = 0 'Long command, it is incomplete.
            Else
                ProcessReceivedCommand = 6 'Long command.
                If IsEmpty(varTTypes) Then
                    SendTermType "UNKNOWN"
                Else
                    intTType = intTType + 1
                    If intTType > UBound(varTTypes) Then
                        'At end-of-type-list, send last type,
                        'reset list to start.
                        SendTermType varTTypes(intTType - 1)
                        intTType = -1
                    Else
                        SendTermType varTTypes(intTType)
                    End If
                End If
            End If
        Case Else 'Any other command.
            Select Case Mid$(strCommand, 2, 1)
                Case telDO, telDONT
                    strBounce = telWONT
                Case telWILL, telWONT
                    strBounce = telDONT
            End Select
            Mid$(strCommand, 2, 1) = strBounce
            wscTelnet.SendData strCommand
    End Select
End Function

Public Property Get LocalPort() As Long
    LocalPort = wscTelnet.LocalPort
End Property

Public Property Let LocalPort(ByVal Port As Long)
    wscTelnet.LocalPort = Port
End Property

Public Property Get RemoteHost() As String
    RemoteHost = wscTelnet.RemoteHost
End Property

Public Property Let RemoteHost(ByVal Host As String)
    wscTelnet.RemoteHost = Host
End Property

Public Property Get RemoteHostIP() As String
    RemoteHostIP = wscTelnet.RemoteHostIP
End Property

Public Property Get RemotePort() As Long
    RemotePort = wscTelnet.RemotePort
End Property

Public Property Let RemotePort(ByVal Port As Long)
    wscTelnet.RemotePort = Port
End Property

Public Property Get TermType() As String
    If Len(strTTypeNegotiated) > 0 Then
        TermType = strTTypeNegotiated
    Else
        If IsEmpty(varTTypes) Then
            TermType = ""
        Else
            TermType = Join$(varTTypes, "|")
        End If
    End If
End Property

Public Property Let TermType(ByVal Types As String)
    If Len(Types) = 0 Then
        varTTypes = Empty
    Else
        varTTypes = Split(Types, "|")
    End If
End Property

Public Sub Connect(Optional ByVal RemoteHost As String, _
                   Optional ByVal RemotePort As Long)
    On Error Resume Next
    With wscTelnet
        .Bind 0
        If setsockopt(.SocketHandle, SOL_SOCKET, SO_OOBINLINE, 1, 4) = SOCKET_ERROR Then
            RaiseEvent Error(TTC_SOCKET_ERROR, "Error setting SO_OOBINLINE socket option")
            Exit Sub
        Else
            .Close
            .LocalPort = 0
            .Connect RemoteHost, RemotePort
        End If
    End With
End Sub

Public Sub Disconnect()
    Do While colSendCommandQ.Count > 0
        colSendCommandQ.Remove 1
    Loop
    wscTelnet.Close
    wscTelnet.LocalPort = 0
    strTTypeNegotiated = ""
End Sub

Public Sub Echo(ByVal Yes As Boolean)
    If Yes Then
        QueueCommand telDO_ECHO
    Else
        QueueCommand telDONT_ECHO
    End If
End Sub

Public Function GetData() As String
    GetData = strRcvData
    strRcvData = ""
End Function

Public Sub SendData(ByVal Data As String)
On Error Resume Next
    wscTelnet.SendData Data
   
End Sub

Private Sub wscTelnet_Close()
    strTTypeNegotiated = ""
    RaiseEvent Disconnect
End Sub

Private Sub wscTelnet_Connect()
    RaiseEvent Connect
End Sub

Private Sub wscTelnet_DataArrival(ByVal bytesTotal As Long)
    Dim strBuf As String
    Dim intRawChar As Integer
    Dim intCommandStart As Integer
    Dim intCommandLen As Integer
    
    wscTelnet.GetData strBuf
    strRawBuf = strRawBuf & strBuf
    intRawChar = 1
    Do
        intCommandStart = InStr(intRawChar, strRawBuf, telIAC)
        If intCommandStart > 0 Then
            'Check for full command in buffer.
            If Len(strRawBuf) < intCommandStart + 2 Then
                Exit Do 'Come back when we've received more.
            Else
                'Check further (possible long command) and
                'process if complete.
                intCommandLen = ProcessReceivedCommand(intCommandStart)
                If intCommandLen < 1 Then
                    'Incomplete command.  Come back when we've received more.
                    Exit Do
                Else
                    'We've processed a command.  Extract any data
                    'preceding it, skip our index over the command.
                    strRcvData = strRcvData _
                               & Mid$(strRawBuf, intRawChar, intCommandStart - intRawChar)
                    intRawChar = intCommandStart + intCommandLen
                End If
            End If
        Else
            'No more commands in buffer.  Extract all data after
            'the last command as "received data."
            strRcvData = strRcvData & Mid$(strRawBuf, intRawChar)
            SendCommand
        End If
    Loop Until intCommandStart = 0
    
    strRawBuf = ""
    
    If Len(strRcvData) > 0 Then
        RaiseEvent DataArrival
    End If
End Sub

Private Sub wscTelnet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Number, Description)
    CancelDisplay = True
End Sub
