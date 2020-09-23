VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIX PINGER"
   ClientHeight    =   4245
   ClientLeft      =   1740
   ClientTop       =   2025
   ClientWidth     =   5670
   Icon            =   "dsPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "dsPing.frx":0442
   ScaleHeight     =   4245
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar amount 
      Height          =   255
      LargeChange     =   10
      Left            =   1440
      Max             =   1000
      Min             =   1
      TabIndex        =   14
      Top             =   2880
      Value           =   1
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   2880
      TabIndex        =   13
      Text            =   "1"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      TabIndex        =   8
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ping Series"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txTTL 
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Text            =   "5"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton btnPing 
      Caption         =   "Ping Single"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox lbReturn 
      Height          =   1500
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox txIPAddress 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "0.0.0.0"
      Top             =   960
      Width           =   4575
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ping series pings the set you have next to the series label.  It does not do anything with the single box."
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   1440
      TabIndex        =   16
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5520
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Series"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time To Live (TTL) :"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Single:"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'   PingVB
'
'   This application implements a TCP/IP network ping
'   using the ICMP.DLL provided as part of Windows 95 and
'   Windows NT.
'   IP_STATUS codes returned from IP APIs
'
'*******************************************************************
Private Const IP_STATUS_BASE = 11000

Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
'
'   The next group are status codes passed up on status indications to
'   transport layer protocols.
'
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)

Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)

'   option information for network ping, we don't implement these here as this is
'   a simple sample (simon says).
Private Type ip_option_information
    Ttl             As Byte     'Time To Live
    Tos             As Byte     'Type Of Service
    Flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

'   structure that is returned from the ping to give status and error information
Private Type icmp_echo_reply
    Address         As Long             'Replying address
    Status          As Long             'Reply IP_STATUS, values as defined above
    RoundTripTime   As Long             'RTT in milliseconds
    DataSize        As Integer          'Reply data size in bytes
    Reserved        As Integer          'Reserved for system use
    DataPointer     As Long             'Pointer to the reply data
    Options         As ip_option_information    'Reply options
    Data            As String * 250     'Reply data which should be a copy of the string sent, NULL terminated
                                        ' this field length should be large enough to contain the string sent
End Type

'   declares for function to be used from icmp.dll
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, _
                                                    ByVal DestinationAddress As Long, _
                                                    ByVal RequestData As String, _
                                                    ByVal RequestSize As Integer, _
                                                    RequestOptions As ip_option_information, _
                                                    ReplyBuffer As icmp_echo_reply, _
                                                    ByVal ReplySize As Long, _
                                                    ByVal Timeout As Long) As Long

Private Const PING_TIMEOUT = 200        ' number of milliseconds to wait for the reply

Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYSSTATUS_LEN = 256
Private Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Private Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Private Const SOCKET_ERROR = -1

Private Type tagWSAData
        wVersion            As Integer
        wHighVersion        As Integer
        szDescription       As String * WSADESCRIPTION_LEN_1
        szSystemStatus      As String * WSASYSSTATUS_LEN_1
        iMaxSockets         As Integer
        iMaxUdpDg           As Integer
        lpVendorInfo        As String * 200
        End Type
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Private Declare Function WSACleanup Lib "wsock32" () As Integer

Private Sub amount_Change()
Label5.Caption = amount.Value
End Sub

Private Sub btnExit_Click()
    End
End Sub

'   btnPing
'
'   This routine is called when the button is clicked. The Ip address to be pinged
'   is taken from the text box and converted to a long value for the Icmp call
'
Private Sub pinger()

''''''''###I ADDED THE FOLLOWING 2 LINES ONTO YOUR CODE.... IT REALLY HELPS ALOT
lbReturn.AddItem ("-----------")
lbReturn.AddItem (txIPAddress.Text)
'''''####END OF WHAT I ADDED

    Dim hFile       As Long             ' handle for the icmp port opened
    Dim lRet        As Long             ' hold return values as required
    Dim lIPAddress  As Long
    Dim strMessage  As String
    Dim pOptions    As ip_option_information
    Dim pReturn     As icmp_echo_reply
    Dim iVal        As Integer
    Dim lPingRet    As Long
    Dim pWsaData    As tagWSAData
    
    strMessage = "Echo this string of data"
    
    iVal = WSAStartup(&H101, pWsaData)
    
    '   convert the IP address to a long, lIPAddress will be zero
    '   if the function failed. Normally you wouldn't ping if the address
    '   was no good to start with but we don't mind seeing bad return status
    '   as that is what samples are all about
    lIPAddress = ConvertIPAddressToLong(txIPAddress)
    
    '   open up a file handle for doing the ping
    hFile = IcmpCreateFile()
    
    '   set the TTL from the text box, try values of 1 to 255
    pOptions.Ttl = Val(txTTL)
    
    '   Call the function that actually does the ping. It is a blocking call so we
    '   don't get control back until it completes.
    lRet = IcmpSendEcho(hFile, _
                        lIPAddress, _
                        strMessage, _
                        Len(strMessage), _
                        pOptions, _
                        pReturn, _
                        Len(pReturn), _
                        PING_TIMEOUT)

    If lRet = 0 Then
        ' the ping failed for some reason, hopefully the error is in the return buffer
        lbReturn.AddItem "Ping failed with error " & pReturn.Status
        lbReturn.ListIndex = lbReturn.ListCount - 1
    Else
        ' the ping succeeded, .Status will be 0, .RoundTripTime is the time in ms for
        '   the ping to complete, .Data is the data returned (NULL terminated), .Address
        '   is the Ip address that actually replied, .DataSize is the size of the string in
        '   .Data
        If pReturn.Status <> 0 Then
            lbReturn.AddItem "Error -> Ping failed to complete, code = " & pReturn.Status
            lbReturn.ListIndex = lbReturn.ListCount - 1
        Else
            lbReturn.AddItem "Success -> completion time is " & pReturn.RoundTripTime & "ms."
            lbReturn.ListIndex = lbReturn.ListCount - 1
        End If
    End If
                        
    '   close the file handle that was used
    lRet = IcmpCloseHandle(hFile)
    
    iVal = WSACleanup()
    
End Sub

'
'   ConvertIPAddressToLong
'
'   Converts a dotted IP address (eg: "123.234.2.45") to a long
'   integer for use in sending a ping. This routine converts
'   the string as required by an Intel system.
'
'   Essentially we take the 4 numbers, flip them around and make
'   a long by shifting all the parts into the correct byte. We
'   do it here by making a hex string and converting it to a long.
'   Not pretty but it works (most of the time<g>).
'
'   When we get in "a.b.c.d" what we want out is Val(&Hddccbbaa).
'

Function ConvertIPAddressToLong(strAddress As String) As Long

    Dim strTemp             As String
    Dim lAddress            As Long
    Dim iValCount           As Integer
    Dim lDotValues(1 To 4)  As String
    
    ' set up the initial storage and counter
    strTemp = strAddress
    iValCount = 0
    
    ' keep going while we still have dots in the string
    While InStr(strTemp, ".") > 0
        iValCount = iValCount + 1   ' count the number
        lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)    ' pick it off and convert it
        strTemp = Mid(strTemp, InStr(strTemp, ".") + 1) ' chop off the number and the dot
        Wend
        
    ' the string only has the last number in it now
    iValCount = iValCount + 1
    lDotValues(iValCount) = strTemp
    
    ' if we didn't get four pieces then the IP address is no good
    If iValCount <> 4 Then
        ConvertIPAddressToLong = 0
        Exit Function
        End If
        
    '   take the four value, hex them, pad to 2 digits, make a hex
    '   string and then convert the whole mess to a long for returning
    lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & _
                Right("00" & Hex(lDotValues(3)), 2) & _
                Right("00" & Hex(lDotValues(2)), 2) & _
                Right("00" & Hex(lDotValues(1)), 2))
                
    '   set the return value
    ConvertIPAddressToLong = lAddress
    
End Function

Private Sub btnPing_Click()
For x = 1 To amount.Value
pinger
Next
End Sub

Private Sub Command1_Click()
startno = Text4.Text
endno = Text5.Text
For x = startno To endno
txIPAddress.Text = Text1.Text & "." & Text2.Text & "." & Text3.Text & "." & x
btnPing_Click
Next
 

End Sub

