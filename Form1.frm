VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "SMS PDU FORMAT by boedot"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Tx Receive"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox TxPDU 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   6135
   End
   Begin VB.TextBox TxDNo 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3240
      Width           =   6135
   End
   Begin MSCommLib.MSComm MSC 
      Left            =   7440
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   57600
   End
   Begin VB.TextBox TxMsg 
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton CmdZ 
      Caption         =   "Ctrl Z"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton CmdEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "PDU (put PDU Format  from SMS Receive)"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "62815234567      (62 = country code)"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Destination No."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "AT Command"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I have tested this code for some GSM Provider in my country (Indonesia)successfully.
'The first step before you learn my code, you should be have a knowledge about SMS PDU Format.
'You can find out this in search engine.
'I'm not very good in VB, but I always try the best.
'I don't mind to disscuss this topic later for sharing knowledge.
'email:boedot@hotpop.com

Dim Hookdata As String
Private Sub CmdEnter_Click()
    MSC.Output = Text5.Text + Chr(13)
    Text5.Text = ""
End Sub
Private Sub CmdZ_Click()
    MSC.Output = Text5.Text + Chr(26)
End Sub

Private Sub Command1_Click()
    ''Non Flash="00" (default); Flash="F0"
    'MsgType = "00"
    
    ''Display Report SMS : no="11"(default);yes="31"
    MsgReport = "31"
    
    ''Validity 1-5 ; 1=1hour,2=12hour,3=1day(default),4=2day,5=1week
    'MsgTime = 4
    
    'sendmsg destinationno,message
    SendMsg TxDNo.Text, TxMsg.Text
End Sub
Private Sub Command2_Click()
    ReadMessage TxPDU.Text
End Sub

Private Sub Form_Load()
   MSC.CommPort = 1
   MSC.Settings = "19200,n,8,1"
   MSC.PortOpen = True
   MSC.Handshaking = comNone

End Sub

Public Sub ReadMessage(ByVal xData As String)
    Dim RMsg As String
    Dim FO As String, i As Integer, DCS As String
    Dim OA As String, Tgl As String, Jam As String
    Dim Tgl_Akhir As String, Jam_akhir As String, UDL As String, SCA As String
    

        RMsg = TxtReceive(xData)

        FO = vFO
        SCA = vnoSCA
        OA = vnoOA
        DCS = vDCS
        Tgl = vSCTS_Tgl
        Jam = vSCTS_Jam
        Tgl_Akhir = vSCTS_Tgl_A
        Jam_akhir = vSCTS_Jam_A
        UDL = vUDL
        
    RMsg = "Message=" & RMsg & vbCrLf
    RMsg = RMsg & "FO =" & FO & vbCrLf
    RMsg = RMsg & "SCA =" & SCA & vbCrLf
    RMsg = RMsg & "OA =" & OA & vbCrLf
    RMsg = RMsg & "DCS =" & DCS & vbCrLf
    RMsg = RMsg & "Tgl =" & Tgl & vbCrLf
    RMsg = RMsg & "Jam =" & Jam & vbCrLf
    RMsg = RMsg & "Tgl_A =" & Tgl_Akhir & vbCrLf
    RMsg = RMsg & "Jam_A =" & Jam_akhir & vbCrLf
    RMsg = RMsg & "UDL =" & UDL & vbCrLf
    
    MsgBox RMsg
End Sub
Public Sub SendMsg(ByVal DNo As String, ByVal xData As String)
        MSC.Output = "AT+CMGS=" & Len(TxtSend(DNo, xData)) / 2 - 1 & Chr(13)
        Tunda 0.1
        MSC.Output = TxtSend(DNo, xData) & Chr(26)
        
End Sub

Private Sub MSC_OnComm()
    If MSC.CommEvent = comEvReceive Then
            Hookdata = Hookdata & MSC.Input
            Text2.Text = Hookdata
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
    Case 13: Call CmdEnter_Click
    Case 26: Call CmdZ_Click
   End Select
End Sub
