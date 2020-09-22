Attribute VB_Name = "Sending"
Private mvarMsgType As String
Private mvarMsgReport As String
Private mvarMsgTime As String
Public Property Let MsgType(ByVal data As String)
    mvarMsgType = data
End Property
Public Property Let MsgTime(ByVal period As Integer)
    Dim data As String
    Select Case period
        Case 1: data = "0B" '11d/0Bh utk 1 jam interval
        Case 2: data = "8F" '143d : 12 hour
        Case 3: data = "A7" '167d : 1 day
        Case 4: data = "A8" '167d : 2 day
        Case 5: data = "AD" '167d : 1 week
    End Select
    mvarMsgTime = data
End Property
Public Property Let MsgReport(ByVal data As String)
    mvarMsgReport = data
End Property
Public Function TxtSend(ByVal DestinationNo As String, ByVal Message As String)
On Error Resume Next
    Dim SCA As String, PDU As String, MR As String
    Dim DA As String, PID As String, DCS As String
    Dim VP As String, UDL As String, UD As String
    
    
    SCA = "00"
    PDU = mvarMsgReport 'unreceived:"11"/received:"31")
    If PDU = "" Then PDU = "11" 'default:unreceived
    MR = "00"
    
    'DA: Destination Address
    
    DA = DecToHex(Len(DestinationNo)) 'Panjang DestinationNo dlm Hex
    DA = DA & "91" '"91":Int. Number(62...),"81":Loc. Number(081..)
    DA = DA & RevNum(DestinationNo)
    
    PID = "00"
    DCS = mvarMsgType 'Normal:"00",Flash:"F0"
    If DCS = "" Then DCS = "00" 'default normal
    
    VP = mvarMsgTime 'Limit Period of delivery
    If VP = "" Then VP = "A7" ' default:1 days
    
    UDL = DecToHex(Len(Message)) ' length of message in Hex
    UD = CharHex(Message, 8) 'Message in Hex 8bit /octet
    
    'Format of SMS Submit PDU
    TxtSend = SCA & PDU & MR & DA & PID & DCS & VP & UDL & UD

End Function
