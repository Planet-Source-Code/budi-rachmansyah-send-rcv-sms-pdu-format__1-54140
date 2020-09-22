Attribute VB_Name = "Receiving"
Option Explicit
Private mvarnoSCA As String
Private mvarnoOA As String
Private mvarFO As String
Private mvarDCS As String
Private mvarSCTS_Tgl As String
Private mvarSCTS_Jam As String
Private mvarSCTS_Tgl_A As String
Private mvarSCTS_Jam_A As String
Private mvarIndexSend As String
Private mvarUDL As String
Public Property Get vUDL() As String
    vUDL = mvarUDL
End Property
Public Property Get vnoSCA() As String
    vnoSCA = mvarnoSCA
End Property
Public Property Get IndexSend() As String
    IndexSend = mvarIndexSend
End Property
Public Property Get vnoOA() As String
    vnoOA = mvarnoOA
End Property
Public Property Get vFO() As String
    vFO = mvarFO
End Property
Public Property Get vDCS() As String
    vDCS = mvarDCS
End Property
Public Property Get vSCTS_Tgl() As String
    vSCTS_Tgl = mvarSCTS_Tgl
End Property
Public Property Get vSCTS_Jam() As String
    vSCTS_Jam = mvarSCTS_Jam
End Property
Public Property Get vSCTS_Tgl_A() As String
    vSCTS_Tgl_A = mvarSCTS_Tgl_A
End Property
Public Property Get vSCTS_Jam_A() As String
    vSCTS_Jam_A = mvarSCTS_Jam_A
End Property

Public Function TxtReceive(ByVal msg As String) As String
    Dim FO As String, PID As String, DCS As String, SCTS As String
    Dim UDL As String, UD As String, SCTS_Tgl As String, SCTS_Jam As String
    Dim lnSCA As String, typeSCA As String, noSCA As String
    Dim newMsg As String, lnOA As String, typeOA As String, noOA As String
    Dim SCTS_a As String, SCTS_Tgl_a As String, SCTS_Jam_a As String
    
    newMsg = msg
    lnSCA = HexToDec(Left(msg, 2)) * 2  'length of SCA
    newMsg = Right(newMsg, Len(newMsg) - 2)
    
    typeSCA = Left(newMsg, 2) '91:int,81:local
    newMsg = Right(newMsg, Len(newMsg) - 2)
    
    noSCA = RevNum(Left(newMsg, lnSCA - 2)) 'service center
    If UCase(Right(noSCA, 1)) = "F" Then noSCA = Left(noSCA, Len(noSCA) - 1)
    newMsg = Right(newMsg, Len(newMsg) - lnSCA + 2)
    
    FO = Left(newMsg, 2)
    newMsg = Right(newMsg, Len(newMsg) - 2)
    If FO = "06" Then 'code of send report (Indonesia)
        mvarIndexSend = HexToDec(Left(newMsg, 2))
        newMsg = Right(newMsg, Len(newMsg) - 2)
    End If
    
    'Origine Address
    lnOA = HexToDec(Left(newMsg, 2))
    If lnOA Mod 2 <> 0 Then lnOA = lnOA + 1
    newMsg = Right(newMsg, Len(newMsg) - 2)
    typeOA = Left(newMsg, 2)
    newMsg = Right(newMsg, Len(newMsg) - 2)
    
    noOA = Left(newMsg, lnOA)
    If typeOA = "D0" Then
        noOA = CharHex(noOA, 7)
    Else
        noOA = RevNum(noOA)
        If UCase(Right(noOA, 1)) = "F" Then noOA = Left(noOA, Len(noOA) - 1)
    End If
    newMsg = Right(newMsg, Len(newMsg) - lnOA)
   
    If FO <> "06" Then 'if not report message
        PID = Left(newMsg, 2)
        newMsg = Right(newMsg, Len(newMsg) - 2)
    
        DCS = Left(newMsg, 2)
        newMsg = Right(newMsg, Len(newMsg) - 2)
    
        SCTS = RevNum(Left(newMsg, 14))
        SCTS_Tgl = Mid(SCTS, 3, 2) & "/" & Mid(SCTS, 5, 2) & "/20" & Mid(SCTS, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam = Mid(SCTS, 7, 2) & ":" & Mid(SCTS, 9, 2) & ":" & Mid(SCTS, 11, 2) 'hh:mm:dd
        newMsg = Right(newMsg, Len(newMsg) - 14)
    
        UDL = CInt(HexToDec(Left(newMsg, 2)))
        newMsg = Right(newMsg, Len(newMsg) - 2)
    
        UD = CharHex(newMsg, 7)
        UD = Left(UD, UDL)
    Else
        SCTS = RevNum(Left(newMsg, 14))
        SCTS_Tgl = Mid(SCTS, 3, 2) & "/" & Mid(SCTS, 5, 2) & "/20" & Mid(SCTS, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam = Mid(SCTS, 7, 2) & ":" & Mid(SCTS, 9, 2) & ":" & Mid(SCTS, 11, 2) 'hh:mm:dd
        newMsg = Right(newMsg, Len(newMsg) - 14)
        
        SCTS_a = RevNum(Left(newMsg, 14))
        SCTS_Tgl_a = Mid(SCTS_a, 3, 2) & "/" & Mid(SCTS_a, 5, 2) & "/20" & Mid(SCTS_a, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam_a = Mid(SCTS_a, 7, 2) & ":" & Mid(SCTS_a, 9, 2) & ":" & Mid(SCTS_a, 11, 2) 'hh:mm:dd
        
        
    End If
    
    TxtReceive = UD
    mvarnoSCA = noSCA
    mvarnoOA = noOA
    mvarFO = FO
    mvarDCS = DCS
    mvarSCTS_Tgl = SCTS_Tgl
    mvarSCTS_Jam = SCTS_Jam
    mvarSCTS_Tgl_A = SCTS_Tgl_a
    mvarSCTS_Jam_A = SCTS_Jam_a
    mvarUDL = UDL
End Function
