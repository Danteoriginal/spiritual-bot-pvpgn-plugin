Attribute VB_Name = "modSendBNET"
Option Explicit

Public Sub Send0x52(ByVal intIndex As Integer, strBuffer As String)

    On Error GoTo Err_Handler
    
    Dim Salt As String
    Dim Username As String

    bool0x52 = False

    With Host.pProfile(intIndex).DeBuffer
        Salt = .DebuffRaw(32)
        .DebuffRaw 32
        Username = .DebuffNTString
        .Clear
    End With

    With Host.pProfile(intIndex).PBuffer
        .InsertNonNTString Salt
        .InsertNonNTString Left$(Host.pProfile(intIndex).BNET.Password & String(32, Chr(0)), 32)
        .InsertNTString CStr(Username)
        .SendPacket intIndex, &H52
    End With

    Host.pAddText intIndex, Host.pFrmMain.rtbChat(intIndex), RGB(0, 100, 0), "(BNET-PvPGN): Creating account..."
    
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".modSendBNET", "Send0x52", Err.Description, intIndex

End Sub

Public Sub Send0x54(ByVal intIndex As Integer)

    On Error GoTo Err_Handler
    
    Dim PWHash As String * 20
    
    bool0x54 = False

    With Host.pProfile(intIndex)
        hash_password LCase$(.BNET.Password), PWHash
        .PBuffer.InsertNonNTString CStr(PWHash)
        .PBuffer.SendPacket intIndex, &H54
    End With

    Host.pAddText intIndex, Host.pFrmMain.rtbChat(intIndex), RGB(0, 100, 0), "(BNET-PvPGN): Sending login information..."
        
    Exit Sub

Err_Handler:
    Host.pFunctionError App.EXEName & ".modSendBNET", "Send0x54", Err.Description, intIndex

End Sub


