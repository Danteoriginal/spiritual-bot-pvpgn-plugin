VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const Author = "Alendar"
Private Const Title = "PvPGN"
Private Const ReleaseDate = #6/11/2013#
Private Const Description = "A plugin for Spiritual Bot to handle issues that arise when dealing with private PvPGN servers."

'** Plugin Data
Public Property Get getAuthor() As String: getAuthor = Author: End Property
Public Property Get getTitle() As String: getTitle = Title: End Property
Public Property Get getReleaseDate() As Date: getReleaseDate = ReleaseDate: End Property
Public Property Get getDescription() As String: getDescription = Description: End Property
Public Property Get getVersion() As String: getVersion = App.Major & "." & App.Minor & "." & App.Revision: End Property
'** Plugin Data

Public Property Set Load(objHost As Object)

    On Error Resume Next
   
    Set Host = objHost
    
End Property

Public Sub Unload()

    On Error Resume Next
    
    Set Host = Nothing
    
End Sub

Public Function Initialize(ByVal intIndex As Integer) As Boolean

    On Error GoTo Err_Handler
    
    bool0x52 = True
    bool0x54 = True
    
    Initialize = True
    
    Exit Function
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "Initialize", Err.Description, intIndex
    Initialize = False
    
End Function

Public Sub Finalize(ByVal intIndex As Integer)

    On Error GoTo Err_Handler
    
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "Finalize", Err.Description, intIndex
    
End Sub

Public Sub ProfileUnloaded(ByVal intIndex As Integer)

    On Error GoTo Err_Handler
    
    Exit Sub

Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "ProfileUnloaded", Err.Description, intIndex
    
End Sub

Public Sub ActiveProfile(ByVal intIndex As Integer)

    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "ActiveProfile", Err.Description, intIndex
    
End Sub

Public Function BNET_Connected(ByVal intIndex As Integer) As Boolean
    
    On Error GoTo Err_Handler
    
    Dim boolIntercept As Boolean
    boolIntercept = False
    
    BNET_Connected = boolIntercept
    Exit Function
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "BNET_Connected", Err.Description, intIndex
    BNET_Connected = False
    
End Function

Public Function BNET_DataArrival(ByVal intIndex As Integer, ByVal strBuffer As String) As Boolean

    On Error GoTo Err_Handler
    
    Dim boolIntercept As Boolean
    Dim PacketID As Byte
    
    With Host.pProfile(intIndex).DeBuffer
        .DebuffPacket strBuffer
        PacketID = .DebuffPacketID
    
        Select Case PacketID
            Case &H26 'SID_READUSERDATA
                boolIntercept = True
            Case Else
                boolIntercept = False
        End Select
    End With
    
    BNET_DataArrival = boolIntercept
    Exit Function
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "BNET_DataArrival", Err.Description, intIndex
    BNET_DataArrival = False

End Function

Public Function BNET_DataDeparture(ByVal intIndex As Integer, strBuffer As String) As Boolean

    On Error GoTo Err_Handler
    
    Dim boolIntercept As Boolean
    Dim PacketID As Byte
    
    If Not bool0x52 Then Exit Function
    If Not bool0x54 Then Exit Function
    
    With Host.pProfile(intIndex).DeBuffer
        .DebuffPacket strBuffer
        PacketID = .DebuffPacketID
    
        Select Case PacketID
            Case &H26 'SID_READUSERDATA
                boolIntercept = True
            Case &H52 'SID_AUTH_ACCOUNTCREATE
                boolIntercept = True
                Send0x52 intIndex, strBuffer
                bool0x52 = True
            Case &H54 'SID_AUTH_ACCOUNTLOGONPROOF
                boolIntercept = True
                Send0x54 intIndex
                bool0x54 = True
            Case Else
                boolIntercept = False
        End Select
    End With
    
    BNET_DataDeparture = boolIntercept
    Exit Function
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "BNET_DataDeparture", Err.Description, intIndex
    BNET_DataDeparture = False
    
End Function

Public Sub BNET_Disconnected(ByVal intIndex As Integer)

    On Error GoTo Err_Handler

    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "BNET_Disconnected", Err.Description, intIndex
    
End Sub

Public Sub BNET_Error(ByVal intIndex As Integer, ByVal Number As Integer, ByVal Description As String)

    On Error GoTo Err_Handler
    
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "BNET_Error", Err.Description, intIndex
    
End Sub

'** Chat Events
Public Sub EID_SHOWUSER(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_SHOWUSER", Err.Description, intIndex
    
End Sub

Public Sub EID_JOIN(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_JOIN", Err.Description, intIndex
    
End Sub

Public Sub EID_LEAVE(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_LEAVE", Err.Description, intIndex
    
End Sub

Public Sub EID_WHISPER(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_WHISPER", Err.Description, intIndex
    
End Sub

Public Sub EID_TALK(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_TALK", Err.Description, intIndex
    
End Sub

Public Sub EID_BROADCAST(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_BROADCAST", Err.Description, intIndex
    
End Sub

Public Sub EID_CHANNEL(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_CHANNEL", Err.Description, intIndex
    
End Sub

Public Sub EID_USERFLAGS(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_USERFLAGS", Err.Description, intIndex
    
End Sub

Public Sub EID_WHISPERSENT(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_WHISPERSENT", Err.Description, intIndex
    
End Sub

Public Sub EID_CHANNELFULL(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_CHANNELFULL", Err.Description, intIndex
    
End Sub

Public Sub EID_CHANNELDOESNOTEXIST(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_CHANNELDOESNOTEXIST", Err.Description, intIndex
    
End Sub

Public Sub EID_CHANNELRESTRICTED(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_CHANNELRESTRICTED", Err.Description, intIndex
    
End Sub

Public Sub EID_INFO(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_INFO", Err.Description, intIndex
    
End Sub

Public Sub EID_ERROR(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
        
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_ERROR", Err.Description, intIndex
    
End Sub

Public Sub EID_EMOTE(ByVal intIndex As Integer, ByVal lngFlags As Long, ByVal lngPing As Long, ByVal strUsername As String, ByVal strMessage As String, ByVal strProduct As String, ByVal boolFiltered As Boolean)
    
    On Error GoTo Err_Handler
            
    Exit Sub
    
Err_Handler:
    Host.pFunctionError App.EXEName & ".clsPlugin", "EID_EMOTE", Err.Description, intIndex
    
End Sub
'** Chat Events
