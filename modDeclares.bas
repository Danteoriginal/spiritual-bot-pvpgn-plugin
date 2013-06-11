Attribute VB_Name = "modDeclares"
Option Explicit

Public Host As Object

Public Declare Sub hash_password Lib "libbnet.dll" (ByVal Password As String, ByVal OutBuffer As String)

Public bool0x52 As Boolean
Public bool0x54 As Boolean
