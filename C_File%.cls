VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------
' C_File%.cls (c)  Serious Solutions
'
' Programmer:  Stephen S. Coleman
'
' Date: July 16, 2004
'
' Comments:
'
' Description: This is a class i'm just starting to build not many functions but its a start
'
' Modifications:  SSC - Created at 2:4 PM
'-------------------------------------------------------------------------------
Option Explicit ' FILE STREAM CLASS
'Requires Public ErrC As New C_Error
Private tstream As Object
Private Fs1 As Object

Private Sub Class_Terminate()

    Set tstream = Nothing
    Set Fs1 = Nothing

End Sub

Public Sub CloseFileObject()

    Set tstream = Nothing
    Set Fs1 = Nothing

End Sub

Public Sub NewFile(Filename As String)

    On Error GoTo Error_Handler '--Error Trap------------------------------------------------------
    If Err.Number Then
Error_Handler:
        ErrC.ErrorTrap "C_File%.cls", "NewFile", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '--Error Trap---------------------------------------------------------------------------

    Set Fs1 = CreateObject("Scripting.Filesystemobject")
    Set tstream = Fs1.CreateTextFile(Filename, True)

End Sub

Public Function OpenFile(Filename As String) As String

    On Error GoTo Error_Handler '--Error Trap------------------------------------------------------
    If Err.Number Then
Error_Handler:
        ErrC.ErrorTrap "C_File%.cls", "OpenFile", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '--Error Trap---------------------------------------------------------------------------

    Set Fs1 = CreateObject("Scripting.Filesystemobject")
    Set tstream = Fs1.opentextfile(Filename)
    OpenFile = tstream.readall

End Function

Public Sub WriteFile(Str As String)

    On Error GoTo Error_Handler '--Error Trap------------------------------------------------------
    If Err.Number Then
Error_Handler:
        ErrC.ErrorTrap "C_File%.cls", "WriteFile", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '--Error Trap---------------------------------------------------------------------------

    tstream.Writeline Str

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:36) 16 + 58 = 74 Lines
