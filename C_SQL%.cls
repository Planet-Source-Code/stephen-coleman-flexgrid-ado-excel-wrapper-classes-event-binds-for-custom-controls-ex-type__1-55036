VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' C_SQL%.cls (c)  Serious Solutions
'
' Programmer:  Stephen S. Coleman
'
' Date: July 16, 2004
'
' Comments: I use this class for all my SQL or MDB connections
'
' Description: Self Contained SQL Connection\RecordSet
'
' Modifications:  SSC - Created at 1:59 PM
'-------------------------------------------------------------------------------
Option Explicit
'Requires Public ErrC As New C_Error
'Refrence Microsoft ActiveX Data Objects
Public RS    As ADODB.Recordset
Private CN    As ADODB.Connection
Private M_DATABASE As String
Private M_Server As String
Private M_CommandTimeout As Integer
Private M_ConnectionTimeout As Integer
Private M_LastSqlQuery As String
Private ConnectionSet As Boolean
Public LeaveConnectionOpen As Boolean
Public OutPutToQueryA As Boolean
Public PrntSql As Boolean
Public shownotepad As Boolean

Private Sub Class_Initialize()

    ConnectionSet = False
    LeaveConnectionOpen = True
    Set RS = Nothing
    Set CN = Nothing

End Sub

Private Sub Class_Terminate()

    Set CN = Nothing
    Set RS = Nothing

End Sub

Public Function ClipSqlFormated(StrSQL As String, Clip As Boolean) As String

    ClipSqlFormated = Replace(UCase$(StrSQL), "WHERE", vbCrLf & "WHERE")
    ClipSqlFormated = Replace(ClipSqlFormated, "FROM", vbCrLf & "FROM")
    ClipSqlFormated = Replace(ClipSqlFormated, "HAVING", vbCrLf & "HAVING")
    ClipSqlFormated = Replace(ClipSqlFormated, "GROUP", vbCrLf & "GROUP")
    ClipSqlFormated = Replace(ClipSqlFormated, "LEFT", vbCrLf & "LEFT")
    ClipSqlFormated = Replace(ClipSqlFormated, "RIGHT", vbCrLf & "RIGHT")
    ClipSqlFormated = Replace(ClipSqlFormated, "INNER", vbCrLf & "INNER")
    If Clip Then
        Clipboard.SetText ClipSqlFormated
    End If

End Function

Public Property Get CommandTimeOut() As String

    CommandTimeOut = LeaveConnectionOpen

End Property

Public Property Let CommandTimeOut(ByVal vNewValue As String)

    M_CommandTimeout = vNewValue

End Property

Public Property Get ConnectionTimeOut() As String

    ConnectionTimeOut = M_ConnectionTimeout

End Property

Public Property Let ConnectionTimeOut(ByVal vNewValue As String)

    M_ConnectionTimeout = vNewValue

End Property

Public Property Get DataBase() As String

    DataBase = M_DATABASE

End Property

Public Property Let DataBase(ByVal vNewValue As String)

    If Server = "Access Database" Then
        Server = "" ' Resets Server if not using MDB so you can't try to connect
    End If
    If InStr(UCase$(vNewValue), ".MDB") <> 0 Then
        Server = "Access Database"
    End If
    M_DATABASE = vNewValue
    ConnectionSet = False

End Property

Private Sub LastQuery()

    sql M_LastSqlQuery

End Sub

Public Property Get LastQueryStr() As String

    LastQueryStr = M_LastSqlQuery

End Property

Private Function OpenConnection() As Boolean

    On Error GoTo Error_Handler '--Error Trap-----------------------------------------------
    If Err.Number Then
Error_Handler:
        OpenConnection = False
        Exit Function
    End If '--Error Trap---------------------------------------------------------------------------

    If Not ConnectionSet Then
        If M_Server <> "" And M_DATABASE <> "" Then
            Set CN = New Connection
            With CN
                .CursorLocation = adUseClient
                .CommandTimeOut = M_CommandTimeout
                .ConnectionTimeOut = M_ConnectionTimeout
                If InStr(UCase$(M_DATABASE), ".MDB") = 0 Then
                    .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & M_DATABASE & ";Data Source=" & M_Server
                  Else 'For Access.MDB Connection'NOT INSTR(UCASE(M_DATABASE),...
                    .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & M_DATABASE & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                End If
            End With
            ConnectionSet = True
          Else
            Exit Function
        End If
    End If
    OpenConnection = True

End Function

Public Function SendQuery(StrSQL As String) As Boolean

    On Error GoTo exitfun
    If OpenConnection Then
        Set RS = New Recordset
        If RS.State <> 1 Then
            RS.CursorLocation = adUseServer
        End If
        If PrntSql Then
            sql StrSQL
        End If
        M_LastSqlQuery = StrSQL
        Set RS = CN.Execute(StrSQL)
        If Not LeaveConnectionOpen Then
            Set RS.ActiveConnection = Nothing
            Set CN = Nothing
            ConnectionSet = False
        End If
      Else
        Exit Function
    End If
    SendQuery = True
exitfun:

End Function

Public Function SendQueryNoRecordset(StrSQL As String) As Boolean

    On Error GoTo exitfun
    If OpenConnection Then
        If PrntSql Then
            sql StrSQL
        End If
        M_LastSqlQuery = StrSQL
        CN.Execute StrSQL, , adExecuteNoRecords
        If Not LeaveConnectionOpen Then
            Set CN = Nothing
            ConnectionSet = False
        End If
    End If
    SendQueryNoRecordset = True
exitfun:

End Function

Public Property Get Server() As String

    Server = M_Server

End Property

Public Property Let Server(ByVal vNewValue As String)

    M_Server = vNewValue
    ConnectionSet = False

End Property

Private Sub sql(sqlstr As String)

  Dim strErrorLog As String

    strErrorLog = FreeFile()
    Open "c:\C_SQL.LOG" For Append As strErrorLog
    Print #strErrorLog, Now & Chr$(9) & ClipSqlFormated(sqlstr, True)
    Close #strErrorLog
    If shownotepad Then
        Shell "notepAD.exe c:\C_SQL.log", vbMaximizedFocus
    End If
    If OutPutToQueryA Then
        Open "c:\C_SQL1.SQL" For Output As strErrorLog
        Print #strErrorLog, ClipSqlFormated(sqlstr, True)
        Close #strErrorLog
        Shell "C:\Program Files\Microsoft SQL Server\80\Tools\Binn\isqlw.exe  c:\C_SQL1.SQL", vbMaximizedFocus
    End If

End Sub

Public Sub WriteToFile(Path As String, Filename As String, Optional Delimeter As String)
Dim Cfile As New C_File 'Requires C_File%.cls
    Cfile.NewFile Path & Filename
    RS.MoveFirst
    If Delimeter = "" Then
        Delimeter = vbTab
    End If
    'CFile.WriteFile (RS.GetString(adClipString, , vbTab, vbCr, vbNullString)) ' CSV TAB Format
    Cfile.WriteFile (RS.GetString(adClipString, , Delimeter, vbCr, vbNullString)) ' CSV Comma Format
    Cfile.CloseFileObject

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:36) 31 + 210 = 241 Lines
