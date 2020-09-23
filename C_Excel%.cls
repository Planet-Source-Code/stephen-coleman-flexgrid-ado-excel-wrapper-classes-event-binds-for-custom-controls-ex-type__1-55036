VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' C_Excel%.cls (c)  Serious Solutions
'
' Programmer:  Stephen S. Coleman
'
' Date: July 16, 2004
'
' Comments: Copies from a FlexGrid to Excel.  Assumes you have top Row Column Names
'
' Description: Very Simple Excel class
'
' Modifications:  SSC - Created at 2:7 PM
'-------------------------------------------------------------------------------

Option Explicit
'Refrence Microsoft Excel Oject Library
Private XLS As New Excel.Application 'Tested with Office 97
'Requires Public ErrC As New C_Error
Private Stor As New C_Stack 'Requires C_Stack%.cls
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private ExcelWasNotRunning As Boolean
Private PreviousRow  As Long

Private Wbs As Excel.Workbooks
Private Wb  As Excel.Workbook
Private ws  As Excel.Worksheet
Private XRange As Range
Private Const xlMaximized = &HFFFFEFD7
Public ReportPath As String
Public Title As String

Private Sub Class_Initialize()

    ReportPath = App.Path
    Title = "Testing"

End Sub

Private Sub Class_Terminate()

    XLS.Quit
    Set XLS = Nothing
    Set Wb = Nothing

End Sub

Public Sub ClipReport(Grd As C_FlexGrid, showEXCEL As Boolean, Save As Boolean, Optional XLfilename As String, Optional stacked As Boolean)

  Dim Cntr As Integer
  Dim XLArrayStr As String

    On Error GoTo Error_Handler '--Error Trap------------------------------------------------------
    If Err.Number Then
Error_Handler:
        If Err.Number = -2147417851 Then
            Wb.Save
            Resume Next
          ElseIf Err.Number = 1004 Then
            XLS.StatusBar = "The Excel File is locked.  WorkBook not saved!!!!"
            Resume Next
        End If
        ErrC.ErrorTrap "C_Excel%.cls", "ClipReport", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '--Error Trap---------------------------------------------------------------------------

    If Grd.Flxgrd.Rows <= 1 Then
        Exit Sub  'if grd empty exit
    End If
    If Not stacked Then
        PreviousRow = 0
        Set XLS = New Excel.Application
    End If

    If XLfilename = "" Then
        Set Wb = XLS.Workbooks.Add
      Else
        XLfilename = Replace((ReportPath & "\" & XLfilename & ".xls"), "/", " ")
    End If

    Set ws = Wb.ActiveSheet
    ws.Activate
    XLS.Visible = True
    With XLS
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .StatusBar = "Writing data from Software Management Tool..."
    End With

    SetupTitle
    If Grd.Flxgrd.Rows > 600 Then
        CopyByLine Grd, PreviousRow, stacked
      Else
        CopyAll Grd, PreviousRow, stacked
    End If

    FormatPAstedData Grd

    With XLS
        .StatusBar = False
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    If Save Then
        XLS.DisplayAlerts = False
        XLS.AlertBeforeOverwriting = False
        Wb.SaveAs Replace((ReportPath & "\" & Stor.Pop & ".xls"), "/", " "), True
        'XLS.DisplayAlerts = True
    End If
    If showEXCEL Then
        XLS.Visible = True
      ElseIf XLfilename <> "C:\Program Files\Software Management Tool\Report.xls" Then 'NOT SHOWEXCEL...
        XLfilename = "C:\Program Files\Software Management Tool\Report.xls"
    End If

End Sub

Public Sub ClipToXl(Grd As C_FlexGrid, Save As Boolean, Optional XLfilename As String)

  Dim Cntr As Integer
  Dim XLArrayStr As String

    On Error GoTo Error_Handler '**Error Trap**
    If Err.Number Then
Error_Handler:
        If Err.Number = -2147417851 Then
            Wb.Save
            Resume Next
          ElseIf Err.Number = 1004 Then
            XLS.StatusBar = "The Excel file is locked or invalid file name.  WorkBook not saved!!!!"
            Resume Next
        End If
        ErrC.ErrorTrap "C_Excel%.cls", "Class_Initialize", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '**Error Trap**

    If Grd.Flxgrd.Rows <= 1 Then
        Exit Sub  'if grd empty exit
    End If
    Set XLS = New Excel.Application
    Set Wb = XLS.Workbooks.Add
    XLfilename = Replace((ReportPath & "\" & XLfilename & ".xls"), "/", " ")
    Set ws = Wb.ActiveSheet
    ws.Activate
    XLS.Visible = True
    With XLS
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .StatusBar = "Writing data from " & App.Title & "..."
    End With

    SetupTitle

    If Grd.Flxgrd.Rows > 600 Then
        CopyByLine Grd, 0, False
      Else
        CopyAll Grd, 0, False
    End If

    FormatPAstedData Grd

    With XLS
        .StatusBar = False
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    If Save Then
        XLS.DisplayAlerts = False
        XLS.AlertBeforeOverwriting = False
        Wb.SaveAs XLfilename, True
        'XLS.DisplayAlerts = True
    End If

End Sub

Private Sub CopyAll(Grd As C_FlexGrid, PreviousRow As Long, stacked As Boolean)

  Dim LOOPCOPY As Integer
  Dim Lcol As Long
  Dim Rcol As Long
  Dim Rrow As Long
  Dim Lrow As Long
  Dim lastloop As Long

    If PreviousRow = 0 Or Not stacked Then
        Lcol = 1
        Lrow = 4
        Rcol = Grd.Flxgrd.Cols
        Rrow = LOOPCOPY + 4
        PreviousRow = LOOPCOPY + 4
      Else
        Lcol = 1
        Lrow = PreviousRow + 2
        Rcol = Grd.Flxgrd.Cols
        Rrow = PreviousRow + 1 + Grd.Flxgrd.Rows
    End If

    Set XRange = ws.Range(ws.Cells(Lrow, Lcol), ws.Cells(Rrow, Rcol))
    XRange.HorizontalAlignment = xlCenter
    Grd.EditSelectandCopy True, True
    XRange.PasteSpecial xlPasteAll, xlPasteSpecialOperationNone

End Sub

Private Sub CopyByLine(Grd As C_FlexGrid, PreviousRow As Long, stacked As Boolean)

  Dim LOOPCOPY As Integer
  Dim Lcol As Long
  Dim Rcol As Long
  Dim Rrow As Long
  Dim Lrow As Long
  Dim lastloop As Long
  Dim STEPINT As Integer
  Dim StartAtRow As Integer
  Dim XlLrow As Long
  Dim XlRrow As Long

    On Error GoTo Error_Handler '**Error Trap**
    If Err.Number Then
Error_Handler:
        ErrC.ErrorTrap "C_Excel%.cls", "CopyByLine", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '**Error Trap**

    STEPINT = 400
    StartAtRow = 4 ' Starts under  title
    For LOOPCOPY = STEPINT To Grd.Flxgrd.Rows Step STEPINT
        If PreviousRow = 0 Or Not stacked Then
            Lcol = 1
            If PreviousRow = 0 Then
                Lrow = StartAtRow ' Starts under  title
                Rcol = Grd.Flxgrd.Cols
                Rrow = LOOPCOPY
                PreviousRow = Rrow
                XlRrow = Rrow + StartAtRow
                XlLrow = StartAtRow
              Else
                XlLrow = XlRrow + 1
                XlRrow = XlLrow + STEPINT
                'Lrow = lastloop
                Rcol = Grd.Flxgrd.Cols
                Rrow = Rrow + STEPINT + 1
                PreviousRow = Rrow
            End If
          Else
            Lcol = 1
            Lrow = PreviousRow + 2
            Rcol = Grd.Flxgrd.Cols
            Rrow = PreviousRow + 1 + Grd.Flxgrd.Rows
        End If

        If (Grd.Flxgrd.Rows - Rrow) < STEPINT Then 'Paste the rest if next set is less than step amount
            'XlRrow = XlRrow + (Grd.Flxgrd.Rows - Rrow) - 1
            XlRrow = XlLrow + Grd.Flxgrd.Rows - Rrow
        End If
        Set XRange = ws.Range(ws.Cells(XlLrow, Lcol), ws.Cells(XlRrow, Rcol))
        XRange.HorizontalAlignment = xlCenter

        With XLS
            XRange.Select
            .StatusBar = False
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
        End With
        Grd.EditSelectandCopy True, True, Rrow, lastloop
        XRange.PasteSpecial xlPasteAll, xlPasteSpecialOperationNone
        lastloop = STEPINT + lastloop + 1

    Next LOOPCOPY

End Sub

Private Sub FormatPAstedData(Grd As C_FlexGrid)

    Set XRange = ws.Range(ws.Cells(4, 1), ws.Cells(4, Grd.Flxgrd.Cols))
    'XRange.Select
    XRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    XRange.Borders(xlEdgeLeft).Weight = xlMedium
    XRange.Borders(xlEdgeLeft).ColorIndex = 0
    XRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    XRange.Borders(xlEdgeTop).Weight = xlMedium
    XRange.Borders(xlEdgeTop).ColorIndex = 0
    XRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    XRange.Borders(xlEdgeBottom).Weight = xlMedium
    XRange.Borders(xlEdgeBottom).ColorIndex = 0
    XRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    XRange.Borders(xlEdgeRight).Weight = xlMedium
    XRange.Borders(xlEdgeRight).ColorIndex = 0
    XRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    XRange.Borders(xlInsideVertical).Weight = xlMedium
    XRange.Borders(xlInsideVertical).ColorIndex = 0
    Set XRange = ws.Range("A1", "M9000")
    XRange.EntireColumn.AutoFit
    ws.Columns.Range("A1", "A9000").ColumnWidth = 6
    ws.PageSetup.CenterFooter = "&""Arial,Bold""&8&UCompany Use Only"
    ws.PageSetup.LeftFooter = "&""Arial,Bold""&8Page &P  &F"
    ws.PageSetup.RightFooter = "&""Arial,Bold""&8 " & Date

End Sub

Private Sub SetupTitle()

  'Creates Title on XL sheet

    ws.Cells(1, 1).Value = Title
    ws.Cells(2, 1).Value = "Second Title"
    With ws.Cells.Font
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .ColorIndex = xlAutomatic
    End With
    Set XRange = ws.Range("A1:I1")
    With XRange
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .ShrinkToFit = False
    End With
    With XRange.Font
        .Name = "Times New Roman"
        .FontStyle = "Bold"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Set XRange = ws.Range("A2:I2")
    With XRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .ShrinkToFit = False
        .MergeCells = True
    End With
    With XRange.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleSingle
        .ColorIndex = xlAutomatic
    End With

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:37) 33 + 330 = 363 Lines
