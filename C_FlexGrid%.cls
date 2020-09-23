VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_FlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' C_FlexGrid%.cls (c)  Serious Solutions
'
' Programmer:  Stephen S. Coleman
'
' Date: July 16, 2004
'
' Comments: You will need C_SQL%.CLS to make many of the SQL related functions work.
'
' Description: Hooks the events of a Flexgrid on a form or Adds a Flexgrid dynamic to the form
' Adds TypeAhead and Alpha Sorting
'
' Modifications:  SSC - Created at 2:1 PM
'-------------------------------------------------------------------------------
Option Explicit
'Requires Public ErrC As New C_Error
Private Stor As New C_Stack 'Requires C_Stack%.cls
'Stack Class is used here to store column names and sizes for custom sizing
Const BuildGridInt = 115 ' Integer for Coulmn Width autosize... Increase for wider coulmns
Private GrdSortAtoZ As Variant ' For Grd Alpha sort
Private GrdSortZtoA As Variant ' For Grd Alpha sort
Public WithEvents Flxgrd As MSFlexGrid ' To get events and bind handling of FlexGrid Control
Attribute Flxgrd.VB_VarHelpID = -1
Private StupidTimer As Date  '  Typeahead
Private TypeAheadStor As Variant  '  Typeahead
Private FlexGridRowStor As Long ' TypeAhead

Public Sub AddColumn(ColumnName As String, Width As Integer, ColumnNum As Integer) ' Adds a column to the flex grid

  Dim Cnt As Integer

    With Flxgrd
        If .Rows <> 0 Then
            .Visible = False
            .Cols = .Cols + 1
            If ColumnNum < .Cols Then
                .Col = ColumnNum - 1
                EditSelectColumn ColumnNum - 1, .Cols - 2
                EditCopy
                'EditDelete  ' remmed this for faster run... Un rem if you want the new column to be blank
                EditSelectColumn ColumnNum, .Cols - 1
                EditPaste
                For Cnt = 0 To .Cols - ColumnNum - 1
                    .Col = .Cols - 1 - Cnt
                    .ColWidth(.Col) = .ColWidth(.Col - 1)
                Next Cnt
              Else
                .Col = ColumnNum - 1
                .Row = 0
                .Text = ColumnName
                .ColWidth(.Cols - 1) = Width
            End If

            .Visible = True
        End If
    End With

End Sub

Public Sub CheckBox(column As Integer)

  Dim y As Integer
  Const strChecked = "þ"
  Const strUnChecked = "q"

    AddColumn "CheckBox", 1800, column ' Adds new column
    With Flxgrd
        If .Rows <> 0 Then
            .Visible = False
            .Row = 0
            .Col = column - 1
            .Text = "Check"
            .ColWidth(.Col) = 600
            For y = 1 To .Rows - 1
                'For x = 1 To .Cols - 1
                .Row = y
                .Col = column - 1
                .CellFontName = "Wingdings"
                .CellFontSize = 14
                .CellAlignment = flexAlignCenterCenter
                .Text = strUnChecked
                .Visible = True
            Next y
        End If
    End With

End Sub

Public Sub CheckboxTrigger(iRow As Integer, iCol As Integer)

    Dim Rowstor As Long
    Dim ColStor As Long
    With Flxgrd
        If .Rows > 0 Then
            Rowstor = .Row
            ColStor = .Col
            .Row = 0
            .Col = iCol
            If .Text = "Check" Then
                If iRow <> 0 And iCol <> 0 Then
                    If .TextMatrix(iRow, iCol) = "q" Then
                        .TextMatrix(iRow, iCol) = "þ"
                      Else
                        .TextMatrix(iRow, iCol) = "q"
                    End If
                End If
            End If
        .Row = Rowstor
        .Col = ColStor
        End If

    End With
End Sub
Private Sub Class_Terminate()

    Set Flxgrd = Nothing

End Sub
Public Sub Colors(R As Integer, G As Integer, B As Integer) ' Set the Colors of a flexgrid


   EditSelectAll

    With Flxgrd

     ' .Visible = False
        If .Rows > 0 Then
            Randomize
            .FillStyle = flexFillRepeat
            .CellBackColor = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            '.CellForeColor = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            '.ForeColor = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            .ForeColorFixed = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            .ForeColorSel = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            .GridColor = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            .GridColorFixed = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
             .BackColorFixed = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
             .BackColorBkg = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
             '.BackColorSel = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
             .BackColor = RGB(R + Int((255 * Rnd) + 1), G + Int((255 * Rnd) + 1), B + Int((255 * Rnd) + 1))
            .Refresh
        End If
      .Visible = True

    End With

End Sub


Private Sub ColumnSet()

  Dim Count As Integer
  Dim arasplit() As String

    ReDim arasplit(2) As String

    Flxgrd.Rows = 2
    Flxgrd.Cols = 2
    Flxgrd.FixedRows = 1
    Flxgrd.FixedCols = 0
    Flxgrd.Rows = 1
    Flxgrd.Cols = 1
    Flxgrd.Clear
    Flxgrd.Row = 0
    Flxgrd.Cols = Stor.Count

    For Count = Stor.Count To 1 Step -1
        arasplit = Split(Stor.Pop, ",")
        Flxgrd.Col = Count - 1
        Flxgrd.Text = arasplit(0)
        Flxgrd.ColWidth(Count - 1) = arasplit(1)
    Next Count
    Stor.Clear

End Sub

Public Sub dumptofile(Path As String, Filename As String, Optional Show As Boolean) ' Very Simple way to dump a FlexGrid to a file.

    On Error GoTo Error_Handler '**Error Trap**
    If Err.Number Then
Error_Handler:
        ErrC.ErrorTrap "C_FlexGrid%.cls", "dumptofile", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '**Error Trap**

    EditSelectAllWithFixed ' By changeing the select you can dump out different parts of the file
    Open Path & "\" & Filename For Output As #1
    Print #1, Replace(Flxgrd.Clip, vbTab, ",") ' Clip deafults as a tab delimeted file, Here i changed the tabs to commas
    Close #1
    If Show Then
        Shell "Notepad.exe " & Path & "\" & Filename, vbNormalFocus
    End If

End Sub

Public Sub DynamicAddToform(frm As Form, Height As Long, Width As Long, Top As Long, Left As Long, GridCounter As Integer)

    Set Flxgrd = frm.Controls.Add("Msflexgridlib.MSFlexGrid", "Grd" & GridCounter, frm)
    With Flxgrd
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 0
        .Cols = 0
        .Visible = True
        .Height = Height
        .Width = Width
        .Left = Left
        .Top = Top
    End With

End Sub

Public Sub DynamicRemove(frm As Form, ControlName As String)

    On Error GoTo 0
    frm.Controls.Remove ControlName

End Sub

Public Sub EditCopy()

  'Copy the selection and put it on the Clipboard

    Clipboard.Clear
    Clipboard.SetText Flxgrd.Clip

End Sub

Public Sub EditCut()

  'Cut the selection and put it on the Clipboard

    EditCopy
    EditDelete

End Sub

Public Sub EditDelete()

  'Deletes the selection

  Dim i As Integer
  Dim j As Integer
  Dim strClip As String

    With Flxgrd
        For i = 1 To .RowSel
            For j = 1 To .ColSel
                strClip = strClip & "" & vbTab
            Next j
            strClip = strClip & vbCr
        Next i
        .Clip = strClip
    End With

End Sub

Public Sub EditPaste()

  'Insert Clipboard contents

    If Len(Clipboard.GetText) Then
        Flxgrd.Clip = Replace(Clipboard.GetText, ChrW$(10), "")
    End If

End Sub

Public Sub EditSelectAll()

  'Selects the whole Grid excludes fixed rows

    With Flxgrd
        If .Rows > 0 Then
            .Visible = False
            .Row = 1
            .Col = 0
            .RowSel = .Rows - 1
            .ColSel = .Cols - 1
            .TopRow = 1
            .Visible = True
        End If
    End With

End Sub

Public Sub EditSelectAllWithFixed()

  'Selects the whole Grid excludes fixed rows

    With Flxgrd
        .Visible = False
        .FixedRows = 0
        .FixedCols = 0
        .Row = 0
        .Col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .TopRow = 1
        .Visible = True
    End With

End Sub

Public Sub EditSelectandCopy(Optional FixRows As Boolean, Optional FixCols As Boolean, Optional NumRows As Long, Optional StartAtRow As Long)

  Dim Fr As Integer
  Dim Fc As Integer

    With Flxgrd
        If NumRows > 0 Then  ' Selected rows copied
            If .Rows > 0 Then
                .Visible = False
                .Row = StartAtRow
                .Col = 0
                .RowSel = NumRows
                .ColSel = .Cols - 1
                .Visible = True
                EditCopy
            End If
          ElseIf .Rows > 0 Then ' All rows copied
            .Visible = False
            .Row = 0
            .Col = 0
            .RowSel = .Rows - 1
            .ColSel = .Cols - 1
            .Visible = True
            EditCopy
        End If
    End With

End Sub

Public Sub EditSelectColumn(FirstColumn As Integer, lastCoulmn As Integer)

  'Selects the whole Grid

    With Flxgrd
        '.Visible = False
        .Row = 0
        .Col = FirstColumn
        .RowSel = .Rows - 1
        .ColSel = lastCoulmn
        If .Rows > 1 Then
            .TopRow = 1
        End If
        ' .Visible = True
    End With

End Sub
'Public Sub Typeahead(Key As String)
'
'  Dim cmbInd As Long, lstInd As Long
'
'    If DateDiff("s", Now(), StupidTimer) < -1 Then 'Greater than 1 seconds start over with first char
'        Stop
'    End If
'    StupidTimer = Now()
'    For cmbInd = (Flxgrd.Rows - 1) To 0 Step -1
'        Flxgrd.Row = cmbInd
'        If UCase$(Left$(Flxgrd.Text, Len(Key))) = UCase$(Key) Then
'            Flxgrd.Row = cmbInd 'Find and Set the selected combo item
'        End If
'    Next cmbInd
'
'End Sub
Private Sub Flxgrd_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Used for Type Ahead

  Dim cmbInd As Long, lstInd As Long
With Flxgrd
    .Redraw = False
    If KeyCode <> 16 Then
        If DateDiff("s", Now(), StupidTimer) < -1 Then 'Greater than 2 seconds start over with first char
            TypeAheadStor = ""
            FlexGridRowStor = 0
        End If
        StupidTimer = Now()
        TypeAheadStor = TypeAheadStor & Chr$(KeyCode)
        For FlexGridRowStor = (.Rows - 1) To 0 Step -1
            .Row = FlexGridRowStor
            If UCase$(Left$(.Text, Len(TypeAheadStor))) = UCase$(TypeAheadStor) Then
                cmbInd = FlexGridRowStor ' Last Match
            End If
        Next FlexGridRowStor
        If cmbInd Then
            .TopRow = cmbInd
            .Row = cmbInd 'Find and Set the selected item
            .RowSel = cmbInd 'Find and Set the selected item
        End If
    End If
        .Redraw = True
End With
End Sub

'Public Event Click()
Private Sub Flxgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        SortGrid
        CheckboxTrigger Flxgrd.MouseRow, Flxgrd.MouseCol
    End If

End Sub

Public Sub Kilp(CSQLC As Object, Optional AddCol As Integer, Optional Delimeter As String) ' Clip From Recordset in C_SQL class

  Dim RowStore As Integer
  'Debug.Print FLXGRD.Name

    If CSQLC.RS.RecordCount > 0 Then
        If AddCol > 0 Then
            AddColumn CSQLC.RS.Fields(0).Name, Len(CSQLC.RS.Fields(0).Name) * BuildGridInt, AddCol
            Flxgrd.Col = AddCol - 1
            Flxgrd.Rows = CSQLC.RS.RecordCount + 1
          Else
            RowStore = Flxgrd.Rows
            Flxgrd.Rows = CSQLC.RS.RecordCount + RowStore
            Flxgrd.Cols = CSQLC.RS.Fields.Count
            If RowStore > 1 Then
                Flxgrd.Row = RowStore
              Else
                Flxgrd.Row = RowStore
            End If
            Flxgrd.Col = 0
        End If
        Flxgrd.RowSel = Flxgrd.Rows - 1
        Flxgrd.ColSel = Flxgrd.Cols - 1
    End If
    If CSQLC.RS.RecordCount <> 0 Then
        If Delimeter = "" Then
            Delimeter = Chr$(13)
        End If
        Flxgrd.Clip = CSQLC.RS.GetString(adClipString, -1, Chr$(9), Delimeter, vbNullString)
      Else
    End If
    Flxgrd.RowSel = 0
    Flxgrd.ColSel = 0

End Sub

Public Sub KlipAutoSize(CSQLC As Object, Optional GetColumnlegnth As Boolean) ' Clip From Recordset in C_SQL class and auto size the columns base on title lenghts

  Dim Cnt As Integer

    With Flxgrd
        .Clear
        .Rows = 0
        .Cols = 0
        .AllowUserResizing = 1
        If GetColumnlegnth Then
            For Cnt = 0 To CSQLC.RS.Fields.Count - 1
                ResetGridSQL CSQLC.RS.Fields(Cnt).Name, Len(CSQLC.RS.Fields(Cnt).Name) * BuildGridInt
            Next Cnt
        End If
        Kilp CSQLC
    End With

End Sub
Public Function AraToString(ara As Variant) As String
    StrComma As String

End Function
Public Function LoopGrid(AraCells() As String) As Boolean
'exceute inside "do until LoopGrid = false" to loop through a grid without having to think about it.
' Kinda silly function just here to show you how to loop through a flexgrid

  Dim Cnt As Integer
  Static StaticCnt As Integer
  Dim ColCnt As Integer

   
    If Flxgrd.Cols <= 0 Then
        LoopGrid = True
        Exit Function
    End If
     ReDim AraCells(Flxgrd.Cols - 1)
     For Cnt = StaticCnt To Flxgrd.Rows - 1
            Flxgrd.Row = Cnt
        For ColCnt = 1 To Flxgrd.Cols
            Flxgrd.Col = ColCnt - 1
            AraCells(ColCnt - 1) = Flxgrd.Text
            LoopGrid = False
        Next ColCnt
        StaticCnt = StaticCnt + 1
        Exit Function
    Next Cnt
    StaticCnt = 0
    LoopGrid = True

End Function

Public Sub PrintCsv(Filename As String)

    EditSelectAllWithFixed
    Open App.Path & "\" & Filename & ".csv" For Output As #1
    Print #1, Replace(Flxgrd.Clip, vbTab, ",")
    Close #1

End Sub

Public Sub Reset()

    Flxgrd.Visible = False
    Flxgrd.Rows = 0
    Flxgrd.Cols = 0
    Flxgrd.Clear
    Flxgrd.Visible = True

End Sub

Public Sub ResetGridSQL(ColumnName As String, Width As Integer)

  'Rows and Cols must = 0 for this function to work

    With Flxgrd
        If .Cols >= 1 Then
            .Cols = .Cols + 1
        End If
        If .Cols = 0 Then
            .Rows = 2
            .Cols = 2
            .FixedRows = 1
            .FixedCols = 0
            .Rows = 1
            .Cols = 1
        End If
        .Rows = 1
        .Row = 0
        .Col = .Cols - 1
        .Text = ColumnName
        .ColWidth(.Cols - 1) = Width
    End With

End Sub

Private Sub SortGrid()  'Sort the flexgrid according to the column that is selected.

  'FLex Grid Sorting Critera
  'flexSortNone = 0
  'flexSortGenericAscending = 1
  'flexSortGenericDescending = 2
  'flexSortNumericAscending = 3
  'flexSortNumericDescending = 4
  'flexSortStringNoCaseAsending = 5
  'flexSortNoCaseDescending = 6
  'flexSortStringAscending = 7
  'flexSortStringDescending = 8
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Dim selrow As Integer
  Dim selcol As Integer

    If Flxgrd.MouseRow = 0 Then
        If Flxgrd.Rows > 1 Then
            selcol = Flxgrd.MouseCol
            selrow = Flxgrd.MouseRow
            If GrdSortZtoA <> Flxgrd.Text Then
                GrdSortAtoZ = "1"
            End If
            If GrdSortAtoZ = "1" Then
                Flxgrd.Sort = 1
                GrdSortZtoA = Flxgrd.Text
                GrdSortAtoZ = "2"
                Flxgrd.Row = Flxgrd.RowSel
              Else
                Flxgrd.Sort = 2
                GrdSortZtoA = Flxgrd.Text
                GrdSortAtoZ = "1"
                Flxgrd.Row = Flxgrd.RowSel
            End If
            Flxgrd.TopRow = 1
            Flxgrd.ColSel = selcol
            Flxgrd.RowSel = selrow
        End If
    End If

End Sub



':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:36) 27 + 536 = 563 Lines
