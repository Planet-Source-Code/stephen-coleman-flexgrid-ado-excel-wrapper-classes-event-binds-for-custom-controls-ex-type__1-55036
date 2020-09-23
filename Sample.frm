VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "How to use shared classes. Excel, Flexgrids, Error logging, Ado connectivity."
   ClientHeight    =   7440
   ClientLeft      =   2805
   ClientTop       =   3840
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   13545
   Begin VB.CommandButton Command1 
      Caption         =   "Readme.txt"
      Height          =   375
      Index           =   10
      Left            =   3360
      TabIndex        =   18
      ToolTipText     =   "Open the readme.txt to hear me ramble on."
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox TXTColor 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Text            =   "35"
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox TXTColor 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Text            =   "10"
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox TXTColor 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Text            =   "50"
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Sample.frx":0000
      Top             =   5280
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Colors"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Wow change the colors!"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dynamic Grid"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Adds a new flexgrid to the from dynamically!!!"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print in Excel"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Copies the flexgrid data from the Flexgrid to Excel and formats it!"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dump Grid To File"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Makes a csv file from the clip control of the flexgrid"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clip Sample Data"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Copies data from MDB to the flexgrid using clip.. Very Fast"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox TXTDatabase 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Hmmmm. Don't click this. I don't know what it does! AHHHHHHH!!!!"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kill Dynamic Grid"
      Height          =   375
      Index           =   8
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Removes the dynamic flexgrid"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Terminate Excel"
      Height          =   375
      Index           =   7
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Terminates the excel object. Leave the created excel spread sheet open to see the results."
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Grid"
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Wipes the grid"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add CheckBox"
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   4
      ToolTipText     =   "Very handy check box function!"
      Top             =   5280
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4455
      Left            =   0
      TabIndex        =   17
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   3
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LoopGrid"
      Height          =   375
      Index           =   11
      Left            =   3360
      TabIndex        =   19
      ToolTipText     =   "A stupid or... is it unique way to loop through the flexgrid."
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Project: OOP_Shared_Classes (c)  Serious Solutions
'
' Programmer:  Stephen S. Coleman
'
' Email:  Steve@Webnexis.com
'
' Date:  July 16, 2004
'
' Comments:
' I setup this Example to show how Classes designed for shared Access(By multiple forms or applications)
'can really make creating new apps or just another form with similar controls easy.
'
' Modifications:  SSC - Created at 1:26 PM
'-------------------------------------------------------------------------------
Option Explicit
'Requires Public ErrC As New C_Error
Private Grd As New C_FlexGrid 'Brings the grid class to life.  Notice that GRD.Flxgrd. allows to access all the properties of the flexgrid
Private C As New C_SQL 'This can be Public as it is for a persistent connection to a Database or set in a sub or function for a single query.
Private XL As New C_Excel ' Excel handler
'When the sub ends the class and all memory allocated is terminated and you don't have to manage it.
'"SET C = Nothing" will terminate the class.

Private Sub Command1_Click(Index As Integer)

  Dim query As String
Dim AraGrdTx() As String
    On Error GoTo Error_Handler '--Error Trap------------------------------------------------------
    If Err.Number Then
Error_Handler:
        ErrC.ErrorTrap Me.Name, "Command1_Click", Err.Number, Err.Description, Err.Source
        Resume Next
    End If '--Error Trap---------------------------------------------------------------------------

    Select Case Index
      Case Is = 0
        C.DataBase = TXTDatabase
        query = "SELECT  DISTINCTROW  [Order Details].OrderID, [Order Details].ProductID, Products.ProductName, [Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount, CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS ExtendedPrice FROM Products INNER JOIN [Order Details] ON Products.ProductID = [Order Details].ProductID ORDER BY [Order Details].OrderID;"
        'query = "SELECT  DISTINCTROW  Products.ProductName FROM Products INNER JOIN [Order Details] ON Products.ProductID = [Order Details].ProductID ORDER BY [Order Details].OrderID;"

        If C.SendQuery(query) Then
            Grd.KlipAutoSize C, True
            L(1) = "Clip " & C.RS.RecordCount
          Else 'C.SendQuery(query) = False / 0
            MsgBox "Query Failed Verify your settings", vbCritical
        End If
      Case Is = 1
        Grd.dumptofile App.Path, "Dump.CSV", True
      Case Is = 2
        XL.ClipToXl Grd, True, "Test report"
      Case Is = 3
        Grd.DynamicAddToform Me, 4575, 6975, 600, 7080, 1  'Adds dynamic msflexgrid to form as GRD1, Terminates link to current msflexgrid
        'Sets dynamic grid on form and now everything works off the new msflexgrid
      Case Is = 4
        Grd.Colors TXTColor(0), TXTColor(1), TXTColor(2)   ' Sets Colors on The Grid
      Case Is = 5
        Grd.CheckBox InputBox("Enter a column number")
      Case Is = 6
        Grd.Reset
      Case Is = 7
        Set XL = Nothing ' See that XL is killed and so is the visible spread sheet.
      Case Is = 8
        Grd.DynamicRemove Me, "Grd1"
        Set Grd.Flxgrd = MSFlexGrid1 ' Relates the grid on the form to the C_flexgrid Class as Grd
      Case Is = 9
        ErrC.UnloadAll
        Case Is = 10
        Shell "notepad.exe " & App.Path & "\ReadMe.txt", vbMaximizedFocus
        Case Is = 11 ' loopgrid function
            Text1 = ""
            Do While Grd.LoopGrid(AraGrdTx) = False
                Text1 = Text1 & AraToComma(AraGrdTx) & vbCrLf
                Text1.SelStart = Len(Text1)
            Loop
    End Select

End Sub

Private Sub Form_Load()

    Set Grd.Flxgrd = MSFlexGrid1 ' Relates the grid on the form to the C_flexgrid Class as Grd
    TXTDatabase = App.Path & "\Nwind.mdb"
    ErrC.ErrResume = True
    ErrC.ErrPost = True

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:36) 22 + 56 = 78 Lines
