VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' C_Stack%.cls
'
' Programmer:  Stephen S. Coleman
'
' Date: July 19, 2004
'
' Comments: I use this class for storage for many of my other classes.
'
' Description: 'Very Useful Stack class I found don't remember who to give credit to.
'
' Modifications:  SSC - Created at 10:25 AM
'-------------------------------------------------------------------------------
Option Explicit
Private Const BUFFER_SIZE As Long = 100
Private Items() As Variant
Private Index As Long

Private Sub Class_Initialize()

    Clear

End Sub

Private Sub Class_Terminate()

    Erase Items

End Sub

Public Sub Clear()

    ReDim Items(BUFFER_SIZE - 1)
    Index = 0

End Sub

Public Property Get Count() As Long

    Count = Index

End Property

Public Function Peek() As Variant

    If Index > 0 Then
        Peek = Items(Index - 1)
        If IsObject(Items(Index - 1)) Then
            Set Peek = Items(Index - 1)
          Else
            Peek = Items(Index - 1)
        End If
    End If

End Function

Public Function Pop() As Variant

  Dim L As Long

    If Index > 0 Then
        Index = Index - 1
        If IsObject(Items(Index)) Then
            Set Pop = Items(Index)
          Else
            IsObject (Items(Index))
            Pop = Items(Index)
        End If
        '   Else: Clear
    End If

    If (UBound(Items) - Index) \ BUFFER_SIZE > 0 Then
        ReDim Preserve Items(UBound(Items) - BUFFER_SIZE)
    End If

End Function

Public Sub Push(Item As Variant)

    If UBound(Items) = Index Then
        ReDim Preserve Items(Index + BUFFER_SIZE)
    End If

    If IsObject(Item) Then
        Set Items(Index) = Item
      Else
        IsObject (Item)
        Items(Index) = Item
    End If
    Index = Index + 1

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-16 14:36) 5 + 77 = 82 Lines
