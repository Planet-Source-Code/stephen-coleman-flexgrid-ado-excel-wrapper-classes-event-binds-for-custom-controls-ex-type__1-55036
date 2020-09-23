Attribute VB_Name = "M_Global"
Option Explicit
'errc as public so all the classes and form can use it for error handling.
' I counld have put it as private in every class and the form
'but there is no need to have many instances of this class
'one is all you need for error handling.
Public ErrC As New c_error

Public Function AraToComma(ara() As String) As String
Dim Cnt As Integer
        
        For Cnt = 0 To (UBound(ara) + 1) - 3
            AraToComma = AraToComma & ara(Cnt) & ","
        Next Cnt
        AraToComma = AraToComma & ara(Cnt + 1)
End Function
