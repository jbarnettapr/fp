Attribute VB_Name = "Test"
Option Explicit

'@Test
Public Function Try() As Boolean
    Dim expected
    expected = Add(6, 5.5)
    Debug.Print expected
End Function

Public Function Add(A, B)
    Add = Dispatch("Add", A, B)
End Function

Public Function Add_Variant_Variant(A, B)
    Add_Variant_Variant = A + B
End Function

Public Function Add_Integer_Integer(A As Integer, B As Integer) As Integer
    Add_Integer_Integer = A + B
End Function

Public Function Add_Double(A As Double, B) As Double
    Add_Double = A + B
End Function
