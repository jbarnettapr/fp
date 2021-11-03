Attribute VB_Name = "Tools"
Option Explicit

Public Function Dispatch(fn, ParamArray args())
    Dim Types() As String
    Dim fwd, arg, i As Byte, signature As String
    
    fwd = args
    ReDim Types(1 To 1 + UBound(fwd) - LBound(fwd))
    
    For Each arg In fwd
        i = i + 1
        Types(i) = TypeName(arg)
    Next
    
    signature = fn & "_" & Join(Types, "_")
    On Error GoTo Catch
    Dispatch = Forward(signature, fwd)
    Exit Function
Catch:
    Debug.Print "ERROR: No such function " & signature
    
End Function


Public Function Forward(fn, args)
    Select Case UBound(args) - LBound(args)
        Case -1
            Forward = Application.Run(fn)
        Case 0
            Forward = Application.Run(fn, args(LBound(args)))
        Case 1
            Forward = Application.Run(fn, args(LBound(args)), args(LBound(args) + 1))
        Case 2
            Forward = Application.Run(fn, args(LBound(args)), args(LBound(args) + 1), args(LBound(args) + 2))
        Case 3
            Forward = Application.Run(fn, args(LBound(args)), args(LBound(args) + 1), args(LBound(args)) + 2, args(LBound(args) + 3))
        Case 4
            Forward = Application.Run(fn, args(LBound(args)), args(LBound(args) + 1), args(LBound(args)) + 2, args(LBound(args) + 3), args(LBound(args) + 4))
        Case 5
            Forward = Application.Run(fn, args(LBound(args)), args(LBound(args) + 1), args(LBound(args)) + 2, args(LBound(args) + 3), args(LBound(args) + 4), args(LBound(args) + 5))
    End Select
End Function
