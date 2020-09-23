Attribute VB_Name = "modHex"
Option Explicit
Public Function ArcSin(X As Double) As Double
    On Error GoTo Error_ArcSin
    ArcSin = Atn(X / Sqr(-X * X + 1))
    Exit Function
Error_ArcSin:
    Debug.Print "Error in ArcSin(" & X & ")"
End Function

Public Function ArcCos(X As Double) As Double
    On Error GoTo Error_ArcCos
    ArcCos = ArcSin(X) + 2 * Atn(1)
    Exit Function
Error_ArcCos:
    Debug.Print "Error in ArcCos(" & X & ")"
End Function

Public Function RadToDeg(r As Double) As Double
    RadToDeg = r * 180 / pi
End Function

Public Function DegToRad(d As Double) As Double
    DegToRad = pi * d / 180
End Function

Public Function pi() As Double
    pi = Atn(1) * 4
End Function


