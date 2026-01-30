Attribute VB_Name = "charttransform"

Function flip_vert(arr As Variant) As Variant
Dim out As Variant
Dim i As Integer, j As Integer
Dim t As Variant

If TypeName(arr) = "Range" Then
out = arr.Value
Else
out = arr
End If

n = UBound(out, 1) - LBound(out, 1) + 1
m = UBound(out, 2) - LBound(out, 2) + 1

For i = 1 To n / 2
    For j = 1 To m
        t = out(i, j)
        out(i, j) = out(n - i + 1, j)
        out(n - i + 1, j) = t
    Next j
Next i
flip_vert = out
End Function
Function flip_horiz(arr As Variant) As Variant
Dim out As Variant
Dim i As Integer, j As Integer
If TypeName(arr) = "Range" Then
    out = arr.Value
Else
    out = arr
End If
m = UBound(out, 1) - LBound(out, 1) + 1
n = UBound(out, 2) - LBound(out, 2) + 1
Dim t As Variant
For i = 1 To n / 2
    For j = 1 To m
        t = out(j, i)
        out(j, i) = out(j, n - i + 1)
        out(j, n - i + 1) = t
    Next j
Next i
flip_horiz = out
End Function

Function rotate_ccw(arr As Variant) As Variant
Dim out As Variant
out = Application.Transpose(flip_horiz(arr))
rotate_ccw = out
End Function
Function rotate_cw(arr As Variant) As Variant
Dim out As Variant
out = Application.Transpose(flip_vert(arr))
rotate_cw = out
End Function


