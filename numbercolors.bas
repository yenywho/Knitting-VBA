Attribute VB_Name = "numbercolors"
'https://makingyenything.wordpress.com/2025/12/17/faster-code-for-counting-stitch-runs/
Sub color()
' used to copy charts made from conditional formatting
' so number() can be used after
Set ref_sheet = Worksheets("final_design")
Dim c As Range
For Each c In ref_sheet.Range("f5:lp178").Cells
If c.DisplayFormat.Interior.color <> rgb(255, 255, 255) Then

ActiveSheet.Range(c.Address).Interior.color = c.DisplayFormat.Interior.color
End If
Next c

End Sub
Sub number()

Dim box As Range
Dim c As Range
Dim m As Range

Dim color As Long

Dim s As Double
Dim a As String
Dim b As String
'Application.ScreenUpdating = False
' numbers runs of colors with it's length in the range below
Set box = Range("f5:lp178")


With box

For i = 1 To .Rows.Count
    a = .Cells(i, 1).Address
    s = .Cells(i, 1).DisplayFormat.Interior.color
    For j = 1 To .Columns.Count
        color = .Cells(i, j).Interior.color
        If color <> s Then
            b = .Cells(i, j - 1).Address
            If s = rgb(255, 255, 255) Then
                a = b
                s = color
            Else
                Set m = Range(a & ":" & b)
                With m
                If .Count > 1 Then
                    '.Merge
                    '.HorizontalAlignment = xlCenter
                    '.Value = .Count
                    .Cells(1, .Count / 2).Value = .Count
                    If is_dark(.DisplayFormat.Interior.color) Then
                        .Font.color = rgb(255, 255, 255)
                    End If
                End If
                End With
            End If
        
            a = .Cells(i, j).Address
            s = color
        End If
    
    Next
    
Next
End With
Application.ScreenUpdating = True
End Sub

Function is_dark(i As Double) As Boolean
Dim r As Double
Dim g As Double
Dim b As Double

r = i Mod 256
g = (i \ 256) Mod 256
b = i \ 65536

is_dark = 0.2126 * r + 0.7152 * g + 0.0722 * b < 128
End Function

