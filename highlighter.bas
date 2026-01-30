Attribute VB_Name = "highlighter"
'https://makingyenything.wordpress.com/2025/12/24/chart-highlighters-and-stitch-count-checking-in-excel/
Sub addhighlighter()
'adds vertical and horizontal highlighters that move with selection

'Use with this in the sheet module of the sheet where it will be used
'
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    On Error Resume Next
'    Shapes("horizontalhighlighter").Top = Target.Top
'    Shapes("verticalhighlighter").Left = Target.Left
'End Sub

    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
        0, 0, Columns(10).width * 1200, Rows(10).RowHeight)
        .Name = "horizontalhighlighter"
        'yellow
        .Fill.ForeColor.rgb = rgb(255, 255, 0)
        .Fill.Transparency = 0.8
    End With
    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
        0, 0, Columns(10).width, Rows(10).RowHeight * 1200)
        .Name = "verticalhighlighter"
        'blue
        .Fill.ForeColor.rgb = rgb(0, 0, 255)
        .Fill.Transparency = 0.8
    End With
End Sub


