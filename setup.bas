Attribute VB_Name = "setup"
Sub setrules()
'https://makingyenything.wordpress.com/2025/12/17/faster-code-for-counting-stitch-runs/
'sets conditional formatting rules for that range based on the first table
'in the sheet that is active
    Dim keytable As Range
    Dim chartrange As Range
    Dim i As Integer
    
    Set keytable = ActiveSheet.ListObjects(1).DataBodyRange
    Set chartrange = ActiveSheet.Range("f5:alq178")
    chartrange.Font.Name = keytable.Cells(1, 1).Font.Name
    
    If chartrange.FormatConditions.Count >= keytable.Rows.Count Then
    chartrange.FormatConditions.Delete
    End If
    
    For Each r In keytable.Rows
    
        With chartrange.FormatConditions _
        .Add(xlCellValue, xlEqual, "=" & r.Cells(1, 1).Address)
        With .Interior
        .color = r.Cells(1, 2).Interior.color
        End With
        With .Font
        .color = r.Cells(1, 2).Interior.color
        End With
        End With
    Next r
End Sub
Sub makegridlines()
'https://makingyenything.wordpress.com/2025/12/21/macros-for-adjusting-excel-row-height-column-width-according-to-knitting-gauge-and-adding-major-grid-lines/
    Dim r As Long, c As Long
    Dim chart As Range, wrange As Range
    Dim i As Integer
    
    Set chart = Selection
    
    r = Application.InputBox("Thick line every how many rows?", "Add Gridlines", 5)
    c = Application.InputBox("Thick line every how many columns?", "Add Gridlines", 5)
    
    'set inside borders
    With chart.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
        .color = rgb(100, 100, 100)
    End With
    With chart.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
        .color = rgb(100, 100, 100)
    End With
    'set section borders starting from the right, bottom
    i = chart.Columns.Count - c
    While i > 0
        Set wrange = chart.Columns(i)
        With wrange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = rgb(0, 0, 0)
        End With
        i = i - c
    Wend
    i = chart.Rows.Count - r
    While i > 0
        Set wrange = chart.Rows(i)
        With wrange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = rgb(0, 0, 0)
        End With
        i = i - c
    Wend
    'set border around the whole selection
    chart.BorderAround _
        LineStyle:=xlContinuous, Weight:=xlMedium, color:=rgb(0, 0, 0)
    
End Sub

Sub matchgauge()
'https://makingyenything.wordpress.com/2025/12/21/macros-for-adjusting-excel-row-height-column-width-according-to-knitting-gauge-and-adding-major-grid-lines/
    Dim r As Long, c As Long
    Dim rgauge As Double, cgauge As Double
    Dim w As Double, unit As Double, minimum As Double, dis
    Dim wrange As Range
    Dim rrange As Range

    r = Application.InputBox("How many rows to adjust?", "Match Gauge", 1000)
    c = Application.InputBox("How many columns to adjust?", "Match Gauge", 1000)
    rgauge = Application.InputBox("How many rows per x?", "Match Gauge")
    cgauge = Application.InputBox("How many stitches per x?", "Match Gauge")

    dis = 1
    minimum = 2
    'starts adjusting at column F and row 5
    Set wrange = Columns(6).Resize(, c)
    Set rrange = Rows(5).Resize(r)

    While Abs(dis) > 0.05
        'adjust columns'
        If rgauge > cgauge Then
        'rows smaller than columns
        wrange.ColumnWidth = minimum * (rgauge / cgauge)
        Else
        'columns get the smaller value
        wrange.ColumnWidth = minimum
        End If
        
        'convert to points
        w = wrange.width / c
        'adjust rows
        rrange.RowHeight = w * (cgauge / rgauge)
    
        'check if ratio correct
        dis = Columns(6).width * cgauge - Rows(5).RowHeight * rgauge
        minimum = minimum + 0.1
    Wend

    MsgBox "Finished adjusting cells, error of " & Round(dis, 3) & " points"
    
        
End Sub
