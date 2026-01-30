Attribute VB_Name = "writtentochart"
'https://makingyenything.wordpress.com/2025/12/30/written-knitting-pattern-to-chart/
Function convertwrittentochart(s As String, key As Range, Optional stitches As Range, _
    Optional rs As Boolean = True, Optional countcolumn As Range, Optional stitchcount As Integer)
If s = "" Then
GoTo done
End If

Dim outputstring As String
Dim outputarr As Variant
Dim x As Variant
Dim tempstring1 As String
Dim tempstring2 As String
Dim keystring As String
  
Dim regexOne As Object
Dim regexTwo As Object
Dim regexThree As Object
Dim theMatches As Object
Dim theMatchesA As Object
Dim theMatchesB As Object
Dim theMatchesC As Object
Dim Match1 As Object
Dim Match2 As Object
  
Dim stitchtionary1 As Object
Dim stitchtionary2 As Object
Dim i As Integer, k As Integer, j As Integer, m As Integer
Set regexOne = New RegExp
Set regexTwo = New RegExp
Set regexThree = New RegExp

regexOne.Global = True
regexOne.IgnoreCase = True
regexTwo.Global = True
regexTwo.IgnoreCase = True
regexThree.Global = True
regexThree.IgnoreCase = True

'key - char with ! to represent extra spaces for cables
Set stitchtionary1 = CreateObject("Scripting.Dictionary")
'char - stitch
Set stitchtionary2 = CreateObject("Scripting.Dictionary")
If LCase(Left(s, 2)) = "ws" Or LCase(Left(s, 5)) = "wrong" Then
    rs = False
End If
outputstring = s
If stitches Is Nothing Then
Set stitches = key
End If
m = key.Cells.Count
j = 65

For i = 1 To m
    If key.Cells(i).Value <> "" Then
    tempstring1 = Chr(j)
    'add stitches for cables
    If Not (countcolumn Is Nothing) Then
         
        If countcolumn.Cells(i).Value > 1 Then
            For k = 1 To countcolumn.Cells(i).Value - 1
            If rs Then
                tempstring1 = "!" & tempstring1
            Else
                tempstring1 = tempstring1 & "!"
            End If
                 
            Next
        Else
        End If
    End If
    tempstring2 = Replace(LCase(key.Cells(i).Value), "1 ", "_")
    stitchtionary1.Add tempstring2, tempstring1
    stitchtionary2.Add Chr(j), stitches.Cells(i).Value
    j = j + 1
    If j = 91 Then
    j = 97
    End If
    End If
     
Next i
'replace ks stuck to ps
regexOne.Pattern = "\b[KkPp]{2,}\b"
Set theMatches = regexOne.Execute(outputstring)
For Each Match1 In theMatches
    tempstring1 = Replace(Match1, "k", "k ")
    tempstring1 = Replace(tempstring1, "K", "K ")
    tempstring1 = Replace(tempstring1, "p", "p ")
    tempstring1 = Replace(tempstring1, "P", "P ")
    outputstring = Replace(outputstring, Match1, tempstring1)
Next

'prepare regex made from key
keystring = Join(stitchtionary1.keys, "|")
'for things like k1 tbl k2 tbl s2 wyif etc
regexOne.Pattern = "\b[^|]+_[^|]+\b"

Set theMatches = regexOne.Execute(keystring)

If theMatches.Count > 0 Then
    For i = 0 To theMatches.Count - 1
        keystring2 = "\b" & Replace(theMatches(i), "_", "\s?(\d*)\s?") & "\b"
        regexTwo.Pattern = keystring2
        Set theMatchesA = regexTwo.Execute(outputstring)
        k = 0
        For Each Match1 In theMatchesA
            If Match1.SubMatches(0) = "" Then
                tempstring2 = theMatches(i)
            Else
                tempstring2 = WorksheetFunction.Rept(theMatches(i) & " ", Match1.SubMatches(0))
            End If
            tempstring1 = Replace(outputstring, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
            outputstring = Left(outputstring, Match1.FirstIndex + k) & tempstring1
            k = k + Len(tempstring2) - Len(Match1)
        Next
    Next i
End If

'find () or [] and the first number after assuming repeats
regexOne.Pattern = "(?:\[|\()([^\(\)\]\[]+)(?:\)|\]).*?(\d+)"
'do 2 rounds of replacing to account for nested groups
Set theMatches = regexOne.Execute(outputstring)
k = 0
For Each Match1 In theMatches
    tempstring2 = WorksheetFunction.Rept(Match1.SubMatches(0) & " ", Match1.SubMatches(1))
    tempstring1 = Replace(outputstring, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
    outputstring = Left(outputstring, Match1.FirstIndex + k) & tempstring1
    k = k + Len(tempstring2) - Len(Match1)
Next
Set theMatches = regexOne.Execute(outputstring)
k = 0
For Each Match1 In theMatches
    tempstring2 = WorksheetFunction.Rept(Match1.SubMatches(0) & " ", Match1.SubMatches(1))
    tempstring1 = Replace(outputstring, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
    outputstring = Left(outputstring, Match1.FirstIndex + k) & tempstring1
    k = k + Len(tempstring2) - Len(Match1)
Next




'finds keys with number with 1 or 0 spaces between them
'within word boundary
regexOne.Pattern = "\b(" & keystring & ")\s?(\d+)\b"
'convertwrittentochart = regexOne.Pattern

Set theMatches = regexOne.Execute(outputstring)
k = 0
For Each Match1 In theMatches
    tempstring2 = WorksheetFunction.Rept(Match1.SubMatches(0) & " ", Match1.SubMatches(1))
    tempstring1 = Replace(outputstring, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
    outputstring = Left(outputstring, Match1.FirstIndex + k) & tempstring1
    k = k + Len(tempstring2) - Len(Match1)
Next
Set theMatches = regexTwo.Execute(outputstring)
'find stitchcount if exists
If stitchcount = 0 Then
    regexOne.Pattern = "\(\b(\d+)\b\ssts\.\)"
    Set theMatches = regexOne.Execute(outputstring)
    If theMatches.Count > 0 Then
        stitchcount = theMatches(0).SubMatches(0)
    End If
End If
'catches stitches in the key
regexThree.Pattern = "\b(" & keystring & ")\b"

'deal with *blah repeat from * pattern
regexTwo.Pattern = "(.*)\*(.+)rep.+\*(?:.*\b(\d+)\b\stimes|(.*)\(\b(\d+)\b\ssts.\)|.*end|)(.*)"

'group 1 is part before *
'group 2 is after * and before rep
'group 3 is number after rep and before times if it exists
'group 4 is stitches after rep and before stitch count if it exists
'group 5 is stitch count if it exists
'group 6 is stitches after rep amount or after end
Set theMatches = regexTwo.Execute(outputstring)

If theMatches.Count = 0 Then
    Set theMatchesA = regexThree.Execute(outputstring)
    outputstring = ""
    For Each Match1 In theMatchesA
        'annoyingly, matches are not type string and will cause an error
        tempstring1 = Match1
        outputstring = outputstring & stitchtionary1(LCase(tempstring1))
    Next
    If stitchcount <> 0 Then
        If Len(outputstring) < stitchcount Then
        k = WorksheetFunction.RoundUp(stitchcount / Len(outputstring), 0)
        outputstring = WorksheetFunction.Rept(outputstring, k)
        outputstring = Left(outputstring, stitchcount)
         
        End If
    End If
     
Else
With theMatches(0)
    m = 0
    Set theMatchesA = regexThree.Execute(.SubMatches(0))
    Set theMatchesB = regexThree.Execute(.SubMatches(1))
    Set theMatchesC = regexThree.Execute(.SubMatches(3))
    If theMatchesC.Count = 0 Then
        Set theMatchesC = regexThree.Execute(.SubMatches(5))
    End If
     
    If .SubMatches(2) = "" Then
        If .SubMatches(4) = "" Then
            If stitchcount = 0 Then
                m = theMatchesA.Count + theMatchesC.Count + theMatchesB.Count
            Else
                m = stitchcount
            End If
        Else
        m = .SubMatches(4)
        End If
        k = WorksheetFunction.RoundUp((m - theMatchesA.Count - theMatchesC.Count) / theMatchesB.Count, 0)
    Else
        k = .SubMatches(2)
    End If
    outputstring = ""
    For Each Match1 In theMatchesA
        tempstring1 = Match1
        outputstring = outputstring & stitchtionary1(LCase(tempstring1))
    Next
    For j = 1 To k
        For Each Match1 In theMatchesB
            tempstring1 = Match1
            outputstring = outputstring & stitchtionary1(LCase(tempstring1))
        Next
    Next
    For Each Match1 In theMatchesC
        tempstring1 = Match1
        outputstring = outputstring & stitchtionary1(LCase(tempstring1))
    Next
    If m <> 0 Then
    outputstring = Left(outputstring, m)
    End If
End With
End If
'reverse if rs
If rs Then
    outputstring = StrReverse(outputstring)
End If
ReDim outputarr(1 To Len(outputstring))
'fill array for output
For i = 1 To Len(outputstring)
    tempstring1 = Mid$(outputstring, i, 1)
     
    If tempstring1 = "!" Then
        outputarr(i) = vbNullString
    Else
        outputarr(i) = stitchtionary2(tempstring1)
    End If
Next
done:
convertwrittentochart = outputarr
End Function
Function convertwrittentochartwcolor(s As String, key As Range, Optional stitches As Range, _
    Optional rs As Boolean = True, Optional countcolumn As Range, Optional stitchcount As Integer, Optional colorkey As Range, Optional colorcol As Range)
If s = "" Then
GoTo done
End If
Dim loc As String
Dim alteredstr As String
Dim outputstr As String
Dim outputarr As Variant
Dim x As Variant
Dim tempstring1 As String
Dim tempstring2 As String
Dim keystring As String
Dim colorsout As New Collection
  
Dim regexOne As Object
Dim regexTwo As Object
Dim regexThree As Object
Dim theMatches As Object
Dim theMatchesA As Object
Dim theMatchesB As Object
Dim theMatchesC As Object
Dim colorMatches As Object
Dim Match1 As Object
Dim Match2 As Object
  
Dim stitchtionary1 As Object
Dim stitchtionary2 As Object
Dim colordic As Object
Dim i As Integer, k As Integer, j As Integer, m As Integer
Set regexOne = New RegExp
Set regexTwo = New RegExp
Set regexThree = New RegExp

loc = ActiveCell.Address

regexOne.Global = True
regexOne.IgnoreCase = True
regexTwo.Global = True
regexTwo.IgnoreCase = True
regexThree.Global = True
regexThree.IgnoreCase = True

'key - char with ! to represent extra spaces for cables
Set stitchtionary1 = CreateObject("Scripting.Dictionary")
'char - stitch
Set stitchtionary2 = CreateObject("Scripting.Dictionary")
'for colors
Set colordict = CreateObject("Scripting.Dictionary")

If LCase(Left(s, 2)) = "ws" Or LCase(Left(s, 5)) = "wrong" Then
    rs = False
End If
alteredstr = s
If stitches Is Nothing Then
Set stitches = key
End If
m = key.Cells.Count
j = 65

For i = 1 To m
    If key.Cells(i).Value <> "" Then
    tempstring1 = Chr(j)
    'add stitches for cables
    If Not (countcolumn Is Nothing) Then
         
        If countcolumn.Cells(i).Value > 1 Then
            For k = 1 To countcolumn.Cells(i).Value - 1
            If rs Then
                tempstring1 = "!" & tempstring1
            Else
                tempstring1 = tempstring1 & "!"
            End If
                 
            Next
        Else
        End If
    End If
    tempstring2 = Replace(LCase(key.Cells(i).Value), "1 ", "_")
    stitchtionary1.Add tempstring2, tempstring1
    stitchtionary2.Add Chr(j), stitches.Cells(i).Value
    j = j + 1
    If j = 91 Then
    j = 97
    End If
    End If
     
Next i


'replace ks stuck to ps
regexOne.Pattern = "\b[KkPp]{2,}\b"
Set theMatches = regexOne.Execute(alteredstr)
For Each Match1 In theMatches
    tempstring1 = Replace(Match1, "k", "k ")
    tempstring1 = Replace(tempstring1, "K", "K ")
    tempstring1 = Replace(tempstring1, "p", "p ")
    tempstring1 = Replace(tempstring1, "P", "P ")
    alteredstr = Replace(alteredstr, Match1, tempstring1)
Next

'prepare regex made from key
keystring = Join(stitchtionary1.keys, "|")

'for things like k1 tbl k2 tbl s2 wyif etc
regexOne.Pattern = "\b[^|]+_[^|]+\b"

Set theMatches = regexOne.Execute(keystring)

If theMatches.Count > 0 Then
    For i = 0 To theMatches.Count - 1
        keystring2 = "\b" & Replace(theMatches(i), "_", "\s?(\d*)\s?") & "\b"
        regexTwo.Pattern = keystring2
        Set theMatchesA = regexTwo.Execute(alteredstr)
        k = 0
        For Each Match1 In theMatchesA
            If Match1.SubMatches(0) = "" Then
                tempstring2 = theMatches(i)
            Else
                tempstring2 = WorksheetFunction.Rept(theMatches(i) & " ", Match1.SubMatches(0))
            End If
            tempstring1 = Replace(alteredstr, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
            alteredstr = Left(alteredstr, Match1.FirstIndex + k) & tempstring1
            k = k + Len(tempstring2) - Len(Match1)
        Next
    Next i
End If

'find () or [] and the first number after assuming repeats
regexOne.Pattern = "(?:\[|\()([^\(\)\]\[]+)(?:\)|\]).*?(\d+)"
'do 2 rounds of replacing to account for nested groups
Set theMatches = regexOne.Execute(alteredstr)
k = 0
For Each Match1 In theMatches
    tempstring2 = WorksheetFunction.Rept(Match1.SubMatches(0) & " ", Match1.SubMatches(1))
    tempstring1 = Replace(alteredstr, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
    alteredstr = Left(alteredstr, Match1.FirstIndex + k) & tempstring1
    k = k + Len(tempstring2) - Len(Match1)
Next
Set theMatches = regexOne.Execute(alteredstr)
k = 0
For Each Match1 In theMatches
    tempstring2 = WorksheetFunction.Rept(Match1.SubMatches(0) & " ", Match1.SubMatches(1))
    tempstring1 = Replace(alteredstr, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
    alteredstr = Left(alteredstr, Match1.FirstIndex + k) & tempstring1
    k = k + Len(tempstring2) - Len(Match1)
Next




'finds keys with number with 1 or 0 spaces between them
'within word boundary
regexOne.Pattern = "\b(" & keystring & ")\s?(\d+)\b"
'convertwrittentochart = regexOne.Pattern

Set theMatches = regexOne.Execute(alteredstr)
k = 0
For Each Match1 In theMatches
    tempstring2 = WorksheetFunction.Rept(Match1.SubMatches(0) & " ", Match1.SubMatches(1))
    tempstring1 = Replace(alteredstr, Match1, tempstring2, Match1.FirstIndex + k + 1, 1)
    alteredstr = Left(alteredstr, Match1.FirstIndex + k) & tempstring1
    k = k + Len(tempstring2) - Len(Match1)
Next
Set theMatches = regexTwo.Execute(alteredstr)
'find stitchcount if exists
If stitchcount = 0 Then
    regexOne.Pattern = "\(\b(\d+)\b\ssts\.\)"
    Set theMatches = regexOne.Execute(alteredstr)
    If theMatches.Count > 0 Then
        stitchcount = theMatches(0).SubMatches(0)
    End If
End If
'catches stitches in the key
regexThree.Pattern = "\b(" & keystring & ")\b"

'deal with *blah repeat from * pattern
regexTwo.Pattern = "(.*)\*(.+)rep.+\*(?:.*\b(\d+)\b\stimes|(.*)\(\b(\d+)\b\ssts.\)|.*end|)(.*)"


outputstr = ""
Set theMatches = regexTwo.Execute(alteredstr)
'group 1 is part before *
'group 2 is after * and before rep
'group 3 is number after rep and before times if it exists
'group 4 is stitches after rep and before stitch count if it exists
'group 5 is stitch count if it exists
'group 6 is stitches after rep amount or after end

If theMatches.Count > 0 Then
With theMatches(0)
    m = 0
    Set theMatchesA = regexThree.Execute(.SubMatches(0))
    Set theMatchesB = regexThree.Execute(.SubMatches(1))
    Set theMatchesC = regexThree.Execute(.SubMatches(3))
    tempstring1 = .SubMatches(3)
    If theMatchesC.Count = 0 Then
        Set theMatchesC = regexThree.Execute(.SubMatches(5))
        tempstring = .SubMatches(5)
    End If
    k = 1
    If .SubMatches(2) = "" Then
        If .SubMatches(4) = "" Then
            If stitchcount = 0 Then
                m = theMatchesA.Count + theMatchesC.Count + theMatchesB.Count
            Else
                m = stitchcount
            End If
        Else
        m = .SubMatches(4)
        End If
        k = WorksheetFunction.RoundUp((m - theMatchesA.Count - theMatchesC.Count) / theMatchesB.Count, 0)
    Else
        k = .SubMatches(2)
    End If
    
    tempstring2 = WorksheetFunction.Rept(.SubMatches(1) & " ", k)
    
    alteredstr = .SubMatches(0) & tempstring2 & tempstring1

End With
End If

If colorkey Is Nothing Then

    Set theMatchesA = regexThree.Execute(alteredstr)

    For Each Match1 In theMatchesA
        'annoyingly, matches are not type string and will cause an error
        tempstring1 = Match1
        outputstr = outputstr & stitchtionary1(LCase(tempstring1))
    Next
    If stitchcount <> 0 Then
        If Len(alteredstr) < stitchcount Then
        k = WorksheetFunction.RoundUp(stitchcount / Len(outputstr), 0)
        outputstr = WorksheetFunction.Rept(outputstr, k)
        outputstr = Left(outputstr, stitchcount)
         
        End If
    End If
Else
    m = colorkey.Cells.Count
    For i = 1 To m
        colordict.Add colorkey.Cells(i).Value, colorcol.Cells(i).Interior.Color
    Next i
    colorstring = Join(colordict.keys, "|")
    'get with {color} sections
    regexOne.Pattern = "with\s(" & colorstring & ")(.*?)(?=with|$)"
    
    Set colorMatches = regexOne.Execute(alteredstr)
        For Each Match2 In colorMatches
            tempstring2 = ""
    
                Set theMatchesA = regexThree.Execute(Match2.SubMatches(1))
            
                For Each Match1 In theMatchesA
                    tempstring1 = Match1
                    tempstring2 = tempstring2 & stitchtionary1(LCase(tempstring1))
                Next
                If stitchcount <> 0 And colorMatches.Count = 1 Then
                    If Len(tempstring2) < stitchcount Then
                    k = WorksheetFunction.RoundUp(stitchcount / Len(tempstring2), 0)
                    tempstring2 = WorksheetFunction.Rept(tempstring2, k)
                    tempstring2 = Left(tempstring2, stitchcount)
                    
                    End If
                End If
    
            For i = 1 To Len(tempstring2)
                colorsout.Add colordict(Match2.SubMatches(0))
            Next
            outputstr = outputstr & tempstring2
        Next
    
    m = colorsout.Count
    
    For i = 1 To m
        'reverse if rs
        If rs Then
            k = m - i + 1
        Else
            k = i
        End If
        Evaluate ("colorcell(" & Application.Caller.Offset(0, i - 1).Address & "," & colorsout(k) & ")")
        
    Next

End If

'reverse if rs
If rs Then
    outputstr = StrReverse(outputstr)
End If
ReDim outputarr(1 To Len(outputstr))
'fill array for output
For i = 1 To Len(outputstr)
    tempstring1 = Mid$(outputstr, i, 1)
     
    If tempstring1 = "!" Then
        outputarr(i) = vbNullString
    Else
        outputarr(i) = stitchtionary2(tempstring1)
    End If
Next


done:
convertwrittentochartwcolor = outputarr
End Function



Function organizepattern(s As String, Optional delimiter As String = ":", Optional number_sts As String = "")
Dim regex_one As Object
Set regex_one = New RegExp
Dim regex_two As Object
Set regex_two = New RegExp
Dim regex_three As Object
Set regex_three = New RegExp

Dim r_arr() As Variant
Dim t_arr As Variant
Dim out_arr As Variant
Dim i As Integer, n As Integer, k As Integer, m As Integer
Dim max_row As Integer
Dim x As Integer
Dim tempstring1 As String
Dim tempstring2 As String
Dim matches0 As Object
Dim matches1 As Object
Dim matches2 As Object
Dim match As Object

regex_one.Global = True
regex_one.IgnoreCase = True
regex_two.Global = True
regex_two.IgnoreCase = True
regex_three.Global = True
regex_three.IgnoreCase = True

s = WorksheetFunction.Trim(s)
'common replacements
s = Replace(s, "knit", "k")
s = Replace(s, "purl", "p")
s = Replace(s, "Knit", "k")
s = Replace(s, "Purl", "p")
s = Replace(s, "once", "1")
s = Replace(s, "twice", "2")

'parses a through b a-b
regex_three.Pattern = "(\d+)(?:\sthrough\s|-)(\d+)"
Set matches1 = regex_three.Execute(s)
k = 0

For Each match In matches1
    tempstring1 = ""
    For i = match.SubMatches(0) To match.SubMatches(1)
        tempstring1 = tempstring1 & i & " "
    Next
    s = Left(s, match.FirstIndex + k) & Replace(s, match, tempstring1, match.FirstIndex + k + 1, 1)
    k = k + Len(tempstring1) - Len(match)
Next

'rewrites this pattern K3, * p4, k4; rep from *, end last rep k1 instead of k4
regex_three.Pattern = "(\*([^\*.]+\b)[^\*.]+rep[^\*.]+\*)[^\*.]+end\slast\srep[a-z\s]*(\b[^.]+)\sinstead\sof\s([^.]+)\."
k = 0
Set matches1 = regex_three.Execute(s)
For Each match In matches1
    tempstring1 = match.SubMatches(0) & Replace(match.SubMatches(1), match.SubMatches(3), match.SubMatches(2)) & "."
    s = Left(s, match.FirstIndex + k) & Replace(s, match, tempstring1, match.FirstIndex + k + 1, 1)
    k = k + Len(tempstring1) - Len(match)
Next

'finds rs ws
regex_two.Pattern = "\b(wrong|ws|right|rs)\b"

'finds row numbers
regex_one.Pattern = "\b\d+\b"

'replacements specific to barbara walker's treasury
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'replaces the older way to write through back loop
'i.e. p2-b -> p2 tbl
regex_three.Pattern = "\b(p|k)(\d+)-b\b"

s = regex_three.Replace(s, "$1$2 tbl")

'replace these sentences with the abbreviation
s = Replace(s, "Knit all sts through back loops", "ktbl")
s = Replace(s, "Purl all sts through back loops", "ptbl")

'common ocr mistake that makes p1 k1 into pl kl
regex_three.Pattern = "\b(p|k)l\b"
s = regex_three.Replace(s, "$1")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'splits into 1 col with row numbers and another w pattern
regex_three.Pattern = Replace("([^:\.]*):([^:\.]*)\.", ":", delimiter)
Set matches0 = regex_three.Execute(s)
n = matches0.Count

ReDim r_arr(1 To n)
x = 0
tempstring1 = ""

max_row = 1


For i = 0 To n - 1
    'finds row numbers
    Set matches1 = regex_one.Execute(matches0(i).SubMatches(0))
    Set matches2 = regex_two.Execute(matches0(i).SubMatches(0))
    If x = 0 Then
    If matches2.Count <> 0 Then
        If LCase(Left(matches2(0), 1)) = "r" Then
            x = matches1(0)
        Else
            x = matches1(0) + 1
        End If
    End If

    End If
    m = matches1.Count - 1
    ReDim t_arr(0 To m)

    For k = 0 To m
        t_arr(k) = matches1(k)
        If matches1(k) > max_row Then
            max_row = matches1(k)
        End If
    Next
    r_arr(i + 1) = t_arr
Next

ReDim out_arr(1 To max_row, 1 To 2)
'Find Repeat * Pattern
regex_three.Pattern = "rep.*\*"
'find repeat row x pattern
regex_two.Pattern = "rep.+row"
If number_sts <> "" Then
number_sts = Application.Evaluate(number_sts)
End If
For i = 1 To n
    k = 0
    For Each t In r_arr(i)
        
        tempstring1 = matches0(i - 1).SubMatches(1)

        'if has rep(eat) row(s)
        If regex_two.test(tempstring1) Then
            'find the numbers
            Set matches1 = regex_one.Execute(tempstring1)
            tempstring1 = out_arr(max_row - matches1(k) + 1, 2)
            If matches1.Count > 1 Then
                k = k + 1
            End If
        Else
    
        If number_sts <> "" Then
            tempstring1 = tempstring1 & " (" & number_sts & " sts.)"
        End If

        End If
        ''''''''''''''''''''''''''''''''''''''''''''
        'adds rs/ws to beginning of row
        If (Left(tempstring1, 2) = "RS" Or Left(tempstring1, 2) = "WS") Then
            tempstring1 = Right(tempstring1, Len(tempstring1) - 3)
        End If
        If x = 0 Then
            tempstring1 = "RS " & tempstring1
        ElseIf t Mod 2 = x Mod 2 Then
            tempstring1 = "RS " & tempstring1
        Else
            tempstring1 = "WS " & tempstring1
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        'row number
        out_arr(max_row - t + 1, 1) = t
        'instructions
        out_arr(max_row - t + 1, 2) = tempstring1
    Next
    
Next

organizepattern = out_arr

End Function



Sub mergeslst()
Dim rng As Range
Dim c As Range
Dim n As Integer, i As Integer, k As Integer, j As Integer
Dim stitch As String
stitch = InputBox("What represents a slip stitch?", , "*")
Set rng = Selection
n = rng.Rows.Count
For j = 1 To rng.Columns.Count
    For i = 1 To n
        With rng.Cells(i, j)
            k = 0
            While .Offset(k, 0).Value = "*"
                .Offset(k, 0).Value = ""
                k = k + 1
            Wend
            If k > 0 Then
                If .Offset(k, 0).Value = "" Then
                .Offset(-1, 0).Resize(k + 1, 1).Merge
                Else
                .Resize(k + 1, 1).Merge
                End If
            i = i + k
            End If
        End With
    Next
Next
rng.VerticalAlignment = xlCenter


End Sub
