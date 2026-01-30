Attribute VB_Name = "charttowritten"
'https://makingyenything.wordpress.com/2025/12/27/excel-function-for-converting-a-chart-into-a-written-pattern/
Function convertcharttowritten(arr As Variant, Optional rs As Boolean = True)
Dim outputstring As String
Dim tempstring1 As String
Dim tempstring2 As String
 
Dim regexOne As Object
Dim regexTwo As Object
Dim regexThree As Object
Dim theMatches As Object
Dim Match1 As Object
 
Dim stitchtionary As Object
Dim l As Integer
Dim n As Integer
 
Set regexOne = New RegExp
Set regexTwo = New RegExp
Set regexThree = New RegExp
Set stitchtionary = CreateObject("Scripting.Dictionary")
 
'captures repeating strings greedily, finds big repeating patterns
regexOne.Pattern = "([^\(\)\d]{3,})\1+"
regexOne.Global = True
regexOne.IgnoreCase = False
'captures repeating strings non greedily, finds small repeating patterns
regexTwo.Pattern = "([^\(\)\d]+?)\1+"
regexTwo.Global = True
regexTwo.IgnoreCase = False
Dim a As Integer, b As Integer
 
'allows both ranges and arrays to be used
If TypeName(arr) = "Range" Then
    inputstring = arr.Value
Else
    inputstring = arr
End If
 
'where we start getting characters
l = 65
'a and b find the boundaries of the chart so we can find the stitch count
'simultaneously adding stitches to the dictionary and creating a new string
'to represent the row using unique characters
a = -1
i = 1

For Each c In inputstring
    If c <> "" Then
        If a = -1 Then
            a = i
        End If
        b = i
        If Not stitchtionary.exists(c) Then
            stitchtionary.Add c, Chr(l)
            l = l + 1
            If l = 91 Then
                l = 97
            End If
        End If
            outputstring = outputstring & stitchtionary(c)
    End If
    i = i + 1
Next

'n is number of stitches we just need it for the when we put it in at the end
n = b - a + 1
 
'if on the rightside we read the chart from right to left
If rs Then
    outputstring = StrReverse(outputstring)
End If
 
'1st round of replacements
tempstring1 = outputstring
'find the big consecutively repeating strings and replaces them
Set theMatches = regexOne.Execute(tempstring1)
For Each Match1 In theMatches
    tempstring1 = Replace(tempstring1, Match1, _
    "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
     
Next
 
tempstring2 = outputstring
'find the small consecutively repeating strings and replaces them
Set theMatches = regexTwo.Execute(tempstring2)
For Each Match1 In theMatches
    tempstring2 = Replace(tempstring2, Match1, _
    "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
Next
 
 
'2nd round of replacements
 
'can't use regex replace because can't get the lengths of the matches with that have to for loop
'find the small consecutively repeating strings and replaces them
Set theMatches = regexTwo.Execute(tempstring1)
For Each Match1 In theMatches
    tempstring1 = Replace(tempstring1, Match1, _
    "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
Next
'find the big consecutively repeating strings and replaces them
Set theMatches = regexOne.Execute(tempstring2)
For Each Match1 In theMatches
    tempstring2 = Replace(tempstring2, Match1, _
    "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
Next
 
regexThree.Pattern = "[^\(\)\d\s\*]"
regexThree.Global = True
regexThree.IgnoreCase = False
 
'find which string is better, only counting stitch characters
If regexThree.Execute(tempstring1).Count < regexThree.Execute(tempstring2).Count Then
    outputstring = tempstring1
Else
    outputstring = tempstring2
End If
 
'find the consecutively repeating strings and replaces them one last time
Set theMatches = regexTwo.Execute(outputstring)
For Each Match1 In theMatches
    outputstring = Replace(outputstring, Match1, _
    "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
Next
 
 
'removes stuff that ends up like (k3)x2, which should be k6, unlikely to be needed but just in case
regexThree.Pattern = "\(\((.)\)\*(\d+)\s\)\*(\d+)"
outputstring = regexThree.Replace(outputstring, "($1)*" & "$2" * "$3" & " ")
 
 
'converts characters back to stitches and adds spaces to avoid ambiguity
For Each x In stitchtionary.keys
    outputstring = Replace(outputstring, stitchtionary(x), x & " ")
Next
 
 
'remove spaces in front of right paranthesis and replace * back with x and accidental double spaces
outputstring = Replace(outputstring, " )", ")")
outputstring = Replace(outputstring, "  ", " ")
 
'conventionally, something like k x10, is simplified to k10
regexThree.Pattern = "\(([a-zA-Z])\)\*(\d+)"
outputstring = regexThree.Replace(outputstring, "$1$2")
 
'if there's nested parentheses, change the outside one to brackets
regexThree.Pattern = "\(((?:[^()]*\([^()]*\)[^()]*)+)\)"
outputstring = regexThree.Replace(outputstring, "[$1]")
 
outputstring = Replace(outputstring, "*", " x")
 
If rs Then
outputstring = "RS: " & outputstring
Else
outputstring = "WS: " & outputstring
End If
 
'add number of stitches to the end of the string
convertcharttowritten = outputstring & "(" & n & " sts.) "
 
End Function
Function convertcharttowritten2(arr As Range, Optional flat As Boolean = True, _
 Optional stitchcolumn As Range, Optional rscolumn As Range, _
 Optional wscolumn As Range, Optional startonws As Boolean = False)
 
Dim outputstring As String
Dim inputarr As Variant
Dim tempstring1 As String
Dim tempstring2 As String
Dim rs As Boolean
Dim regexOne As Object
Dim regexTwo As Object
Dim regexThree As Object
Dim theMatches As Object
Dim Match1 As Object
Dim r As Integer, w As Integer, c As String
Dim stitchtionary As Object
Dim a As Integer, b As Integer

Dim l As Integer
Dim n As Integer
Dim i As Integer, j As Integer
Dim x As Variant

Set regexOne = New RegExp
Set regexTwo = New RegExp
Set regexThree = New RegExp
Set stitchtionary = CreateObject("Scripting.Dictionary")

r = arr.Rows.Count
w = arr.Columns.Count
ReDim inputarr(1 To r)
'where we start getting characters
l = 65


'captures repeating strings greedily, finds big repeating patterns
regexOne.Pattern = "([^\(\)\d]{3,})\1+"
regexOne.Global = True
regexOne.IgnoreCase = False
'captures repeating strings non greedily, finds small repeating patterns
regexTwo.Pattern = "([^\(\)\d]+?)\1+"
regexTwo.Global = True
regexTwo.IgnoreCase = False

Dim outputarr As Variant
ReDim outputarr(1 To r)
With arr
For i = 1 To r
    a = -1
    outputstring = ""
    'assumes first row is rs unless you specify otherwise
    If flat Then
        rs = (r - i + startonws) Mod 2 = 0
    Else
        rs = True
    End If
    'start new string each row
    outputstring = ""
    For j = 1 To w
        'perform xlookup if there's columns to do it with otherwise just do it with the
        'text that's in the cell
        If Not (stitchcolumn Is Nothing) Then
            If rs Then
                c = WorksheetFunction.XLookup(.Cells(i, j), stitchcolumn, rscolumn, "")
            Else
                c = WorksheetFunction.XLookup(.Cells(i, j), stitchcolumn, wscolumn, "")
            End If
            Else
            c = .Cells(i, j)
        End If
        If c <> "" Then
            If a = -1 Then
            a = i
            End If
            b = i
            If Not stitchtionary.exists(c) Then
                stitchtionary.Add c, Chr(l)
                l = l + 1
                If l = 91 Then
                    l = 97
                End If
            End If
            outputstring = outputstring & stitchtionary(c)
        End If
    Next j


    'n is number of stitches we just need it for the when we put it in at the end
    n = b - a + 1
    
    'if on the rightside we read the chart from right to left
    If rs Then
        outputstring = StrReverse(outputstring)
    End If
    
    '1st round of replacements
    tempstring1 = outputstring
    'find the big consecutively repeating strings and replaces them
    Set theMatches = regexOne.Execute(tempstring1)
    For Each Match1 In theMatches
        tempstring1 = Replace(tempstring1, Match1, _
        "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
        
    Next
    
    tempstring2 = outputstring
    'find the small consecutively repeating strings and replaces them
    Set theMatches = regexTwo.Execute(tempstring2)
    For Each Match1 In theMatches
        tempstring2 = Replace(tempstring2, Match1, _
        "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
    Next
    
    
    '2nd round of replacements
    
    'can't use regex replace because can't get the lengths of the matches with that have to for loop
    'find the small consecutively repeating strings and replaces them
    Set theMatches = regexTwo.Execute(tempstring1)
    For Each Match1 In theMatches
        tempstring1 = Replace(tempstring1, Match1, _
        "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
    Next
    'find the big consecutively repeating strings and replaces them
    Set theMatches = regexOne.Execute(tempstring2)
    For Each Match1 In theMatches
        tempstring2 = Replace(tempstring2, Match1, _
        "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
    Next
    
    regexThree.Pattern = "[^\(\)\d\s\*]"
    regexThree.Global = True
    regexThree.IgnoreCase = False
    
    'find which string is better, only counting stitch characters
    If regexThree.Execute(tempstring1).Count < regexThree.Execute(tempstring2).Count Then
        outputstring = tempstring1
    Else
        outputstring = tempstring2
    End If
    
    'find the consecutively repeating strings and replaces them one last time
    Set theMatches = regexTwo.Execute(outputstring)
    For Each Match1 In theMatches
        outputstring = Replace(outputstring, Match1, _
        "(" & Match1.SubMatches(0) & ")*" & Len(Match1) / Len(Match1.SubMatches(0)) & " ", , 1)
    Next
    
    
    'removes stuff that ends up like (k3)x2, which should be k6, unlikely to be needed but just in case
    regexThree.Pattern = "\(\((.)\)\*(\d+)\s\)\*(\d+)"
    outputstring = regexThree.Replace(outputstring, "($1)*" & "$2" * "$3" & " ")
    
    
    'converts characters back to stitches and adds spaces to avoid ambiguity
    For Each x In stitchtionary.keys
        outputstring = Replace(outputstring, stitchtionary(x), x & " ")
    Next
    
    
    'remove spaces in front of right paranthesis and replace * back with x and accidental double spaces
    outputstring = Replace(outputstring, " )", ")")
    outputstring = Replace(outputstring, "  ", " ")
    
    'conventionally, something like k x10, is simplified to k10
    regexThree.Pattern = "\(([a-zA-Z])\)\*(\d+)"
    outputstring = regexThree.Replace(outputstring, "$1$2")
    
    'if there's nested parentheses, change the outside one to brackets
    regexThree.Pattern = "\(((?:[^()]*\([^()]*\)[^()]*)+)\)"
    outputstring = regexThree.Replace(outputstring, "[$1]")
    
    outputstring = Replace(outputstring, "*", " x")
    
    If rs Then
    outputstring = "RS: " & outputstring
    Else
    outputstring = "WS: " & outputstring
    End If
    
    'add number of stitches to the end of the string

    outputarr(i) = outputstring & "(" & n & " sts.) "
    
Next i
End With
    convertcharttowritten2 = WorksheetFunction.Transpose(outputarr)
End Function
