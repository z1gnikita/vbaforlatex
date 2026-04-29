Attribute VB_Name = "MathConverter"
Option Explicit

Public Sub PasteAndConvertMath()
    Dim startPos As Long
    Dim endPos As Long
    Dim target As Range
    Dim ur As UndoRecord

    On Error GoTo EH

    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Paste and Convert Math"

    startPos = Selection.Start

    On Error Resume Next
    Selection.PasteSpecial DataType:=wdPasteText
    If Err.Number <> 0 Then
        Err.Clear
        Selection.Paste
    End If
    On Error GoTo EH

    endPos = Selection.End

    If endPos > startPos Then
        Set target = ActiveDocument.Range(startPos, endPos)
        ApplyPlainTextFormatting target
        ConvertMathInRange target
    End If

    ur.EndCustomRecord
    Exit Sub

EH:
    On Error Resume Next
    ur.EndCustomRecord
    MsgBox "PasteAndConvertMath failed: " & Err.Description, vbExclamation
End Sub

Public Sub ConvertSelectedMath()
    Dim ur As UndoRecord
    Dim target As Range

    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "Select text first.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH

    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Convert Selected Math"

    Set target = Selection.Range.Duplicate
    ApplyPlainTextFormatting target
    ConvertMathInRange target

    ur.EndCustomRecord
    Exit Sub

EH:
    On Error Resume Next
    ur.EndCustomRecord
    MsgBox "ConvertSelectedMath failed: " & Err.Description, vbExclamation
End Sub

Private Sub ApplyPlainTextFormatting(ByVal rng As Range)
    With rng.Font
        .Name = "Times New Roman"
        .Size = 12
    End With
End Sub

Private Sub ConvertMathInRange(ByVal target As Range)
    Dim re As Object
    Dim matches As Object
    Dim i As Long
    Dim m As Object
    Dim absStart As Long
    Dim absEnd As Long
    Dim tokenRange As Range
    Dim tokenText As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.Pattern = "\$\$[\s\S]+?\$\$|\$[^$]+\$"

    Set matches = re.Execute(target.Text)

    For i = matches.Count - 1 To 0 Step -1
        Set m = matches.Item(i)
        absStart = target.Start + CLng(m.FirstIndex)
        absEnd = absStart + CLng(m.Length)

        Set tokenRange = ActiveDocument.Range(absStart, absEnd)
        tokenText = tokenRange.Text

        If Left$(tokenText, 2) = "$$" And Right$(tokenText, 2) = "$$" Then
            ConvertOneMath tokenRange, Mid$(tokenText, 3, Len(tokenText) - 4), True
        ElseIf Left$(tokenText, 1) = "$" And Right$(tokenText, 1) = "$" Then
            ConvertOneMath tokenRange, Mid$(tokenText, 2, Len(tokenText) - 2), False
        End If
    Next i
End Sub

Private Sub ConvertOneMath(ByVal rng As Range, ByVal eqText As String, ByVal isDisplay As Boolean)
    Dim mathRange As Range

    eqText = Trim$(eqText)
    eqText = Replace(eqText, vbCr, " ")
    eqText = Replace(eqText, vbLf, " ")
    eqText = NormalizeLatex(eqText)

    If Len(eqText) = 0 Then Exit Sub

    rng.Text = eqText

    Set mathRange = ActiveDocument.OMaths.Add(rng)
    mathRange.OMaths(1).BuildUp

    If isDisplay Then
        mathRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
End Sub

Private Function NormalizeLatex(ByVal s As String) As String
    s = Trim$(s)

    ' Remove unsupported LaTeX environments
    s = Replace(s, "\begin{equation*}", "")
    s = Replace(s, "\end{equation*}", "")
    s = Replace(s, "\begin{equation}", "")
    s = Replace(s, "\end{equation}", "")

    s = Replace(s, "\begin{align*}", "")
    s = Replace(s, "\end{align*}", "")
    s = Replace(s, "\begin{align}", "")
    s = Replace(s, "\end{align}", "")

    s = Replace(s, "\begin{aligned}", "")
    s = Replace(s, "\end{aligned}", "")

    ' Remove LaTeX alignment symbols
    s = Replace(s, "\\", " ")
    s = Replace(s, "&=", "=")
    s = Replace(s, "&", "")

    ' Matrices for Word math
    s = Replace(s, "\begin{bmatrix}", "\left[\matrix{")
    s = Replace(s, "\end{bmatrix}", "}\right]")

    s = Replace(s, "\begin{pmatrix}", "\left(\matrix{")
    s = Replace(s, "\end{pmatrix}", "}\right)")

    s = Replace(s, "\begin{vmatrix}", "\left|\matrix{")
    s = Replace(s, "\end{vmatrix}", "}\right|")

    s = Replace(s, "\begin{matrix}", "\matrix{")
    s = Replace(s, "\end{matrix}", "}")

    ' Symbols via Unicode code points
    s = Replace(s, "\cdot", ChrW(&H22C5))      ' dot operator
    s = Replace(s, "\times", ChrW(&HD7))       ' multiplication sign
    s = Replace(s, "\approx", ChrW(&H2248))    ' almost equal
    s = Replace(s, "\sim", ChrW(&H223C))       ' tilde operator
    s = Replace(s, "\pm", ChrW(&HB1))          ' plus-minus

    s = Replace(s, "\Delta", ChrW(&H394))      ' Greek capital Delta
    s = Replace(s, "\delta", ChrW(&H3B4))      ' Greek small delta
    s = Replace(s, "\mu", ChrW(&H3BC))         ' Greek small mu
    s = Replace(s, "\pi", ChrW(&H3C0))         ' Greek small pi
    s = Replace(s, "\varphi", ChrW(&H3C6))     ' Greek small phi
    s = Replace(s, "\phi", ChrW(&H3C6))        ' Greek small phi
    s = Replace(s, "\chi", ChrW(&H3C7))        ' Greek small chi

    s = Replace(s, "\leq", ChrW(&H2264))       ' less-than or equal
    s = Replace(s, "\geq", ChrW(&H2265))       ' greater-than or equal
    s = Replace(s, "\neq", ChrW(&H2260))       ' not equal
    s = Replace(s, "\infty", ChrW(&H221E))     ' infinity

    s = Replace(s, "\sqrt", ChrW(&H221A))      ' square root

    ' LaTeX spacing commands
    s = Replace(s, "\,", " ")
    s = Replace(s, "\;", " ")
    s = Replace(s, "\:", " ")
    s = Replace(s, "\!", "")

    ' Remove \left and \right if Word has problems with them
    s = Replace(s, "\left", "")
    s = Replace(s, "\right", "")

    ' Remove \mathrm{...}, \text{...}, \operatorname{...}
    s = ReplaceSimpleCommandWithBraces(s, "\mathrm")
    s = ReplaceSimpleCommandWithBraces(s, "\text")
    s = ReplaceSimpleCommandWithBraces(s, "\operatorname")

    ' Extra cleanup
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeLatex = Trim$(s)
End Function

Private Function ReplaceSimpleCommandWithBraces(ByVal s As String, ByVal cmd As String) As String
    Dim p As Long
    Dim startBrace As Long
    Dim endBrace As Long
    Dim innerText As String

    p = InStr(1, s, cmd & "{", vbBinaryCompare)

    Do While p > 0
        startBrace = p + Len(cmd)
        endBrace = FindMatchingBrace(s, startBrace)

        If endBrace = 0 Then Exit Do

        innerText = Mid$(s, startBrace + 1, endBrace - startBrace - 1)

        s = Left$(s, p - 1) & innerText & Mid$(s, endBrace + 1)

        p = InStr(1, s, cmd & "{", vbBinaryCompare)
    Loop

    ReplaceSimpleCommandWithBraces = s
End Function

Private Function FindMatchingBrace(ByVal s As String, ByVal openBracePos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String

    depth = 0

    For i = openBracePos To Len(s)
        ch = Mid$(s, i, 1)

        If ch = "{" Then
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1

            If depth = 0 Then
                FindMatchingBrace = i
                Exit Function
            End If
        End If
    Next i

    FindMatchingBrace = 0
End Function
