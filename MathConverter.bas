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
        RemoveLatexSectionsInRange target

        ' ÂÀÆÍÎ:
        ' Ñíà÷àëà ôîðìóëû, ïîòîì òàáëèöû.
        ConvertMathInRange target
        ConvertLatexTablesInRange target
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
    RemoveLatexSectionsInRange target

    ' ÂÀÆÍÎ:
    ' Ñíà÷àëà ôîðìóëû, ïîòîì òàáëèöû.
    ConvertMathInRange target
    ConvertLatexTablesInRange target

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

' =========================================================
' REMOVE \section{...}, \subsection{...}, \subsubsection{...}
' =========================================================

Private Sub RemoveLatexSectionsInRange(ByVal target As Range)
    ' ÂÀÆÍÎ:
    ' Ñíà÷àëà áîëåå äëèííûå êîìàíäû, ïîòîì êîðîòêèå.
    ' Òàê áåçîïàñíåå äëÿ \subsubsection / \subsection / \section.
    UnwrapLatexCommandInRange target, "\subsubsection"
    UnwrapLatexCommandInRange target, "\subsection"
    UnwrapLatexCommandInRange target, "\section"
End Sub

Private Sub UnwrapLatexCommandInRange(ByVal target As Range, ByVal cmd As String)
    Dim docText As String
    Dim p As Long
    Dim openBracePos As Long
    Dim closeBracePos As Long
    Dim innerText As String
    Dim replaceRange As Range

    docText = target.text
    p = InStr(1, docText, cmd & "{", vbBinaryCompare)

    Do While p > 0
        openBracePos = p + Len(cmd)
        closeBracePos = FindMatchingBrace(docText, openBracePos)

        If closeBracePos = 0 Then Exit Do

        innerText = Mid$(docText, openBracePos + 1, closeBracePos - openBracePos - 1)

        Set replaceRange = ActiveDocument.Range( _
            target.Start + p - 1, _
            target.Start + closeBracePos _
        )

        replaceRange.text = innerText

        docText = target.text
        p = InStr(1, docText, cmd & "{", vbBinaryCompare)
    Loop
End Sub

' =========================================================
' LATEX TABLES -> WORD TABLES
' =========================================================

Private Sub ConvertLatexTablesInRange(ByVal target As Range)
    Dim re As Object
    Dim matches As Object
    Dim i As Long
    Dim m As Object
    Dim blockRange As Range

    ' Îáðàáàòûâàåì òîëüêî ïîëíîöåííûå áëîêè:
    ' \begin{table}...\end{table}
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.pattern = "\\begin\{table\}(?:\[[^\]]*\])?[\s\S]*?\\end\{table\}"

    Set matches = re.Execute(target.text)

    For i = matches.Count - 1 To 0 Step -1
        Set m = matches.Item(i)
        Set blockRange = ActiveDocument.Range( _
            target.Start + CLng(m.FirstIndex), _
            target.Start + CLng(m.FirstIndex) + CLng(m.length) _
        )
        ConvertOneLatexTable blockRange, blockRange.text
    Next i
End Sub

Private Sub ConvertOneLatexTable(ByVal blockRange As Range, ByVal blockText As String)
    Dim colSpec As String
    Dim tableBody As String
    Dim captionText As String
    Dim cleanedBody As String
    Dim rows() As String
    Dim cellTexts() As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim tbl As Table
    Dim insertRange As Range
    Dim tmp As String
    Dim afterRange As Range

    If Not ExtractTabular(blockText, colSpec, tableBody) Then Exit Sub
    captionText = ExtractCaption(blockText)

    cleanedBody = NormalizeTableBody(tableBody)
    rows = Split(cleanedBody, "§ROW§")

    rowCount = 0
    colCount = 0

    For i = LBound(rows) To UBound(rows)
        tmp = Trim$(rows(i))
        If Len(tmp) > 0 Then
            rowCount = rowCount + 1
            cellTexts = Split(tmp, "&")
            If UBound(cellTexts) - LBound(cellTexts) + 1 > colCount Then
                colCount = UBound(cellTexts) - LBound(cellTexts) + 1
            End If
        End If
    Next i

    If rowCount = 0 Or colCount = 0 Then Exit Sub

    Set insertRange = blockRange.Duplicate
    insertRange.text = ""

    Set tbl = ActiveDocument.Tables.Add(insertRange, rowCount, colCount)
    tbl.Borders.Enable = True

    r = 0
    For i = LBound(rows) To UBound(rows)
        tmp = Trim$(rows(i))
        If Len(tmp) > 0 Then
            r = r + 1
            cellTexts = Split(tmp, "&")

            For c = 1 To colCount
                If c - 1 <= UBound(cellTexts) Then
                    tbl.Cell(r, c).Range.text = Trim$(cellTexts(c - 1))
                Else
                    tbl.Cell(r, c).Range.text = ""
                End If

                FormatAndConvertTableCell tbl.Cell(r, c)
            Next c
        End If
    Next i

    ' Ñîçäà¸ì îáû÷íûé àáçàö ïîñëå òàáëèöû,
    ' ÷òîáû ñëåäóþùèé òåêñò íå ïîïàäàë âíóòðü òàáëèöû.
    Set afterRange = tbl.Range.Duplicate
    afterRange.Collapse wdCollapseEnd
    afterRange.InsertParagraphAfter
    afterRange.Collapse wdCollapseEnd

    If Len(Trim$(captionText)) > 0 Then
        afterRange.InsertAfter captionText & vbCr
    Else
        afterRange.InsertAfter vbCr
    End If
End Sub

Private Sub FormatAndConvertTableCell(ByVal cl As Cell)
    Dim cellRng As Range
    Dim txt As String

    Set cellRng = cl.Range
    cellRng.End = cellRng.End - 1

    txt = Trim$(cellRng.text)
    If Len(txt) = 0 Then Exit Sub

    ApplyPlainTextFormatting cellRng

    If IsLikelyMathText(txt) Then
        txt = NormalizeLatex(txt)
        cellRng.text = txt

        On Error Resume Next
        ActiveDocument.OMaths.Add(cellRng).OMaths(1).BuildUp
        On Error GoTo 0
    Else
        cellRng.text = txt
    End If
End Sub

Private Function IsLikelyMathText(ByVal s As String) As Boolean
    s = Trim$(s)

    If Len(s) = 0 Then
        IsLikelyMathText = False
        Exit Function
    End If

    If InStr(s, "\") > 0 _
        Or InStr(s, "_") > 0 _
        Or InStr(s, "^") > 0 _
        Or InStr(s, "=") > 0 _
        Or InStr(s, "+") > 0 _
        Or InStr(s, "-") > 0 _
        Or InStr(s, "g_") > 0 _
        Or InStr(s, "C_") > 0 _
        Or InStr(s, "R_") > 0 _
        Or InStr(s, "U_") > 0 _
        Or InStr(s, "I_") > 0 _
        Or InStr(s, "K") > 0 _
        Or InStr(s, "j") > 0 _
        Or InStr(s, "?") > 0 Then
        IsLikelyMathText = True
    Else
        IsLikelyMathText = False
    End If
End Function

Private Function ExtractTabular(ByVal blockText As String, ByRef colSpec As String, ByRef tableBody As String) As Boolean
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = True
    re.pattern = "\\begin\{tabular\}\{([^}]*)\}([\s\S]*?)\\end\{tabular\}"

    Set matches = re.Execute(blockText)

    If matches.Count = 0 Then
        ExtractTabular = False
        Exit Function
    End If

    colSpec = matches(0).SubMatches(0)
    tableBody = matches(0).SubMatches(1)
    ExtractTabular = True
End Function

Private Function ExtractCaption(ByVal blockText As String) As String
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = True
    re.pattern = "\\caption\{([\s\S]*?)\}"

    Set matches = re.Execute(blockText)

    If matches.Count = 0 Then
        ExtractCaption = ""
    Else
        ExtractCaption = Trim$(matches(0).SubMatches(0))
    End If
End Function

Private Function NormalizeTableBody(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    s = Replace(s, "\centering", "")
    s = Replace(s, "\hline", "")
    s = RegexReplaceString(s, "\\cline\{[^}]*\}", "")

    s = RegexReplaceString(s, "\\\\\[[^\]]*\]", "§ROW§")
    s = Replace(s, "\\", "§ROW§")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeTableBody = Trim$(s)
End Function

' =========================================================
' MATH CONVERSION
' Supports:
'   $$ ... $$
'   $ ... $
'   \[ ... \]
'   \( ... \)
' =========================================================

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

    re.pattern = "\$\$[\s\S]+?\$\$|\$[^$]+\$|\\\[[\s\S]+?\\\]|\\\([\s\S]+?\\\)"

    Set matches = re.Execute(target.text)

    For i = matches.Count - 1 To 0 Step -1
        Set m = matches.Item(i)
        absStart = target.Start + CLng(m.FirstIndex)
        absEnd = absStart + CLng(m.length)

        Set tokenRange = ActiveDocument.Range(absStart, absEnd)
        tokenText = tokenRange.text

        If Left$(tokenText, 2) = "$$" And Right$(tokenText, 2) = "$$" Then
            ConvertOneMath tokenRange, Mid$(tokenText, 3, Len(tokenText) - 4), True

        ElseIf Left$(tokenText, 1) = "$" And Right$(tokenText, 1) = "$" Then
            ConvertOneMath tokenRange, Mid$(tokenText, 2, Len(tokenText) - 2), False

        ElseIf Left$(tokenText, 2) = "\[" And Right$(tokenText, 2) = "\]" Then
            ConvertOneMath tokenRange, Mid$(tokenText, 3, Len(tokenText) - 4), True

        ElseIf Left$(tokenText, 2) = "\(" And Right$(tokenText, 2) = "\)" Then
            ConvertOneMath tokenRange, Mid$(tokenText, 3, Len(tokenText) - 4), False
        End If
    Next i
End Sub

Private Sub ConvertOneMath(ByVal rng As Range, ByVal eqText As String, ByVal isDisplay As Boolean)
    Dim mathRange As Range

    On Error GoTo FormulaEH

    eqText = Trim$(eqText)
    eqText = Replace(eqText, vbCr, " ")
    eqText = Replace(eqText, vbLf, " ")
    eqText = NormalizeLatex(eqText)

    If Len(eqText) = 0 Then Exit Sub

    rng.text = eqText

    Set mathRange = ActiveDocument.OMaths.Add(rng)
    mathRange.OMaths(1).BuildUp

    If isDisplay Then
        mathRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If

    Exit Sub

FormulaEH:
    ' Åñëè êîíêðåòíàÿ ôîðìóëà íå ñîáðàëàñü,
    ' îñòàâëÿåì å¸ êàê íîðìàëèçîâàííûé òåêñò è ïðîäîëæàåì ìàêðîñ.
    On Error Resume Next
    rng.text = eqText
End Sub

' =========================================================
' LATEX NORMALIZATION
' =========================================================

Private Function NormalizeLatex(ByVal s As String) As String
    s = Trim$(s)

    ' Èñïðàâëåíèå äëÿ ðàçîðâàííûõ LaTeX-ñêîáîê:
    ' \left[ ... \right.  è  \left. ... \right]
    ' Word ïëîõî ïåðåíîñèò òàêèå ïàðû ìåæäó ðàçíûìè \[...\] áëîêàìè.
    s = FixSplitLeftRightDelimiters(s)

    ' Äåñÿòè÷íàÿ çàïÿòàÿ LaTeX: 0{,}2 -> 0,2
    s = Replace(s, "{,}", ",")

    ' Íîðìàëèçóåì äðîáè:
    ' \dfrac è \tfrac ïåðåâîäèì â \frac,
    ' à \frac{...} {...} ïðèâîäèì ê \frac{...}{...}
    s = Replace(s, "\dfrac", "\frac")
    s = Replace(s, "\tfrac", "\frac")
    s = CompactLatexFractions(s)

    ' Ìàòðèöû:
    ' Word ÍÅ ïîääåðæèâàåò \begin{matrix}...\end{matrix}.
    ' Äëÿ Word-LaTeX èñïîëüçóåì:
    ' \matrix{a & b \\ c & d}
    '
    ' bmatrix -> \left[\matrix{...}\right]
    ' pmatrix -> \left(\matrix{...}\right)
    ' vmatrix -> \left|\matrix{...}\right|
    ' matrix  -> \matrix{...}
    '
    ' Âíóòðè ìàòðèö \\ âðåìåííî çàùèùàåì êàê §MROW§,
    ' ÷òîáû îáùàÿ ñòðîêà s = Replace(s, "\\", " ") íå óíè÷òîæèëà ñòðîêè ìàòðèöû.
    s = ConvertMatrixEnvironment(s, "bmatrix", "§MLEFT§[", "§MRIGHT§]")
    s = ConvertMatrixEnvironment(s, "pmatrix", "§MLEFT§(", "§MRIGHT§)")
    s = ConvertMatrixEnvironment(s, "vmatrix", "§MLEFT§|", "§MRIGHT§|")
    s = ConvertMatrixEnvironment(s, "matrix", "", "")

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

    ' Remove LaTeX alignment symbols outside matrices.
    ' Âíóòðè ìàòðèö \\ óæå çàìåíåíû íà §MROW§.
    s = Replace(s, "\\", " ")
    s = Replace(s, "&=", "=")
    ' Ãëîáàëüíî "&" ÍÅ óäàëÿåì: îí íóæåí äëÿ ñòîëáöîâ ìàòðèöû.

    ' Symbols via Unicode code points
    s = Replace(s, "\cdot", ChrW(&H22C5))
    s = Replace(s, "\times", ChrW(&HD7))
    s = Replace(s, "\approx", ChrW(&H2248))
    s = Replace(s, "\sim", ChrW(&H223C))
    s = Replace(s, "\pm", ChrW(&HB1))

    s = Replace(s, "\Delta", ChrW(&H394))
    s = Replace(s, "\delta", ChrW(&H3B4))
    s = Replace(s, "\mu", ChrW(&H3BC))
    s = Replace(s, "\pi", ChrW(&H3C0))
    s = Replace(s, "\omega", ChrW(&H3C9))
    s = Replace(s, "\varphi", ChrW(&H3C6))
    s = Replace(s, "\phi", ChrW(&H3C6))
    s = Replace(s, "\chi", ChrW(&H3C7))

    s = Replace(s, "\leq", ChrW(&H2264))
    s = Replace(s, "\geq", ChrW(&H2265))
    s = Replace(s, "\neq", ChrW(&H2260))
    s = Replace(s, "\infty", ChrW(&H221E))

    s = Replace(s, "\sqrt", ChrW(&H221A))

    ' LaTeX spacing commands
    s = Replace(s, "\,", " ")
    s = Replace(s, "\;", " ")
    s = Replace(s, "\:", " ")
    s = Replace(s, "\!", "")

    ' Remove \left and \right if Word has problems with them,
    ' íî ÍÅ òðîãàåì çàùèù¸ííûå matrix-left/right.
    s = Replace(s, "\left", "")
    s = Replace(s, "\right", "")

    ' Remove \mathrm{...}, \text{...}, \operatorname{...}
    s = ReplaceSimpleCommandWithBraces(s, "\mathrm")
    s = ReplaceSimpleCommandWithBraces(s, "\text")
    s = ReplaceSimpleCommandWithBraces(s, "\operatorname")

    ' Âîññòàíàâëèâàåì matrix syntax ïîñëå îáùåé ÷èñòêè.
    s = Replace(s, "§MROW§", "\\")
    s = Replace(s, "§MLEFT§", "\left")
    s = Replace(s, "§MRIGHT§", "\right")

    ' Extra cleanup
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeLatex = Trim$(s)
End Function

Private Function FixSplitLeftRightDelimiters(ByVal s As String) As String
    ' Ñëó÷àé 1:
    ' \left[ ... \right.
    ' Â Word ýòî îòäåëüíàÿ ôîðìóëà ñ íåçàêðûòîé ëåâîé ñêîáêîé.
    If InStr(1, s, "\right.", vbBinaryCompare) > 0 Then
        s = Replace(s, "\right.", "")

        s = Replace(s, "\left[", "")
        s = Replace(s, "\left(", "")
        s = Replace(s, "\left|", "")
        s = Replace(s, "\left\{", "")
        s = Replace(s, "\left\lbrace", "")
        s = Replace(s, "\left\langle", "")
    End If

    ' Ñëó÷àé 2:
    ' \left. ... \right]
    ' Â Word ýòî îòäåëüíàÿ ôîðìóëà ñ íåçàêðûòîé ïðàâîé ñêîáêîé.
    If InStr(1, s, "\left.", vbBinaryCompare) > 0 Then
        s = Replace(s, "\left.", "")

        s = Replace(s, "\right]", "")
        s = Replace(s, "\right)", "")
        s = Replace(s, "\right|", "")
        s = Replace(s, "\right\}", "")
        s = Replace(s, "\right\rbrace", "")
        s = Replace(s, "\right\rangle", "")
    End If

    FixSplitLeftRightDelimiters = s
End Function

Private Function CompactLatexFractions(ByVal s As String) As String
    Dim p As Long
    Dim fracStart As Long
    Dim numOpen As Long
    Dim numClose As Long
    Dim denOpen As Long
    Dim denClose As Long
    Dim beforeText As String
    Dim afterText As String
    Dim numText As String
    Dim denText As String

    p = InStr(1, s, "\frac", vbBinaryCompare)

    Do While p > 0
        fracStart = p

        ' Èùåì îòêðûâàþùóþ ñêîáêó ÷èñëèòåëÿ ïîñëå \frac
        numOpen = fracStart + Len("\frac")

        Do While numOpen <= Len(s) And Mid$(s, numOpen, 1) = " "
            numOpen = numOpen + 1
        Loop

        If numOpen > Len(s) Or Mid$(s, numOpen, 1) <> "{" Then
            p = InStr(fracStart + Len("\frac"), s, "\frac", vbBinaryCompare)
            GoTo ContinueLoop
        End If

        numClose = FindMatchingBrace(s, numOpen)
        If numClose = 0 Then
            p = InStr(fracStart + Len("\frac"), s, "\frac", vbBinaryCompare)
            GoTo ContinueLoop
        End If

        ' Èùåì îòêðûâàþùóþ ñêîáêó çíàìåíàòåëÿ ïîñëå ÷èñëèòåëÿ.
        ' Òóò óáèðàåì ïðîáåëû/ïåðåíîñû ìåæäó } è {
        denOpen = numClose + 1

        Do While denOpen <= Len(s) And _
            (Mid$(s, denOpen, 1) = " " Or _
             Mid$(s, denOpen, 1) = vbTab Or _
             Mid$(s, denOpen, 1) = vbCr Or _
             Mid$(s, denOpen, 1) = vbLf)
            denOpen = denOpen + 1
        Loop

        If denOpen > Len(s) Or Mid$(s, denOpen, 1) <> "{" Then
            p = InStr(fracStart + Len("\frac"), s, "\frac", vbBinaryCompare)
            GoTo ContinueLoop
        End If

        denClose = FindMatchingBrace(s, denOpen)
        If denClose = 0 Then
            p = InStr(fracStart + Len("\frac"), s, "\frac", vbBinaryCompare)
            GoTo ContinueLoop
        End If

        beforeText = Left$(s, fracStart - 1)
        numText = Mid$(s, numOpen + 1, numClose - numOpen - 1)
        denText = Mid$(s, denOpen + 1, denClose - denOpen - 1)
        afterText = Mid$(s, denClose + 1)

        s = beforeText & "\frac{" & Trim$(numText) & "}{" & Trim$(denText) & "}" & afterText

        p = InStr(fracStart + Len("\frac"), s, "\frac", vbBinaryCompare)

ContinueLoop:
    Loop

    CompactLatexFractions = s
End Function

Private Function ConvertMatrixEnvironment(ByVal s As String, ByVal envName As String, ByVal leftDelim As String, ByVal rightDelim As String) As String
    Dim beginTag As String
    Dim endTag As String
    Dim p As Long
    Dim q As Long
    Dim innerText As String
    Dim matrixText As String
    Dim replacement As String

    beginTag = "\begin{" & envName & "}"
    endTag = "\end{" & envName & "}"

    p = InStr(1, s, beginTag, vbBinaryCompare)

    Do While p > 0
        q = InStr(p + Len(beginTag), s, endTag, vbBinaryCompare)
        If q = 0 Then Exit Do

        innerText = Mid$(s, p + Len(beginTag), q - (p + Len(beginTag)))
        matrixText = PrepareMatrixContent(innerText)

        If leftDelim = "" And rightDelim = "" Then
            replacement = "\matrix{" & matrixText & "}"
        Else
            replacement = leftDelim & "\matrix{" & matrixText & "}" & rightDelim
        End If

        s = Left$(s, p - 1) & replacement & Mid$(s, q + Len(endTag))
        p = InStr(1, s, beginTag, vbBinaryCompare)
    Loop

    ConvertMatrixEnvironment = s
End Function

Private Function PrepareMatrixContent(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    ' \\[8pt] -> çàùèù¸ííûé LaTeX-ðàçäåëèòåëü ñòðîêè
    s = RegexReplaceString(s, "\\\\\[[^\]]*\]", "§MROW§")

    ' îáû÷íûå ïåðåõîäû ñòðîê ìàòðèöû \\ -> çàùèù¸ííûé ðàçäåëèòåëü
    s = Replace(s, "\\", "§MROW§")

    ' Óáèðàåì LaTeX-ïðîáåëû âíóòðè ìàòðèöû
    s = Replace(s, "\qquad", " ")
    s = Replace(s, "\quad", " ")
    s = Replace(s, "\,", " ")
    s = Replace(s, "\;", " ")
    s = Replace(s, "\:", " ")
    s = Replace(s, "\!", "")
    s = Replace(s, "~", " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    PrepareMatrixContent = Trim$(s)
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

Private Function RegexReplaceString(ByVal text As String, ByVal pattern As String, ByVal replacement As String) As String
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.pattern = pattern

    RegexReplaceString = re.Replace(text, replacement)
End Function

