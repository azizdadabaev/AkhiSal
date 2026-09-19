Attribute VB_Name = "AkhiSalFuncs"
'==============================================================================
' AkhiSalFuncs - custom worksheet functions (UDFs)
'
' Import:  Alt+F11  ->  File > Import File...  ->  pick this .bas
' Use:     type them in a cell like any built-in, e.g.  =TRANSLIT(A2)
'
' Cyrillic is written as ChrW() codes rather than literal letters on purpose:
' the VBA editor imports .bas files using the machine's ANSI codepage, which
' would turn pasted Cyrillic into garbage on a Latin-locale Windows. ChrW()
' survives that trip intact.
'==============================================================================
Option Explicit

'--- Uzbek Cyrillic -> Latin -------------------------------------------------

' TRANSLIT("Наманган") -> "namangan"
' Lower-cased and Latin, so Cyrillic and Latin spellings of the same client
' finally compare equal. Built for matching, not for display.
Public Function TRANSLIT(ByVal s As String) As String
    Dim i As Long, w As Long, out As String
    For i = 1 To Len(s)
        w = AscW(Mid$(s, i, 1))

        ' fold uppercase down to lowercase before matching
        Select Case w
            Case 1040 To 1071            ' А-Я
                w = w + 32
            Case 65 To 90                ' A-Z
                w = w + 32
            Case 1025                    ' Ё
                w = 1105
            Case 1038                    ' Ў
                w = 1118
            Case 1178                    ' Қ
                w = 1179
            Case 1170                    ' Ғ
                w = 1171
            Case 1202                    ' Ҳ
                w = 1203
        End Select

        Select Case w
            Case 1072: out = out & "a"
            Case 1073: out = out & "b"
            Case 1074: out = out & "v"
            Case 1075: out = out & "g"
            Case 1076: out = out & "d"
            Case 1077: out = out & "e"
            Case 1105: out = out & "yo"
            Case 1078: out = out & "j"
            Case 1079: out = out & "z"
            Case 1080: out = out & "i"
            Case 1081: out = out & "y"
            Case 1082: out = out & "k"
            Case 1083: out = out & "l"
            Case 1084: out = out & "m"
            Case 1085: out = out & "n"
            Case 1086: out = out & "o"
            Case 1087: out = out & "p"
            Case 1088: out = out & "r"
            Case 1089: out = out & "s"
            Case 1090: out = out & "t"
            Case 1091: out = out & "u"
            Case 1092: out = out & "f"
            Case 1093: out = out & "x"
            Case 1094: out = out & "ts"
            Case 1095: out = out & "ch"
            Case 1096: out = out & "sh"
            Case 1097: out = out & "sh"
            Case 1098: out = out & ""     ' ъ
            Case 1099: out = out & "i"
            Case 1100: out = out & ""     ' ь
            Case 1101: out = out & "e"
            Case 1102: out = out & "yu"
            Case 1103: out = out & "ya"
            Case 1118: out = out & "o"    ' ў
            Case 1179: out = out & "q"    ' қ
            Case 1171: out = out & "g"    ' ғ
            Case 1203: out = out & "h"    ' ҳ
            Case Else
                ' already Latin, a digit, or punctuation
                If (w >= 97 And w <= 122) Or (w >= 48 And w <= 57) Then
                    out = out & ChrW(w)
                ElseIf w = 32 Or w = 45 Then
                    out = out & " "
                End If
        End Select
    Next i
    TRANSLIT = Application.WorksheetFunction.Trim(out)
End Function

'--- Fuzzy name matching -----------------------------------------------------

' SIMILARITY("Xalimjon", "Халимжон") -> 1
' Transliterates both sides first, then compares. 0 = nothing in common,
' 1 = identical. Anything above roughly 0.85 is usually the same person.
Public Function SIMILARITY(ByVal a As String, ByVal b As String) As Double
    Dim x As String, y As String, dist As Long, longest As Long
    x = TRANSLIT(a)
    y = TRANSLIT(b)
    longest = Application.WorksheetFunction.Max(Len(x), Len(y))
    If longest = 0 Then
        SIMILARITY = 1
        Exit Function
    End If
    dist = Levenshtein(x, y)
    SIMILARITY = 1 - (dist / longest)
End Function

' MATCHNAME(A2, Clients!A:A) -> the closest client name, or "" below cutoff.
' Third argument is the cutoff and defaults to 0.85.
Public Function MATCHNAME(lookupValue As Variant, searchRange As Range, _
                          Optional cutoff As Double = 0.85) As String
    Dim c As Range, score As Double, best As Double, bestName As String
    Dim needle As String
    needle = CStr(lookupValue)
    If Len(Trim$(needle)) = 0 Then Exit Function

    For Each c In searchRange.Cells
        If Not IsEmpty(c.Value) Then
            score = SIMILARITY(needle, CStr(c.Value))
            If score > best Then
                best = score
                bestName = CStr(c.Value)
                If best = 1 Then Exit For
            End If
        End If
    Next c

    If best >= cutoff Then MATCHNAME = bestName
End Function

' Standard edit distance, iterative two-row version.
Private Function Levenshtein(ByVal s As String, ByVal t As String) As Long
    Dim prev() As Long, curr() As Long
    Dim i As Long, j As Long, cost As Long, n As Long, m As Long

    n = Len(s): m = Len(t)
    If n = 0 Then Levenshtein = m: Exit Function
    If m = 0 Then Levenshtein = n: Exit Function

    ReDim prev(0 To m)
    ReDim curr(0 To m)
    For j = 0 To m
        prev(j) = j
    Next j

    For i = 1 To n
        curr(0) = i
        For j = 1 To m
            If Mid$(s, i, 1) = Mid$(t, j, 1) Then cost = 0 Else cost = 1
            curr(j) = Application.WorksheetFunction.Min( _
                          curr(j - 1) + 1, prev(j) + 1, prev(j - 1) + cost)
        Next j
        For j = 0 To m
            prev(j) = curr(j)
        Next j
    Next i

    Levenshtein = prev(m)
End Function

'--- Money in words (for invoices and contracts) -----------------------------

' UZSWORDS(13500000) -> "o'n uch million besh yuz ming so'm"
' Uzbek numerals are fully regular, so this is a straight place-value walk.
Public Function UZSWORDS(ByVal amount As Double) As String
    Dim n As Double, out As String
    n = Int(Abs(amount))
    If n = 0 Then
        UZSWORDS = "nol so'm"
        Exit Function
    End If

    ' GroupWords pads each place with a trailing space, so collapse the runs
    out = Application.WorksheetFunction.Trim(GroupWords(n))
    If amount < 0 Then out = "minus " & out
    UZSWORDS = out & " so'm"
End Function

Private Function GroupWords(ByVal n As Double) As String
    Dim out As String
    If n >= 1000000000# Then
        out = out & GroupWords(Int(n / 1000000000#)) & " milliard "
        n = n - Int(n / 1000000000#) * 1000000000#
    End If
    If n >= 1000000 Then
        out = out & GroupWords(Int(n / 1000000)) & " million "
        n = n - Int(n / 1000000) * 1000000
    End If
    If n >= 1000 Then
        out = out & GroupWords(Int(n / 1000)) & " ming "
        n = n - Int(n / 1000) * 1000
    End If
    If n >= 100 Then
        out = out & Ones(Int(n / 100)) & " yuz "
        n = n - Int(n / 100) * 100
    End If
    If n >= 20 Then
        out = out & Tens(Int(n / 10)) & " "
        n = n - Int(n / 10) * 10
    ElseIf n >= 10 Then
        out = out & Teens(CLng(n)) & " "
        n = 0
    End If
    If n > 0 Then out = out & Ones(CLng(n)) & " "
    GroupWords = out
End Function

Private Function Ones(ByVal d As Long) As String
    Select Case d
        Case 1: Ones = "bir"
        Case 2: Ones = "ikki"
        Case 3: Ones = "uch"
        Case 4: Ones = "to'rt"
        Case 5: Ones = "besh"
        Case 6: Ones = "olti"
        Case 7: Ones = "yetti"
        Case 8: Ones = "sakkiz"
        Case 9: Ones = "to'qqiz"
    End Select
End Function

Private Function Teens(ByVal d As Long) As String
    ' 10-19: o'n, o'n bir, o'n ikki, ...
    If d = 10 Then
        Teens = "o'n"
    Else
        Teens = "o'n " & Ones(d - 10)
    End If
End Function

Private Function Tens(ByVal d As Long) As String
    Select Case d
        Case 2: Tens = "yigirma"
        Case 3: Tens = "o'ttiz"
        Case 4: Tens = "qirq"
        Case 5: Tens = "ellik"
        Case 6: Tens = "oltmish"
        Case 7: Tens = "yetmish"
        Case 8: Tens = "sakson"
        Case 9: Tens = "to'qson"
    End Select
End Function

'--- Small everyday helpers --------------------------------------------------

' M2(1200, 600, 40) -> 28.8   (millimetres in, square metres out)
Public Function M2(ByVal lengthMM As Double, ByVal widthMM As Double, _
                   Optional ByVal qty As Double = 1) As Double
    M2 = (lengthMM / 1000) * (widthMM / 1000) * qty
End Function

' AGEDAYS(scheduledDate) -> whole days from that date until today.
' Negative while the date is still in the future, so an aging column reads
' straight: 0 and below is not yet due, above 0 is overdue by that many days.
Public Function AGEDAYS(ByVal d As Variant) As Variant
    If Not IsDate(d) Then
        AGEDAYS = CVErr(xlErrValue)
    Else
        Application.Volatile
        AGEDAYS = DateDiff("d", CDate(d), Date)
    End If
End Function

' AGEBUCKET(scheduledDate) -> "not due" / "0-30" / "31-60" / "61-90" / "90+"
Public Function AGEBUCKET(ByVal d As Variant) As Variant
    Dim n As Variant
    n = AGEDAYS(d)
    If IsError(n) Then
        AGEBUCKET = n
    ElseIf n <= 0 Then
        AGEBUCKET = "not due"
    ElseIf n <= 30 Then
        AGEBUCKET = "0-30"
    ElseIf n <= 60 Then
        AGEBUCKET = "31-60"
    ElseIf n <= 90 Then
        AGEBUCKET = "61-90"
    Else
        AGEBUCKET = "90+"
    End If
End Function
