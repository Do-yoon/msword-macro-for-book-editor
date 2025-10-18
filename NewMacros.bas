Sub 영문_단위_띄어쓰기()
'
' 영문_단위_띄어쓰기 Macro
'
'
    Dim findRange As Range
    Set findRange = ActiveDocument.Content

    Dim unitList As Variant
    unitList = Array("m", "mm", "cm", "km", "nm", _
                     "byte", "kb", "kB", "KB", "mb", "MB", "gb", "GB", "tb", "TB", "kbps", _
                     "ml", "mL", "l", "L", _
                     "ms", "Hz", "GHz", _
                     "kcal")

    Dim functionList As Variant
    functionList = Array("log", "ln", "sin", "cos", "tan", "atan", "asin", "acos")

    Application.UndoRecord.StartCustomRecord "Insert space between number and unit"

    With findRange.Find
        .ClearFormatting
        .Text = "([0-9])([a-zA-Z])"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With

    Do While findRange.Find.Execute
        Dim foundText As String
        foundText = findRange.Text

        ' 예외 처리: 고정 표현
        If foundText = "2D" Or foundText = "3D" Then
            GoTo SkipReplacement
        End If

        ' 앞 문맥 함수 예외 검사
        Dim funcName As Variant
        For Each funcName In functionList
            Dim funcLen As Integer
            funcLen = Len(funcName)

            Dim contextRange As Range
            Set contextRange = findRange.Duplicate
            On Error Resume Next
            contextRange.MoveStart wdCharacter, -funcLen
            On Error GoTo 0
            Dim contextText As String
            contextText = LCase(contextRange.Text)

            If contextText = funcName Then
                GoTo SkipReplacement
            End If
        Next funcName

        ' 단위 접두사 일치 검사
        Dim unitName As Variant
        For Each unitName In unitList
            If LCase(Mid(findRange.Text, 2)) Like LCase(unitName & "*") Then
                ' 공백 삽입
                If Len(foundText) = 2 Then
                    findRange.Text = Left(foundText, 1) & " " & Right(foundText, 1)
                ElseIf Len(foundText) > 2 Then
                    findRange.Text = Left(foundText, 1) & " " & Mid(foundText, 2)
                End If
                Exit For
            End If
        Next unitName

SkipReplacement:
        findRange.Collapse Direction:=wdCollapseEnd
    Loop

    Application.UndoRecord.EndCustomRecord
End Sub


Sub 영어_단어와_식을_띄우기()
'
' 예: CASE식 -> CASE 식, '식' 대신 '절' 등에도 적용 가능
'
'
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "([A-Z])식"
        .Replacement.Text = "\1 식"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub 책_제목_괄호_대치()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    ' 왼쪽 괄호 대치: 『 → 《
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "『"
        .Replacement.Text = "《"
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 오른쪽 괄호 대치: 』 → 》
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "』"
        .Replacement.Text = "》"
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    MsgBox "『』 → 《》 로 변환 완료!", vbInformation
End Sub

