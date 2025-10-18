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

Sub 본문의_프로그래밍_언어명을_한글로_변환()
'
' Java → 자바, 스타일명은 원고에 맞게 각자 조절하기
'
'
    Dim pairs, styles
    Dim i As Long
    Dim s
    Dim sty As Style
    Dim rng As Range
    Dim f As Find
    Dim replacedCount As Long

    ' [영문, 한글] 대치 목록
    pairs = Array( _
        "JavaScript", "자바스크립트", _
        "Java", "자바", _
        "Python", "파이썬", _
        "TypeScript", "타입스크립트", _
        "Ruby", "루비", _
        "Rust", "러스트" _
    )

    ' 대상 스타일 이름들
    styles = Array("일반 (웹)", "표준")

    replacedCount = 0

    For Each s In styles
        Set sty = Nothing
        On Error Resume Next
        Set sty = ActiveDocument.styles(CStr(s))  ' 문서에 없으면 오류 → sty는 Nothing
        On Error GoTo 0

        If Not sty Is Nothing Then
            For i = LBound(pairs) To UBound(pairs) Step 2
                Set rng = ActiveDocument.Content
                Set f = rng.Find

                ' 찾기/바꾸기 설정
                f.ClearFormatting
                f.Replacement.ClearFormatting
                f.Forward = True
                f.Wrap = wdFindContinue

                f.Format = True          ' ★ 스타일 필터 사용
                f.Style = sty            ' ★ 이 스타일에만 적용

                f.Text = pairs(i)
                f.Replacement.Text = pairs(i + 1)
                f.MatchCase = False
                f.MatchWholeWord = False  ' Java vs JavaScript 의 구분을 위해 JavaScript를 먼저 검사함
                f.MatchWildcards = False

                ' 한 번에 전체 대치
                f.Execute Replace:=wdReplaceAll

                ' 대치 건수 추정(정확 집계를 원하면 별도 루프 필요)
                ' Word VBA는 ReplaceAll의 정확한 카운트를 직접 제공하지 않음
                ' 여기서는 메시지 표시는 생략하거나, 필요 시 개별 치환 루프로 교체
            Next i
        End If
    Next s

    MsgBox "대치 완료 (대상 스타일: 일반(웹), 표준).", vbInformation
End Sub


