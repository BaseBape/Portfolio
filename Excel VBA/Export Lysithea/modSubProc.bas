Attribute VB_Name = "modSubProcess"
Option Explicit

'**********
'新規作成
'**********
'カレンダー
Public Function createCalendar()
    With wbInfo
    'リシテア取込
    '日付
        .wsExport.Cells(.rowExportCalendar, .colShiftCalendar).Value = .targetDay
    '曜日
        .wsExport.Cells(.rowExportCalendarWeekday, .colShiftCalendar).Value = .targetDay
    'シフト表
    '日付
        .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value = .targetDay
    '曜日
        .wsShift.Cells(.rowShiftCalendarWeekday, .colShiftCalendar).Value = .targetDay
                
    '土曜・日曜のときハッチング
        Select Case Weekday(.targetDay)
            Case 1, 7
                .wsExport.Range(.wsExport.Cells(.rowExportCalendarWeekday, .colShiftCalendar), _
                                .wsExport.Cells(.rowExportCalendar, .colShiftCalendar)).Interior.Color = RGB(255, 204, 255)
                .wsShift.Range(.wsShift.Cells(.rowShiftCalendar, .colShiftCalendar), _
                                .wsShift.Cells(.rowShiftCalendarWeekday, .colShiftCalendar)).Interior.Color = RGB(255, 204, 255)
        End Select
                
    '社休日一覧に有ったらハッチング
        Select Case WorksheetFunction.CountIf(Range("holidayList"), .targetDay)
            Case Is > 0
                .wsExport.Range(.wsExport.Cells(.rowExportCalendarWeekday, .colShiftCalendar), _
                                .wsExport.Cells(.rowExportCalendar, .colShiftCalendar)).Interior.Color = RGB(255, 204, 255)
                .wsShift.Range(.wsShift.Cells(.rowShiftCalendar, .colShiftCalendar), _
                                .wsShift.Cells(.rowShiftCalendarWeekday, .colShiftCalendar)).Interior.Color = RGB(255, 204, 255)
        End Select
    End With
End Function

'1年間デフォルトシフト
Public Function createShiftYear()
    With wbInfo
        For .rowShiftRole = .rowShiftStartRole To .rowShiftEndRole Step 2
            For .colShiftCalendar = .colShiftStartCalendar To .colShiftEndCalendar
                Select Case .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value
                    Case Is <> ""
                        Call modSubProcess.reflectClassWorkType
                End Select
            Next .colShiftCalendar
        Next .rowShiftRole
    End With
End Function

'リシテア取込シート
Public Function createExportLysithea()
    With wbInfo
        For .rowExportShift = .rowExportStartShift To .rowExportEndShift - 2 Step 2
            For .colExportCalendar = .colExportStartCalendar To .colExportEndCalendar
                Select Case .wsExport.Cells(.rowExportCalendar, .colExportCalendar).Value
                    Case Is <> ""
                        Call modSubProcess.reflectClassWorkday
                End Select
            Next .colExportCalendar
        Next .rowExportShift
    End With
End Function

'*************
'データの統合
'*************
Public Function inputData()
    With wbInfo
    '検索値の格納(職員名)
        .shiftName = .wsShift.Cells(.rowShiftRole, .colShiftName).Value
        Select Case InStr(.shiftName, "　")
            Case Is <> 0
                .shiftName = Replace(.shiftName, "　", "")
        End Select
        
        Select Case InStr(.shiftName, " ")
            Case Is <> 0
                .shiftName = Replace(.shiftName, " ", "")
        End Select
        
        .targetShift = .wsTarget.Cells(.rowTargetShift, 1).Value & .wsTarget.Cells(.rowTargetShift, 2).Value
        Select Case InStr(.targetShift, "　")
            Case Is <> 0
                .targetShift = Replace(.targetShift, "　", "")
        End Select
        
        Select Case InStr(.targetShift, " ")
            Case Is <> 0
                .targetShift = Replace(.targetShift, " ", "")
        End Select
        
        Select Case .targetShift
            Case .shiftName
                For .colTargetCalendar = .colTargetStartCalendar To .colTargetEndCalendar
                '検索値の格納(日時)
                    .shiftCalendar = .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value
                    .targetCalendar = .wsTarget.Cells(4, .colTargetCalendar).Value
                    
                    Select Case .targetCalendar
                        Case .shiftCalendar
                            Select Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value
                            'シフト表シートの値が「空白」
                                Case ""
                                    Select Case .wsTarget.Cells(.rowTargetShift, .colTargetCalendar).Value
                                    '現場シフトが「★」「●」「▲」「■」「夜勤」
                                        Case CATEGORYSITEWORK_1, CATEGORYSITEWORK_2, CATEGORYSITEWORK_3, CATEGORYSITEWORK_4, CATEGORYSITEWORK_5
                                        '「夜8」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYWORK_10
                                    '現場シフトが「休」「休希望」
                                        Case CATEGORYSITEREST_1, CATEGORYSITEREST_6
                                        '「休」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_1
                                    '現場シフトが「代休」
                                        Case CATEGORYSITEREST_2
                                        '「代」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_2
                                    '現場シフトが「/」
                                        Case CATEGORYSITEREST_3
                                        '「明」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_3
                                    '現場シフトが「年休」
                                        Case CATEGORYSITEREST_4
                                        '「年休」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_4
                                    '現場シフトが「振休」
                                        Case CATEGORYSITEREST_5
                                            .targetDate = Year(.startDay) & "/" & Replace(.wsTarget.Cells(.rowTargetShift + 1, .colTargetCalendar).Value, "分", "")
                                            .targetWeekday = Weekday(.targetDate)
                                        '振替出勤日に応じてシフト入力
                                            Select Case .targetWeekday
                                            '振替出勤日が日曜の時「日振」
                                                Case 1
                                                    .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_5
                                            '振り返出勤日が日曜以外の時「土祝振」
                                                Case Else
                                                    .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_6
                                            End Select
                                        Case Else
                                        '現場シフトにグレーハッチングがかかっていた時「休」
                                            If .wsTarget.Cells(.rowTargetShift, .colTargetCalendar).DisplayFormat.Interior.Color = RGB(191, 191, 191) Then
                                                .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_1
                                        'グレーハッチング以外の時「昼8」
                                            Else
                                                .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYWORK_5
                                            End If
                                    End Select
                            'シフト表シートに値が入っている
                                Case Else
                                    Select Case .wsTarget.Cells(.rowTargetShift, .colTargetCalendar).Value
                                    '現場シフトが「★」「●」「▲」「■」「夜勤」
                                        Case CATEGORYSITEWORK_1, CATEGORYSITEWORK_2, CATEGORYSITEWORK_3, CATEGORYSITEWORK_4, CATEGORYSITEWORK_5
                                        '「夜」を含むシフトが入力されていない時「夜8」
                                            If Not (.wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value Like "*夜*") Then
                                                .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYWORK_10
                                            End If
                                    '現場シフトが「休」「休希望」
                                        Case CATEGORYSITEREST_1, CATEGORYSITEREST_6
                                        '「休」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_1
                                    '現場シフトが「代休」
                                        Case CATEGORYSITEREST_2
                                        '「代」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_2
                                    '現場シフトが「/」
                                        Case CATEGORYSITEREST_3
                                        '「明」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_3
                                    '現場シフトが「年休」
                                        Case CATEGORYSITEREST_4
                                        '「年休」を入力
                                            .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_4
                                    '現場シフトが「振休」
                                        Case CATEGORYSITEREST_5
                                            .targetDate = Year(.startDay) & "/" & Replace(.wsTarget.Cells(.rowTargetShift + 1, .colTargetCalendar).Value, "分", "")
                                            .targetWeekday = Weekday(.targetDate)
                                        '振替出勤日に応じてシフト入力
                                            Select Case .targetWeekday
                                            '振替出勤日が日曜の時「日振」
                                                Case 1
                                                    .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_5
                                            '振り返出勤日が日曜以外の時「土祝振」
                                                Case Else
                                                    .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_6
                                            End Select
                                        Case Else
                                        '現場シフトにグレーハッチングがかかっていた時「休」
                                            If .wsTarget.Cells(.rowTargetShift, .colTargetCalendar).DisplayFormat.Interior.Color = RGB(191, 191, 191) Then
                                                .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYREST_1
                                        'グレーハッチング以外の時「昼」を含むシフトが入力されていない時「昼8」
                                            ElseIf Not (.wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value Like "*昼*") Then
                                                .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = CATEGORYWORK_5
                                            End If
                                    End Select
                            End Select
                            Exit For
                    End Select
                Next .colTargetCalendar
        End Select
    End With
End Function

'**********
'配列格納
'**********
'休日区分
Public Function inputArrayClassWorkday()
    With wbInfo
        If Not .classWorkDay.exists(.targetDay) Then
        '条件に応じて各日の休暇区分を配列に格納
            Select Case .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Interior.Color
                Case RGB(255, 204, 255)
                    Select Case Weekday(.wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value)
                    '日曜のとき「社休(日曜)」
                        Case 1
                            .classWorkDay.Add .targetDay, CATEGORYHOLIDAY_1
                    '土・祝のとき「社休(土祝)」
                        Case Else
                            .classWorkDay.Add .targetDay, CATEGORYHOLIDAY_2
                    End Select
            '平日のとき「平日」
                Case Else
                    .classWorkDay.Add .targetDay, CATEGORYWEEKDAY
            End Select
        End If
    End With
End Function

'シフト
Public Function inputArrayClassWorkType()
    With wbInfo
        If Not .classWorkType.exists(.targetDay) Then
            For .rowConditionDay = .rowConditionStartDay To .rowConditionEndDay Step 2
                For .colConditionDay = .colConditionStartDay To .colConditionEndDay
                    If .wsCondition.Cells(.rowConditionDay, .colConditionDay).Value = .targetDay Then
                        Select Case .wsCondition.Cells(.rowConditionDay + 1, .colConditionDay).Value
                            Case "休"
                                .classWorkType.Add .targetDay, CATEGORYREST_1
                            Case 0
                                Select Case Val(.wsCondition.Cells(.rowConditionDay + 1, .colConditionWorkTimePerday).Value)
                                    '昼4
                                    Case 4
                                        .classWorkType.Add .targetDay, CATEGORYWORK_1
                                    '昼6
                                    Case 6
                                        .classWorkType.Add .targetDay, CATEGORYWORK_2
                                    '昼7
                                    Case 7
                                        .classWorkType.Add .targetDay, CATEGORYWORK_3
                                    '昼7.75
                                    Case 7.75
                                        .classWorkType.Add .targetDay, CATEGORYWORK_4
                                    '昼8
                                    Case 8
                                        .classWorkType.Add .targetDay, CATEGORYWORK_5
                                    '昼8.5
                                    Case 8.5
                                        .classWorkType.Add .targetDay, CATEGORYWORK_6
                                    '昼9
                                    Case 9
                                        .classWorkType.Add .targetDay, CATEGORYWORK_7
                                    '昼9.5
                                    Case 9.5
                                        .classWorkType.Add .targetDay, CATEGORYWORK_8
                                    '昼10
                                    Case 10
                                        .classWorkType.Add .targetDay, CATEGORYWORK_9
                                End Select
                            Case Else
                                Select Case Val(.wsCondition.Cells(.rowConditionDay + 1, .colCheckWorkDay).Value)
                                    '昼4
                                    Case 4
                                        .classWorkType.Add .targetDay, CATEGORYWORK_1
                                    '昼6
                                    Case 6
                                        .classWorkType.Add .targetDay, CATEGORYWORK_2
                                    '昼7
                                    Case 7
                                        .classWorkType.Add .targetDay, CATEGORYWORK_3
                                    '昼7.75
                                    Case 7.75
                                        .classWorkType.Add .targetDay, CATEGORYWORK_4
                                    '昼8
                                    Case 8
                                        .classWorkType.Add .targetDay, CATEGORYWORK_5
                                    '昼8.5
                                    Case 8.5
                                        .classWorkType.Add .targetDay, CATEGORYWORK_6
                                    '昼9
                                    Case 9
                                        .classWorkType.Add .targetDay, CATEGORYWORK_7
                                    '昼9.5
                                    Case 9.5
                                        .classWorkType.Add .targetDay, CATEGORYWORK_8
                                    '昼10
                                    Case 10
                                        .classWorkType.Add .targetDay, CATEGORYWORK_9
                                End Select
                        End Select
                    End If
                Next .colConditionDay
                
                If .classWorkType.exists(.targetDay) Then
                    Exit For
                End If
            Next .rowConditionDay
        End If
    End With
End Function

'転記（1か月_日勤）
Public Function inputArrayTranslateMonthDay()
    With wbInfo
        Select Case Val(.wsShift.Cells(.rowShiftRole + 1, .colShiftCalendar).Value)
            '休日
            Case 0
                Select Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value
                    Case "休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_1
                    Case "土祝振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_2
                    Case "日振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_3
                    Case "代"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_4
                    Case "年休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_5
                    Case "明"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_6
                End Select
            '月変4定
            Case 4
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M1
            '月変6定
            Case 6
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M2
            '月変7定
            Case 7
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M3
            '月変7.75定
            Case 7.75
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M4
            '月変8定
            Case 8
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M5
            '月変8.5定
            Case 8.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M6
            '月変9定
            Case 9
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M7
            '月変9.5定
            Case 9.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M8
            '月変10定
            Case 10
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M9
            '月変11定
            Case 11
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M10
            '月変12定
            Case 12
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M11
            '月変16定
            Case 16
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_M12
        End Select
    End With
End Function

'転記（1か月_夜勤）
Public Function inputArrayTranslateMonthNight()
    With wbInfo
        Select Case Val(.wsShift.Cells(.rowShiftRole + 1, .colShiftCalendar).Value)
            '休日
            Case 0
                Select Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value
                    Case "休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_1
                    Case "土祝振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_2
                    Case "日振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_3
                    Case "代"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_4
                    Case "年休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_5
                    Case "明"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_6
                End Select
            '月変4非
            Case 4
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M1
            '月変6非
            Case 6
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M2
            '月変7非
            Case 7
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M3
            '月変7.75非
            Case 7.75
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M4
            '月変8非
            Case 8
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M5
            '月変8.5非
            Case 8.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M6
            '月変9非
            Case 9
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M7
            '月変9.5非
            Case 9.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M8
            '月変10非
            Case 10
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M9
            '月変11非
            Case 11
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M10
            '月変12非
            Case 12
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M11
            '月変16非
            Case 16
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_M12
        End Select
    End With
End Function

'転記（1年_日勤）
Public Function inputArrayTranslateYearDay()
    With wbInfo
        Select Case Val(.wsShift.Cells(.rowShiftRole + 1, .colShiftCalendar).Value)
            '休日
            Case 0
                Select Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value
                    Case "休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_1
                    Case "土祝振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_2
                    Case "日振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_3
                    Case "代"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_4
                    Case "年休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_5
                    Case "明"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_6
                End Select
            '年変4定
            Case 4
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y1
            '年変6定
            Case 6
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y2
            '年変7定
            Case 7
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y3
            '年変7.75定
            Case 7.75
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y4
            '年変8定
            Case 8
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y5
            '年変8.5定
            Case 8.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y6
            '年変9定
            Case 9
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y7
            '年変9.5定
            Case 9.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y8
            '年変10定
            Case 10
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKDAY_Y9
        End Select
    End With
End Function

'転記（1年_夜勤）
Public Function inputArrayTranslateYearNight()
    With wbInfo
        Select Case Val(.wsShift.Cells(.rowShiftRole + 1, .colShiftCalendar).Value)
            '休日
            Case 0
                Select Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value
                    Case "休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_1
                    Case "土祝振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_2
                    Case "日振"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_3
                    Case "代"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_4
                    Case "年休"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_5
                    Case "明"
                        .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYREFLECT_6
                End Select
            '年変4非
            Case 4
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y1
            '年変6非
            Case 6
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y2
            '年変7非
            Case 7
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y3
            '年変7.75非
            Case 7.75
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y4
            '年変8非
            Case 8
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y5
            '年変8.5非
            Case 8.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y6
            '年変9非
            Case 9
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y7
            '年変9.5非
            Case 9.5
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y8
            '年変10非
            Case 10
                .classWorkTranslate.Add .shiftName & .targetDay, CATEGORYWORKNIGHT_Y9
        End Select
    End With
End Function

'**********
'情報転記
'**********
'対象者
Public Function translateTargetPerson()
    With wbInfo
        .rowExportShift = .rowExportStartShift
        
        For .rowShiftRole = .rowShiftStartRole To .rowShiftEndRole Step 2
            '通し番号
            .wsExport.Cells(.rowExportShift, .colExportShift).Value = .countExport
            '対象者名
            .wsExport.Cells(.rowExportShift, .colExportName).Value = _
                    .wsShift.Cells(.rowShiftRole, .colShiftName).Value
            '個人コードを8桁にして転記
            Select Case Len(.wsShift.Cells(.rowShiftRole, .colShiftParsonalCode).Value)
                Case 5
                    .wsExport.Cells(.rowExportShift, .colExportParsonalCode).Value = _
                            "000" & .wsShift.Cells(.rowShiftRole, .colShiftParsonalCode).Value
            End Select
            '通し番号に「+1」
            .countExport = .countExport + 1
            'リシテア転記シートの行番号を「+2」
            .rowExportShift = .rowExportShift + 2
            '最終行に「-1」を入力
            Select Case .rowShiftRole
                Case .rowShiftEndRole - 1
                    .wsExport.Cells(.rowExportShift, .colExportShift).Value = -1
                    .rowExportEndShift = .rowExportShift
            End Select
        Next .rowShiftRole
    End With
End Function

'シフト→リシテア取込
Public Function translateShiftExport()
    With wbInfo
        For .rowExportShift = .rowExportStartShift To .rowExportEndShift Step 2
            For .colExportCalendar = .colExportStartCalendar To .colExportEndCalendar
                Select Case .wsExport.Cells(.rowExportCalendar, .colExportCalendar).Value
                    Case Is <> ""
                        Call modSubProcess.reflectClassTranslate
                        Call modSubProcess.checkClassHoliday
                End Select
            Next .colExportCalendar
        Next .rowExportShift
    End With
End Function

'リシテア取込Excel
Public Function translateExportExcel()
    With wbInfo
        For .rowExportShift = .rowExportStartShift To .rowExportEndShift Step 2
            '通し番号
            .wsTarget.Cells(.rowTranslateNumber, .colTranslateNumber).Value = _
                    .wsExport.Cells(.rowExportShift, .colExportShift).Value
            '対象者名
            .wsTarget.Cells(.rowTranslateNumber, .colTranslateName).Value = _
                    .wsExport.Cells(.rowExportShift, .colExportName).Value
            '個人コード
            .wsTarget.Cells(.rowTranslateNumber, .colTranslateParsonalCode).Value = _
                    .wsExport.Cells(.rowExportShift, .colExportParsonalCode).Value
            '様式Excelの行番号に「+2」
            .rowTranslateNumber = .rowTranslateNumber + 2
            '最終行に「-1」を入力
            Select Case .rowExportShift
                Case .rowExportEndShift - 1
                    .wsTarget.Cells(.rowTranslateNumber + 1, .colTranslateNumber).Value = -1
            End Select
        Next .rowExportShift
        'シフトのコピー
        .wsExport.Range(.wsExport.Cells(.rowExportStartShift, .colExportStartCalendar), _
                        .wsExport.Cells(.rowExportEndShift - 1, .colExportEndCalendar)).Copy
        .wsTarget.Range("D5").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End With
End Function

'**********
'情報反映
'**********
'休日区分
Public Function reflectClassWorkday()
    With wbInfo
        .wsExport.Cells(.rowExportShift + 1, .colExportCalendar).Value = _
                .classWorkDay(.wsExport.Cells(.rowExportCalendar, .colExportCalendar).Value)
    End With
End Function

'シフト
Public Function reflectClassWorkType()
    With wbInfo
        .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value = _
                .classWorkType(.wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value)
    End With
End Function

'転記
Public Function reflectClassTranslate()
    With wbInfo
        .wsExport.Cells(.rowExportShift, .colExportCalendar).Value = _
                .classWorkTranslate(.wsExport.Cells(.rowExportShift, .colExportName).Value & _
                                    .wsExport.Cells(.rowExportCalendar, .colExportCalendar).Value)
    End With
End Function

'法定労働時間総枠
Public Function reflectionMaxTime()
    With wbInfo
        Select Case .wsShift.Range("typeSystem").Value
            Case "1か月"
                Select Case Day(DateSerial(Year(.startDay), Month(.startDay) + 1, 1) - 1)
                    Case 28
                        Range("maxTime").Value = "160"
                    Case 29
                        Range("maxTime").Value = "165.7"
                    Case 30
                        Range("maxTime").Value = "171.4"
                    Case 31
                        Range("maxTime").Value = "177.1"
                End Select
            Case "1年"
                For .rowCheckMonth = .rowCheckStartMonth To .rowCheckEndMonth
                    Select Case .wsCheck.Cells(.rowCheckMonth, .colCheckMonth).Value
                        Case .startDay
                            Range("maxTime").Value = .wsCheck.Cells(.rowCheckMonth, .colCheckWorkTime).Value
                            Range("maxDay").Value = .wsCheck.Cells(.rowCheckMonth, .colCheckWorkDay).Value
                            Exit For
                    End Select
                Next .rowCheckMonth
        End Select
    End With
End Function

'**********
'チェック
'**********
'休日
Public Function checkClassHoliday()
    With wbInfo
        'リシテア取込シートの上段が「休」下段が「平日」の時、下段を「社休（土祝）」に変更
        If .wsExport.Cells(.rowExportShift, .colExportCalendar).Value = CATEGORYREFLECT_1 And _
                .wsExport.Cells(.rowExportShift + 1, .colExportCalendar) = CATEGORYWEEKDAY Then
            .wsExport.Cells(.rowExportShift + 1, .colExportCalendar).Value = CATEGORYHOLIDAY_2
        Else
            Select Case .wsExport.Cells(.rowExportShift, .colExportCalendar).Value
                'リシテア取込シートの上段が「代」「年休」「明」の時、上段のみ「変形休」に変更
                Case CATEGORYREFLECT_4, CATEGORYREFLECT_5, CATEGORYREFLECT_6
                    .wsExport.Cells(.rowExportShift, .colExportCalendar).Value = CATEGORYREFLECT_1
            End Select
        End If
    End With
End Function

'作成対象日
Public Function checkCreateDay()
    With wbInfo
        Select Case Weekday(Date)
            Case 1
                .createDay = Date
            Case 2
                .createDay = Date + 6
            Case 3
                .createDay = Date + 5
            Case 4
                .createDay = Date + 4
            Case 5
                .createDay = Date + 3
            Case 6
                .createDay = Date + 2
            Case 7
                .createDay = Date + 1
        End Select
    End With
End Function

'**********
'フラグ設定
'**********
'変形労働適用単位
Public Function setFlagTypeSystem()
    With wbInfo
        Select Case Range("typeSystem").Value
            Case "1か月"
                .flagTypeSystem = 0
            Case "1年"
                .flagTypeSystem = 1
        End Select
    End With
End Function

'作成可否
Public Function setFlagTypeCreate()
    With wbInfo
        .flagTypeCreate = MsgBox("新規予定表を作成します｡", vbOKCancel)
    End With
End Function

'作成対象日
Public Function setFlagCreateDay()
    With wbInfo
        Select Case .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value
            Case Is > .createDay
                .flagCreateDay = 1
            Case Else
                .flagCreateDay = 0
        End Select
    End With
End Function

'**********
'書式設定
'**********
'条件付き書式
Public Function setFormatCondition()
    With wbInfo
        With .wsShift.Range(.wsShift.Cells(.rowShiftStartRole, .colShiftWorkDay), _
                                .wsShift.Cells(.rowShiftEndRole, .colShiftWorkDay))
            '労働日数
            Set wbInfo.formatConditionWorkday = .FormatConditions.Add _
                                                (Type:=xlExpression, _
                                                    Formula1:=FC_WORKDAY)
            wbInfo.formatConditionWorkday.Interior.Color = RGB(255, 200, 255)
            wbInfo.formatConditionWorkday.StopIfTrue = False
            '休日日数
            Set wbInfo.formatConditionHoliday = .FormatConditions.Add _
                                                (Type:=xlExpression, _
                                                    Formula1:=FC_HOLIDAY)
            wbInfo.formatConditionHoliday.Interior.Color = RGB(255, 200, 255)
            wbInfo.formatConditionHoliday.StopIfTrue = False
        End With
         
        With .wsShift.Range(.wsShift.Cells(.rowShiftStartRole, .colShiftWorkTime), _
                                .wsShift.Cells(.rowShiftEndRole, .colShiftWorkTime))
            '労働時間
            Set wbInfo.formatConditionWorkTime = .FormatConditions.Add _
                                                (Type:=xlExpression, _
                                                    Formula1:=FC_WORKTIME)
            wbInfo.formatConditionWorkTime.Interior.Color = RGB(255, 200, 255)
            wbInfo.formatConditionWorkTime.StopIfTrue = False
        End With
        
        For .rowShiftRole = .rowShiftStartRole To .rowShiftEndRole Step 2
            With .wsShift.Range(.wsShift.Cells(.rowShiftRole, .colShiftStartCalendar), _
                                    .wsShift.Cells(.rowShiftRole + 1, .colShiftEndCalendar))
                '夜勤
                Set wbInfo.formatConditionShiftNight = .FormatConditions.Add _
                                                        (Type:=xlExpression, _
                                                            Formula1:=FC_SHIFT_NIGHT_1 & wbInfo.rowMasterEndShiftRole & _
                                                                        FC_SHIFT_NIGHT_2 & wbInfo.rowMasterEndShiftRole & _
                                                                        FC_SHIFT_NIGHT_3 & wbInfo.rowShiftRole & ")")
                wbInfo.formatConditionShiftNight.Interior.Color = RGB(200, 200, 255)
                wbInfo.formatConditionShiftNight.StopIfTrue = False
                '通し
                Set wbInfo.formatConditionShiftOver = .FormatConditions.Add _
                                                        (Type:=xlExpression, _
                                                            Formula1:=FC_SHIFT_OVER_1 & wbInfo.rowMasterEndShiftRole & _
                                                                        FC_SHIFT_OVER_2 & wbInfo.rowMasterEndShiftRole & _
                                                                        FC_SHIFT_OVER_3 & wbInfo.rowShiftRole & ")")
                wbInfo.formatConditionShiftOver.Interior.Color = RGB(255, 230, 210)
                wbInfo.formatConditionShiftOver.StopIfTrue = False
                '休日
                Set wbInfo.formatConditionShiftHoli = .FormatConditions.Add _
                                                        (Type:=xlExpression, _
                                                            Formula1:=FC_SHIFT_HOLIDAY_1 & wbInfo.rowMasterEndShiftRole & _
                                                                        FC_SHIFT_HOLIDAY_2 & wbInfo.rowMasterEndShiftRole & _
                                                                        FC_SHIFT_HOLIDAY_3 & wbInfo.rowShiftRole & ")")
                wbInfo.formatConditionShiftHoli.Interior.Color = RGB(220, 220, 220)
                wbInfo.formatConditionShiftHoli.StopIfTrue = False
            End With
        Next .rowShiftRole
    End With
End Function

'**********
'情報削除
'**********
'リシテア取込シート
Public Function clearExportLysithea()
    With wbInfo
        For .rowExportShift = .rowExportStartShift To .rowExportEndShift Step 2
            .wsExport.Cells(.rowExportShift, .colExportShift).MergeArea.ClearContents
            .wsExport.Cells(.rowExportShift, .colExportName).MergeArea.ClearContents
            .wsExport.Cells(.rowExportShift, .colExportParsonalCode).MergeArea.ClearContents
        Next .rowExportShift
        
        .wsExport.Range(.wsExport.Cells(.rowExportStartShift, .colExportStartCalendar), _
                        .wsExport.Cells(Rows.Count, .colExportEndCalendar)).ClearContents
    End With
End Function

'シフト表シート
Public Function clearShift()
    With wbInfo
        For .rowShiftRole = .rowShiftStartRole To .rowShiftEndRole Step 2
            .wsShift.Range(.wsShift.Cells(.rowShiftRole, .colShiftStartCalendar), _
                            .wsShift.Cells(.rowShiftRole, .colShiftEndCalendar)).ClearContents
        Next .rowShiftRole
    End With
End Function

'日付情報
Public Function clearCalendar()
    With wbInfo
        .wsExport.Range(.wsExport.Cells(.rowExportCalendarWeekday, .colShiftStartCalendar), _
                        .wsExport.Cells(.rowExportCalendar, .colExportEndCalendar)).ClearContents
        .wsExport.Range(.wsExport.Cells(.rowExportCalendarWeekday, .colShiftStartCalendar), _
                        .wsExport.Cells(.rowExportCalendar, .colExportEndCalendar)).Interior.ColorIndex = xlNone
        .wsShift.Range(.wsShift.Cells(.rowShiftCalendar, .colShiftStartCalendar), _
                        .wsShift.Cells(.rowShiftCalendarWeekday, .colShiftEndCalendar)).ClearContents
        .wsShift.Range(.wsShift.Cells(.rowShiftCalendar, .colShiftStartCalendar), _
                        .wsShift.Cells(.rowShiftCalendarWeekday, .colShiftEndCalendar)).Interior.ColorIndex = xlNone
    End With
End Function

'法定労働時間総枠
Public Function clearMaxTime()
    With wbInfo
        Range("maxTime").ClearContents
        Range("maxDay").ClearContents
    End With
End Function

'条件付き書式
Public Function clearFormatCondition()
    With wbInfo
        .wsShift.Cells.FormatConditions.Delete
    End With
End Function

'**********
'配列解放
'**********
'休日区分
Public Function deleteArrayWorkType()
    With wbInfo
        Set .classWorkType = Nothing
    End With
End Function

'シフト
Public Function deleteArrayWorkDay()
    With wbInfo
        Set .classWorkDay = Nothing
    End With
End Function

'転記
Public Function deleteArrayWorkTranslate()
    With wbInfo
        Set .classWorkTranslate = Nothing
    End With
End Function
