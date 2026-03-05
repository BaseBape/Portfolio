Attribute VB_Name = "modMainProcess"
Option Explicit

'シフト表の作成
Public Function createNewShift()
    With wbInfo
    '作成可否
        Call modSubProcess.setFlagTypeCreate
        
        Select Case .flagTypeCreate
            Case 1
            '情報削除
                Call modSubProcess.clearCalendar
                Call modSubProcess.clearExportLysithea
                Call modSubProcess.clearShift
                Call modSubProcess.clearMaxTime
            '適用単位確認
                Call modSubProcess.setFlagTypeSystem
            '所定労働上限時間設定
                Call modSubProcess.reflectionMaxTime
            'カレンダー作成
                For .colShiftCalendar = .colShiftStartCalendar To .colEndCalendar
                    .targetDay = .startDay + .colShiftCalendar - .colShiftStartCalendar
                    
                    Select Case Month(.targetDay)
                        Case Month(.startDay)
                            Call modSubProcess.createCalendar
                        '適用期間が「1年」のみデフォルトシフトを配列格納
                            Select Case .flagTypeSystem
                                Case 1
                                    Call modSubProcess.inputArrayClassWorkType
                            End Select
                    End Select
                Next .colShiftCalendar
            '適用期間が「1年」のみデフォルトシフト作成
                Select Case .flagTypeSystem
                    Case 1
                        Call modSubProcess.createShiftYear
                        Call modSubProcess.deleteArrayWorkType
                End Select
            '条件付き書式設定
                Call modSubProcess.clearFormatCondition
                Call modSubProcess.setFormatCondition
        End Select
    End With
End Function

'現場シフト転記
Public Function reflectSiteShift()
    With wbInfo
    '作成対象日の設定
        Call modSubProcess.checkCreateDay
        
        For .colShiftCalendar = .colShiftStartCalendar To .colShiftEndCalendar
        '作成対象日の判定
            Call modSubProcess.setFlagCreateDay
            
            Select Case .flagCreateDay
                Case 1
                    For .rowShiftRole = .rowShiftStartRole To .rowShiftEndRole Step 2
                        For .rowTargetShift = .rowTargetStartShift To .rowTargetEndShift Step 2
                        '現場シフトの転記
                            Call modSubProcess.inputData
                        Next .rowTargetShift
                    Next .rowShiftRole
            End Select
        Next .colShiftCalendar
    End With
End Function

'リシテア取込Excelへの転記
Public Function reflectExportExcel()
    With wbInfo
    '適用単位確認
        Call modSubProcess.setFlagTypeSystem
    '対象者情報コピー
        Call modSubProcess.translateTargetPerson
    '配列格納
        For .rowShiftRole = .rowShiftStartRole To .rowShiftEndRole Step 2
            For .colShiftCalendar = .colShiftStartCalendar To .colShiftEndCalendar
                Select Case .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value
                    Case Is <> ""
                    '氏名・日付変数設定
                        .targetDay = .wsShift.Cells(.rowShiftCalendar, .colShiftCalendar).Value
                        .shiftName = .wsShift.Cells(.rowShiftRole, .colShiftName).Value
                    '休日区分設定
                        Call modSubProcess.inputArrayClassWorkday
                    '設定したシフトを配列格納
                        Select Case .flagTypeSystem
                            Case 0
                                Select Case True
                                    Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value Like "*夜*"
                                        Call modSubProcess.inputArrayTranslateMonthNight
                                    Case Not .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value Like "*夜*"
                                        Call modSubProcess.inputArrayTranslateMonthDay
                                End Select
                            Case 1
                                Select Case True
                                    Case .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value Like "*夜*"
                                        Call modSubProcess.inputArrayTranslateYearNight
                                    Case Not .wsShift.Cells(.rowShiftRole, .colShiftCalendar).Value Like "*夜*"
                                        Call modSubProcess.inputArrayTranslateYearDay
                                End Select
                        End Select
                End Select
            Next .colShiftCalendar
        Next .rowShiftRole
    '休日区分設定
        Call modSubProcess.createExportLysithea
        Call modSubProcess.deleteArrayWorkDay
    'リシテア取込シート転記
        Call modSubProcess.translateShiftExport
        Call modSubProcess.deleteArrayWorkTranslate
    '休日区分チェック
    
    '個人コードチェック
        
    'エクスポートファイルの指定
        Call modCommonProcess.selectExportFile
    'データのエクスポート
        Select Case .flagSelectFile
            Case 0
                Call modSubProcess.translateExportExcel
        End Select
    End With
End Function
