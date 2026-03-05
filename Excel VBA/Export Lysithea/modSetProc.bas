Attribute VB_Name = "modSetProcess"
Option Explicit

'ワークシート変数の格納
Public Function funcInfoSet_Shift()
    With wbInfo
        Set .wb = ThisWorkbook
        Set .wsShift = .wb.Worksheets(SHEETNAMESHIFT)
        Set .wsExport = .wb.Worksheets(SHEETNAMEEXPORT)
        Set .wsMaster = .wb.Worksheets(SHEETNAMEMASTER)
        Set .wsCondition = .wb.Worksheets(SHEETNAMECONDITION)
        Set .wsCheck = .wb.Worksheets(SHEETNAMECHECK)
        Set .wsHoliday = .wb.Worksheets(SHEETNAMEHOLIDAY)
        Set .classWorkDay = CreateObject("Scripting.Dictionary")
        Set .classWorkType = CreateObject("Scripting.Dictionary")
        Set .classWorkTranslate = CreateObject("Scripting.Dictionary")
        
        .targetFilePath = Range("targetFilePath").Value
        .startDay = Range("startDay").Value
        .targetMonth = Year(.startDay) & "." & Month(.startDay)
        
    '*****************
    'リシテア取込Excel
    '*****************
    '項番
        .rowTranslateNumber = 5
        .colTranslateNumber = 1
    '氏名
        .colTranslateName = 2
    '個人CD
        .colTranslateParsonalCode = 3
        
    '*****************
    'リシテア取込
    '*****************
    'シフト
        .rowExportStartShift = Range("exportStartShift").Row + 2
        .rowExportEndShift = .wsExport.Cells(Rows.Count, 1).End(xlUp).Row
        .colExportShift = Range("exportStartShift").Column
    'カレンダー
        .rowExportCalendar = Range("exportStartCalendar").Row
        .rowExportCalendarWeekday = .rowExportCalendar - 1
        .colExportStartCalendar = Range("exportStartCalendar").Column
        .colExportEndCalendar = Range("exportEndCalendar").Column
    '項番
        .countExport = 1
    '氏名
        .colExportName = Range("exportName").Column
    '個人CD
        .colExportParsonalCode = Range("exportParsonalCode").Column
        
    '*****************
    'シフト表
    '*****************
    'シフト
        .rowShiftStartRole = Range("shiftName").Row + 3
        .rowShiftEndRole = .wsShift.Cells(Rows.Count, 2).End(xlUp).Row + 1
    'カレンダー
        .rowShiftCalendar = Range("shiftStartCalendar").Row
        .rowShiftCalendarWeekday = .rowShiftCalendar + 1
        .colShiftStartCalendar = Range("shiftStartCalendar").Column
        .colShiftEndCalendar = Range("shiftEndCalendar").Column
        .colEndCalendar = Day(DateSerial(Year(.startDay), Month(.startDay) + 1, 1) - 1) + .colShiftStartCalendar - 1
    '氏名
        .colShiftName = Range("shiftName").Column
    '個人CD
        .colShiftParsonalCode = Range("shiftParsonalCode").Column
    '労働時間
        .colShiftWorkTime = Range("shiftWorkTime").Column
    '労働日数
        .colShiftWorkDay = Range("shiftWorkDay").Column
        
    '*****************
    '条件
    '*****************
    '対象日
        .colConditionStartDay = Range("conditionStartShift").Column
        .colConditionEndDay = Range("conditionEndShift").Column
        .rowConditionStartDay = Range("conditionStartShift").Row
        .rowConditionEndDay = .wsCondition.Cells(Rows.Count, .colConditionStartDay).End(xlUp).Row
    '所定労働時間（1日あたり）
        .colConditionWorkTimePerday = Range("conditionWorkTimePerday").Column
        
    '*****************
    '労働時間チェック
    '*****************
    '対象月
        .rowCheckStartMonth = Range("checkStartCalendar").Row
        .rowCheckEndMonth = Range("checkEndCalendar").Row
        .colCheckMonth = Range("checkStartCalendar").Column
    '労働日数
        .colCheckWorkDay = Range("checkWorkDay").Column
    '所定労働時間
        .colCheckWorkTime = Range("checkWorkTime").Column
        
    '*****************
    'マスタ
    '*****************
    '内容
        .colMasterShiftContent = Range("masterShiftContent").Column
    '勤務区分凡例
        .rowMasterEndShiftRole = .wsMaster.Cells(Rows.Count, .colMasterShiftContent).End(xlUp).Row
        
    '*****************
    '祝日一覧
    '*****************
    '開始終了列（社休日）
        .colHolidayListStart = 5
        .colHolidayListEnd = 6
    '開始終了行（社休日）
        .rowHolidayListStart = 3
        .rowHolidayListEnd = .wsHoliday.Cells(Rows.Count, .colHolidayListStart).End(xlUp).Row
    End With
End Function

'検索シート
Public Function funcInfoSet_target()
    With wbInfo
        Workbooks.Open filename:=.decodeURL
        Set .wbTarget = ActiveWorkbook
        Set .wsTarget = .wbTarget.Worksheets(.targetMonth)
        
        .rowTargetStartShift = 9
        .rowTargetEndShift = .wsTarget.Cells(Rows.Count, 1).End(xlUp).Row + 1
        .colTargetStartCalendar = 5
        .colTargetEndCalendar = .wsTarget.Cells(4, 5).End(xlToRight).Column
    End With
End Function
