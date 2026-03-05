Attribute VB_Name = "modSetProcess"
Option Explicit

'ワークシート変数の格納
Public Function funcInfoSet()
    With wbInfo
        Set .fso = CreateObject("Scripting.FileSystemObject")
        Set .wb = ThisWorkbook
        Set .wsMain = .wb.Worksheets(SHEETNAMECREATE)
        Set .wsShiftRole = .wb.Worksheets(SHEETNAMESHIFTROLE)
        Set .wsWorkRole = .wb.Worksheets(SHEETNAMEWORKROLE)
        Set .wsPlanRole = .wb.Worksheets(SHEETNAMEPLANROLE)
        Set .wsHoliday = .wb.Worksheets(SHEETNAMEHOLIDAY)

    'シフト表
        .colStartCalendar = Range("startPosition").Column
        .colEndCalendar = .colStartCalendar + 30
        .rowStartCalendar = Range("startPosition").Row
        .rowEndCalendar = Range("targetPaste").Row - 1
        .colStartShift = Range("targetPaste").Column
        .colEndShift = .colStartShift + 32
        .rowStartShift = Range("targetPaste").Row
        If .wsMain.Range("targetPaste").Value = "" Then
            .rowEndShift = .wsMain.Cells(Rows.Count, .colStartShift).End(xlUp).Row + 10
        Else
            .rowEndShift = .wsMain.Cells(Rows.Count, .colStartShift).End(xlUp).Row
        End If
        .colStartAggre = Range("targetAggregation").Column
        .colEndAggre = .colStartAggre + 4
        .rowStartAggre = Range("targetAggregation").Row
        
    'シフト凡例シート
        .lastRowShiftRole = .wsShiftRole.Cells(Rows.Count, 1).End(xlUp).Row
        .colTargetShiftRole = Range("oplusRole").Column
        .colChangeShiftRole = Range("shiftRole").Column
        .colShiftRoleTarget = Range("shiftRole").Column
        .colShiftRoleInteriorRed = Range("shiftInteriorRed").Column
        .colShiftRoleInteriorGreen = Range("shiftInteriorGreen").Column
        .colShiftRoleInteriorBlue = Range("shiftInteriorBlue").Column
        .colShiftRoleFontRed = Range("shiftFontRed").Column
        .colShiftRoleFontGreen = Range("shiftFontGreen").Column
        .colShiftRoleFontBlue = Range("shiftFontBlue").Column
        .colShiftRoleFontStyle = Range("shiftFontStyle").Column
        
    '作業内容凡例シート
        .lastRowWorkRole = .wsWorkRole.Cells(Rows.Count, 1).End(xlUp).Row
        .colWorkRoleTarget = Range("workRole").Column
        .colSearchRole = Range("workRole").Column
        .colWorkRoleInteriorRed = Range("workInteriorRed").Column
        .colWorkRoleInteriorGreen = Range("workInteriorGreen").Column
        .colWorkRoleInteriorBlue = Range("workInteriorBlue").Column
        .colWorkRoleFontRed = Range("workFontRed").Column
        .colWorkRoleFontGreen = Range("workFontGreen").Column
        .colWorkRoleFontBlue = Range("workFontBlue").Column
        .colWorkRoleFontStyle = Range("workFontStyle").Column
        
    End With
End Function

