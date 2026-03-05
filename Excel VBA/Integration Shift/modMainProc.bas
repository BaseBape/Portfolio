Attribute VB_Name = "modMainProcess"
Option Explicit

'新規予定表の作成
Public Sub createNewSchedule()

    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet
    Call modSubProcess.clearCalendar
    Call modSubProcess.clearShift
    Call modSubProcess.inputSceduleThisMonth
    Call modCommonProcess.comFuncInitMotionOn
    
End Sub

'シフトデータの統合
Public Sub Main()

    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet
    Call modSubProcess.subInputData
    Call modSubProcess.changeShiftForm
    Call modSubProcess.shiftAggregation
    Call modSubProcess.changeAggregationForm
    Call modCommonProcess.comFuncInitMotionOn
    
End Sub

'予定の色付け
Public Sub changeColoring()
    
    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet
    Call modSubProcess.changeShiftColoring
    Call modSubProcess.changeWorkColoring
    Call modSubProcess.changeShiftForm
    Call modSubProcess.shiftAggregation
    Call modSubProcess.changeAggregationForm
    Call modSubProcess.roleChange
    Call modCommonProcess.comFuncInitMotionOn
    
End Sub

'シフトデータの出力
Public Sub mainDataOutput()

    If Range("saveFilePath") = "" Then
        Call modCommonProcess.comFuncInitMotionOff
        Call modSetProcess.funcInfoSet
        Call modSubProcess.dataAnotherFileOutput
        Call modCommonProcess.comFuncInitMotionOn
    Else
        Call modCommonProcess.comFuncInitMotionOff
        Call modSetProcess.funcInfoSet
        Call modSubProcess.dataOutput
        Call modCommonProcess.comFuncInitMotionOn
    End If
    
End Sub

'シフト表クリア
Public Sub mainShiftClear()

    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet
    Call modSubProcess.clearWork
    Call modSubProcess.clearShift
    Call modSubProcess.reflectionClose
    Call modCommonProcess.comFuncInitMotionOn

End Sub

'祝日一覧の取得
Public Sub mainHolidayGet()

    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet
    Call modGetProcess.HolidayGet
    Call modCommonProcess.comFuncInitMotionOn

End Sub
