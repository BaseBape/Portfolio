Attribute VB_Name = "MainModule"
Option Explicit

'新規予定表の作成
Public Sub createShift()
    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet_Shift
    Call modMainProcess.createNewShift
    Call modCommonProcess.comFuncInitMotionOn
End Sub

'現場シフト表取込
Public Sub inputSiteShift()
    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet_Shift
    Call modCommonProcess.decodeURL
    Call modSetProcess.funcInfoSet_target
    Call modMainProcess.reflectSiteShift
    Call modCommonProcess.comFuncInitMotionOn
    MsgBox "処置が完了しました。", vbInformation, "シフト表転記"
End Sub

'リシテア取込Excel転記
Public Sub exportLysitheaExcel()
    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet_Shift
    Call modMainProcess.reflectExportExcel
    Call modCommonProcess.comFuncInitMotionOn
End Sub

'条件付き書式設定
Public Sub setFormatCondition()
    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet_Shift
    Call modSubProcess.clearFormatCondition
    Call modSubProcess.setFormatCondition
    Call modCommonProcess.comFuncInitMotionOn
End Sub

'祝日・社休日取得
Public Sub getHoliday()
    Call modCommonProcess.comFuncInitMotionOff
    Call modSetProcess.funcInfoSet_Shift
    Call modGetProcess.HolidayGet
    Call modCommonProcess.comFuncInitMotionOn
End Sub


