Attribute VB_Name = "MainModule"
Option Explicit

'新規予定表の作成
Public Sub Create_Schedule()
    Call modMainProcess.createNewSchedule
End Sub

'手順1
Public Sub ExportsFile_Select()
    Call modCommonProcess.funcMsgStartProcess
    MsgBox "読み込むファイルパスを設定しました。", vbInformation, "oplusファイル指定"
End Sub

'手順2
Public Sub Shift_Integration()
    Call modMainProcess.Main
    MsgBox "処理が完了しました。", vbInformation, "勤務予定表作成"
End Sub

'手順3
Public Sub change_Coloring()
    Call modMainProcess.changeColoring
    MsgBox "処理が完了しました。", vbInformation, "予定着色"
End Sub

'手順4
Public Sub SaveFile_Select()
    Call modCommonProcess.funcMsgSaveProcess
    MsgBox "保存先ファイルパスを設定しました。", vbInformation, "保存ファイル指定"
End Sub

'手順5
Public Sub Book_OutPut()
    Call modMainProcess.mainDataOutput
    MsgBox "処理が完了しました。", vbInformation, "勤務予定表出力"
End Sub

'その他(シフト表クリア)
Public Sub Shift_Clear()
    Call modMainProcess.mainShiftClear
    MsgBox "処理が完了しました。", vbInformation, "シフト表クリア"
End Sub

'その他(祝日一覧の取得)
Public Sub Holiday_Get()
    Call modMainProcess.mainHolidayGet
    MsgBox "処理が完了しました。", vbInformation, "祝日一覧取得"
End Sub
