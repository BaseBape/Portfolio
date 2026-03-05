Attribute VB_Name = "modGetProcess"
Option Explicit

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

Public Function HolidayGet()
    With wbInfo
    'ダウンロードするファイルのURLを指定（内閣府の祝日情報）
        .targetURL = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
    '保存するファイルパスを指定
        .activePath = ActiveWorkbook.Path
        .activePath = .activePath & "\syukujitsu.csv"
    'ダウンロードを実行
        .resultDownload = URLDownloadToFile(0, .targetURL, .activePath, 0, 0)
        
        Select Case .resultDownload
            Case 0
            '祝日一覧シートのテーブル内のデータを削除
                .wsHoliday.Range("B2").ListObject.DataBodyRange.ClearContents
            '祝日情報csvを開く
                Workbooks.Open filename:=.activePath
            '変数を設定
                Set .wbHoliday = Workbooks("syukujitsu.csv")
                Set .wsHolidayInput = .wbHoliday.Sheets("syukujitsu")
            '祝日情報csvの全データをコピー
                .wsHolidayInput.Range("A2").CurrentRegion.Copy
            '祝日一覧シートに貼付け
                .wsHoliday.Range("B2").PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
            '祝日情報csvを保存せず閉じる
                .wbHoliday.Saved = True
                .wbHoliday.Close
            '祝日情報csvを削除
                CreateObject("Scripting.FileSystemObject").DeleteFile .activePath
            '祝日一覧シートの社休日をコピー
                .wsHoliday.Range(.wsHoliday.Cells(.rowHolidayListStart, .colHolidayListStart), _
                                .wsHoliday.Cells(.rowHolidayListEnd, .colHolidayListEnd)).Copy
            '祝日一覧の下に貼付け
                .wsHoliday.Range("B2").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
                .wsShift.Select
            Case Else
                MsgBox "ダウンロードできませんでした。" & vbCrLf & "再度実行してください。", vbCritical
        End Select
    End With
End Function
