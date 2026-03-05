Attribute VB_Name = "modGetProcess"
Option Explicit

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

Sub HolidayGet()
    With wbInfo
    'ダウンロードするファイルのURLを指定
        .targetURL = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
        
    '保存するファイルパスを指定
        .saveFilePath = .wb.Path
        .saveFilePath = .saveFilePath & "\syukujitsu.csv"
    
    'ダウンロードを実行
        .resultDownload = URLDownloadToFile(0, .targetURL, .saveFilePath, 0, 0)
        
        If .resultDownload = 0 Then
            .wsHoliday.Range("B2").CurrentRegion.ClearContents
            Workbooks.Open filename:=.saveFilePath
            Workbooks("syukujitsu.csv").Sheets("syukujitsu").Range("A1").CurrentRegion.Copy
            .wsHoliday.Range("B2").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Workbooks("syukujitsu.csv").Saved = True
            Workbooks("syukujitsu.csv").Close
            CreateObject("Scripting.FileSystemObject").DeleteFile .saveFilePath
            .wsMain.Select
            
        Else
            MsgBox ("ダウンロードできませんでした。" & vbCrLf & _
                    "再度実行してください。")
            
        End If
    End With
End Sub
