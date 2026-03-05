Attribute VB_Name = "modCommonProcess"
Option Explicit

'処理開始
Public Function comFuncInitMotionOff()
    With Excel.Application
      .ScreenUpdating = False
      .Cursor = xlWait
      .EnableEvents = False
      .DisplayAlerts = False
      .Calculation = xlCalculationManual
    End With
End Function

'処理終了
Public Function comFuncInitMotionOn()
    With Excel.Application
        .StatusBar = False
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .EnableEvents = True
        .Cursor = xlDefault
        .ScreenUpdating = True
    End With
End Function

'URLデコード
Public Function decodeURL()
    With CreateObject("ScriptControl")
        .Language = "JScript"
        wbInfo.decodeURL = .CodeObject.decodeURI(wbInfo.targetFilePath)
    End With
    wbInfo.decodeURL = Replace(wbInfo.decodeURL, "?web=1", "")
End Function

'行削除
Public Function comFuncDeleteRow()
    With wbInfo
        .wsMain.Range("2:" & Rows.Count).Delete
    End With
End Function

'ファイルパス選択ダイアログ（csvファイル）
Public Function funcSlectCSVFile()
    With wbInfo
    'カレントディレクトリをデスクトップに変更
        Set .wScriptHost = CreateObject("WScript.Shell")
        ChDir .wScriptHost.SpecialFolders("Desktop")
        
        .myFile = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")
        
        If Not VarType(.myFile) = vbBoolean Then
            Range("oplusFilePath").Value = .myFile
        End If
    End With
End Function

'ファイルパス選択ダイアログ（Excelファイル）
Public Function funcSlectExcelFile()
    With wbInfo
    'カレントディレクトリをデスクトップに変更
        Set .wScriptHost = CreateObject("WScript.Shell")
        ChDir .wScriptHost.SpecialFolders("Desktop")
        
        .myFile = Application.GetOpenFilename(filefilter:="Excelブック,*.xlsx")
        
        If Not VarType(.myFile) = vbBoolean Then
            Range("saveFilePath").Value = .myFile
        End If
    End With
End Function

'ファイルパス選択ダイアログ（マクロ有効化ブック）
Public Function funcSlectExcelVBA()
    With wbInfo
    'カレントディレクトリをデスクトップに変更
        Set .wScriptHost = CreateObject("WScript.Shell")
        ChDir .wScriptHost.SpecialFolders("Desktop")
        
        .myFile = Application.GetOpenFilename(filefilter:="Excelマクロブック,*.xlsm")
        
        If Not VarType(.myFile) = vbBoolean Then
            Range("saveFilePath").Value = .myFile
        End If
    End With
End Function

'ファイルパス選択ダイアログ（すべてのファイル）
Public Function funcSlectAllFile()
    With wbInfo
    'カレントディレクトリをデスクトップに変更
        Set .wScriptHost = CreateObject("WScript.Shell")
        ChDir .wScriptHost.SpecialFolders("Desktop")
        
        .myFile = Application.GetOpenFilename("すべてのファイル(*.*),*.*")
        
        If Not VarType(.myFile) = vbBoolean Then
            Range("saveFilePath").Value = .myFile
        End If
    End With
End Function

'ファイルパス選択ダイアログ（インポートフォルダ）
Public Function comFuncImportFolderSelect()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "インポートフォルダの選択"
        .InitialFileName = "c=\"
        If .Show = True Then
            wbInfo.importFolderPath = .SelectedItems(1)
        End If
    End With
End Function

'ファイルパス選択ダイアログ（エクスポートフォルダ）
Public Function comFuncExportFolderSelect()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "エクスポートフォルダの選択"
        .InitialFileName = "c=\"
        If .Show = True Then
            wbInfo.exportFolderPath = .SelectedItems(1)
        End If
    End With
End Function

'メッセージボックス
Public Static Sub sub_msg_error(ByVal in_msg_perttern As String, _
                                ByVal in_msg_str As String, _
                                Optional ByVal in_option_str As String)
    
    Select Case in_msg_perttern
        Case 1
             MsgBox in_msg_str & "が空白です。" & vbLf & "処理を中断します。", vbCritical, "処理中断"
             End
        Case 2
             MsgBox in_msg_str & "が空白です。" & vbLf & in_option_str & vbLf & "処理を中断します。", vbCritical, "処理中断"
             End
        Case 3
             MsgBox in_msg_str & "が存在しません。" & vbLf & in_option_str & vbLf & "処理を中断します。", vbCritical, "処理中断"
             End
        Case 4
             MsgBox in_msg_str & "です。" & vbLf & in_option_str & vbLf & "処理を中断します。", vbCritical, "処理中断"
             End
        Case 5
             MsgBox in_msg_str & "恐れがあります。" & vbLf & in_option_str & vbLf & "処理を中断します。", vbCritical, "処理中断"
             End
    End Select
    
End Sub


'Public Static Sub sub_msg_start(ByVal in_target_process_name As String, _
'                                ByRef sheet_info As sheetinfo)
'
'    Dim Message As Long
'
'    Select Case sheet_info.process_all0_or_single1_or_test2
'        Case 0
'            Message = MsgBox("一括メール送信【" & sheet_info.sheetname & "】" & in_target_process_name & "を開始します。" & vbLf & _
'                     "対象件数：" & sheet_info.count_target & "件" & vbLf & _
'                     "送信対象顧客数：" & sheet_info.count_crient & "社" & vbLf & _
'                     "送信メール数：" & sheet_info.count_target * sheet_info.count_crient & "件", _
'                     vbOKCancel + vbInformation, "処理実行確認")
'
'            If Not Message = vbOK Then
'                MsgBox "処理を中断します。", vbInformation, "処理中断"
'                End
'            End If
'
'
'        Case 1
'            Message = MsgBox("メール送信（１社のみ）" & in_target_process_name & "を開始します。" & vbLf & _
'                                  "対象件数：" & sheet_info.count_target & "件", _
'                                 vbOKCancel + vbInformation, "処理実行確認")
'
'            If Not Message = vbOK Then
'                MsgBox "処理を中断します。", vbInformation, "処理中断"
'                End
'            End If
'
'        Case 2
'            Message = MsgBox("メール送信（テスト用）" & in_target_process_name & "を開始します。" & vbLf & _
'                                  "対象件数：" & sheet_info.count_target & "件", _
'                                 vbOKCancel + vbInformation, "処理実行確認")
'
'            If Not Message = vbOK Then
'                MsgBox "処理を中断します。", vbInformation, "処理中断"
'                End
'            End If
'
'    End Select
'
'End Sub


'--------------------------------------------------------------------------
'ステータスバー
'--------------------------------------------------------------------------
'DoEvents
'sheet_info.wb.Activate
'Application.StatusBar = sheet_info.sheetname & "【処理件数】" & i - 3 & "/" & sheet_info.lastrow_target - 3 & "件　"

'--------------------------------------------------------------------------
'ファイル操作
'--------------------------------------------------------------------------
'ファイル（パス）が存在するか
'    For Each target_filepath In mail_info.attachment_split
'        If Not sheet_info.fso.FileExists(Trim()) Then
'
'        End If
'    Next target_filepath


'ファイル削除
Public Function com_func_delete_file_templatefile()

     Dim fso           As FileSystemObject
     Dim filePath      As String, filename As String, wb As Workbook
     Dim flg           As Boolean
     
     filename = "import.xlsm"
     filePath = ThisWorkbook.Path & "\" & filename
     
     Set fso = New FileSystemObject
     
     For Each wb In Workbooks
        If wb.Name = filename Then
            flg = True
            Exit For
        End If
     Next

     If flg = True Then
          Workbooks(filename).Close
     End If
     
     If (fso.FileExists(filePath) = True) Then
        fso.GetFile(filePath).Delete
     End If
     
     Set fso = Nothing
     
End Function

'ファイルを探す
Public Function com_func_serch_file_harutaka_export()

    Dim fso               As FileSystemObject
    Dim folder_list       As Folder
    Dim filepath_file     As file
    Dim filename          As String
    Dim folderpath_target As String
    Dim filename_target   As String
    Dim wb                As Workbook
    Dim flg               As Boolean
     
    flg = False
    
    folderpath_target = "C:\Users\本社貸し出し用①\Downloads\"
    filename_target = "Candidates"
    
    Set fso = New FileSystemObject
    Set folder_list = fso.GetFolder(folderpath_target)
    
    For Each filepath_file In folder_list.Files
        If Left(filepath_file.Name, 10) = filename_target Then
            For Each wb In Workbooks
                If wb.Name = filepath_file.Name Then
                    flg = True
                    Exit For
                End If
            Next

            If flg = True Then
                Workbooks(filepath_file.Name).Close
            End If
            com_func_serch_file_harutaka_export = filepath_file
        End If
    Next
        
End Function

'電話番号ハイフンチェック
Public Function com_func_telnumber_check(ByVal telnumber As String) As String
    Dim c         As Long
    Dim serch_count As Long
    Dim serch_str As String
    c = 0
    serch_str = "-"
    
    If Not Trim(telnumber) = "" Then
        If Len(telnumber) >= 10 Then
            telnumber = StrConv(telnumber, vbNarrow)
            
            If Left(telnumber, 1) <> 0 Then
                telnumber = "0" & telnumber
            End If
            
            'ハイフンカウント
            For c = 1 To Len(telnumber)
                If Mid(telnumber, c, 1) = serch_str Then
                    serch_count = serch_count + 1
                End If
            Next c
            
            Select Case serch_count
              Case 0
                telnumber = Left(telnumber, Len(telnumber) - 4) & "-" & Right(telnumber, 4)
                telnumber = Left(telnumber, Len(telnumber) - 9) & "-" & Right(telnumber, 9)
              Case 1
                telnumber = Replace(telnumber, "-", "")
                telnumber = Left(telnumber, Len(telnumber) - 4) & "-" & Right(telnumber, 4)
                telnumber = Left(telnumber, Len(telnumber) - 9) & "-" & Right(telnumber, 9)
            End Select
            
        End If
    End If
     
    com_func_telnumber_check = telnumber

End Function

 'ワークシート作成
Public Static Sub create_worksheet()
    Dim ws_export_book As Worksheet
    Dim NewWorkSheet   As Worksheet
    Dim lastrow        As Long
    Dim i              As Long
    
    lastrow = ThisWorkbook.Worksheets("部門リスト").Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastrow
        Worksheets("temp").Select
        Worksheets("temp").Copy after:=ThisWorkbook.Worksheets(Sheets.Count)
        ActiveSheet.Name = ThisWorkbook.Worksheets("部門リスト").Cells(i, 1).Value
    Next i
End Sub

'CSVとして保存
Public Function savefile_csv(ByVal in_sheetname As String, ByVal in_dirpath_save As String) As String
    Dim dblTimer  As Double
    Dim s_return  As String
    Dim getMSec   As String
    Dim nowtime   As String
    Dim filePath  As String
    Dim csvFile As String
    Dim filepath_savepath As String
    
    nowtime = Format(Now, "yyyymmddHHMMSS")
    filepath_savepath = in_dirpath_save
    
    filePath = filepath_savepath & nowtime & "_" & "rpm_import" & "_" & in_sheetname & ".csv"
    
    With ThisWorkbook
        .Worksheets(in_sheetname).Activate
        .Worksheets(in_sheetname).Copy
    End With
    
    If Dir(filepath_savepath, vbDirectory) = "" Then
        MkDir filepath_savepath
    End If
    
    With ActiveWorkbook
        .SaveAs filename:=filePath, FileFormat:=xlCSV, Local:=True
        .Close savechanges:=False
    End With
    
    savefile_csv = filePath
    
End Function

'開いているファイルを閉じる
Public Function com_func_delete_file_se_importfile(ByVal branch As String)
     Dim fso           As FileSystemObject
     Dim filePath      As String
     Dim wb            As Workbook
     Dim filename      As String
     Dim flg           As Boolean
     
     filename = "se_import.xlsx"
     filePath = "" & branch & "\" & filename
     
     Set fso = New FileSystemObject
     
     For Each wb In Workbooks
        If wb.Name = filename Then
            flg = True
            Exit For
        End If
     Next

     If flg = True Then
          Workbooks(filename).Close
     End If
     
     If (fso.FileExists(filePath) = True) Then
        fso.GetFile(filePath).Delete
     End If
     
     Set fso = Nothing
End Function

'別ファイルとして保存する
Public Sub saveFileOther(ByVal in_sheetname As String)
    Dim dblTimer  As Double
    Dim s_return  As String
    Dim getMSec   As String
    Dim nowtime   As String
    Dim filePath  As String
    Dim csvFile As String
    Dim filepath_savepath As String
'    Dim Path As String, WSH As Variant
'
'    Set WSH = CreateObject("WScript.Shell")
'    Path = WSH.SpecialFolders("Desktop") & "\"
'
'    nowtime = Format(Now, "yyyymmddHHMMSS")
'    filepath_savepath =  in_sheetname & "\"
'
'    Filepath = filepath_savepath & nowtime & "_" & in_sheetname & ".xlsx"
    
    With ThisWorkbook
        .Worksheets(in_sheetname).Activate
         .Worksheets(in_sheetname).Copy
        Set wbInfo.wbOutput = ActiveWorkbook
          
    End With
    
'    If Dir(filepath_savepath, vbDirectory) = "" Then
'        MkDir filepath_savepath
'    End If
    
'    With ActiveWorkbook
        '.SaveAs filename:=Filepath, FileFormat:=xlWorkbookDefault
        '.Close SaveChanges:=False
'    End With
    
End Sub
