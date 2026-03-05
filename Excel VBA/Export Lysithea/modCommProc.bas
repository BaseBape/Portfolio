Attribute VB_Name = "modCommonProcess"
Option Explicit

'画面更新OFF
Public Function comFuncInitMotionOff()
    With Excel.Application
      .ScreenUpdating = False
      .Cursor = xlWait
      .EnableEvents = False
      .DisplayAlerts = False
      .Calculation = xlCalculationManual
    End With
End Function

'画面更新ON
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

'ファイルを開く
Public Function fileOpen()
    With wbInfo
        Workbooks.Open filename:=.decodeURL
    End With
End Function

'ファイル選択（リシテア取込Excel）
Public Function selectExportFile()
    With wbInfo
        .myFile = Application.GetOpenFilename("すべてのファイル(*.*),*.*")
        
        Select Case .myFile
            Case "False"
                .flagSelectFile = 1
            Case Else
                .flagSelectFile = 0
                Workbooks.Open .myFile
                
                Set .wbTarget = ActiveWorkbook
                Set .wsTarget = .wbTarget.Worksheets(SHEETNAMETARGET)
        End Select
    End With
End Function

