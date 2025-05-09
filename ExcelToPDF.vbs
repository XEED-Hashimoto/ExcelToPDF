' ExcelToPDF.vbs
' エクセルファイルの各シートをPDFに変換するスクリプト

' コマンドライン引数を取得
If WScript.Arguments.Count < 1 Then
    WScript.Echo "引数が不足しています。エクセルファイルのパスを指定してください。"
    WScript.Quit
End If

excelFilePath = WScript.Arguments(0)
If WScript.Arguments.Count > 1 Then
    outputFolder = WScript.Arguments(1)
Else
    outputFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(excelFilePath)
End If

' ファイル名（拡張子なし）を取得
Set fso = CreateObject("Scripting.FileSystemObject")
fileName = fso.GetBaseName(excelFilePath)

' Excelアプリケーションを起動
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Excelを起動できませんでした。Excelがインストールされているか確認してください。"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

' エクセルファイルを開く
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(excelFilePath)
If Err.Number <> 0 Then
    WScript.Echo "ファイル「" & excelFilePath & "」を開けませんでした。"
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit
End If
On Error GoTo 0

' 各シートをPDFとして保存
For Each objSheet In objWorkbook.Sheets
    sheetName = objSheet.Name
    pdfPath = outputFolder & "\" & fileName & "_" & sheetName & ".pdf"
    
    WScript.Echo "  シート「" & sheetName & "」をPDF化しています..."
    
    ' 現在のシートをアクティブにする
    objSheet.Activate
    
    ' PDFとして保存
    On Error Resume Next
    objWorkbook.ActiveSheet.ExportAsFixedFormat 0, pdfPath, 0, 1, 0, , , 0
    If Err.Number <> 0 Then
        WScript.Echo "  エラー: シート「" & sheetName & "」のPDF化に失敗しました。"
    End If
    On Error GoTo 0
Next

' エクセルを閉じる
objWorkbook.Close False
objExcel.Quit

' オブジェクトの解放
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "ファイル「" & excelFilePath & "」のすべてのシートをPDF化しました。"
