@echo off
setlocal enabledelayedexpansion

echo エクセルシートをPDF化するバッチファイルを実行します
echo.
echo 同じフォルダ内のすべてのエクセルファイルを処理します...
echo.

set outputfolder=%~dp0
echo PDF保存先: %outputfolder%
echo.

echo 処理を開始します...

for %%f in (*.xls *.xlsx *.xlsm) do (
    if not exist "%%~nf.lock" (
        echo ファイル「%%f」を処理中...
        cscript //nologo "%~dp0ExcelToPDF.vbs" "%~dp0%%f" "%outputfolder%"
        echo.> "%%~nf.lock"
    )
)

rem 処理後にロックファイルを削除
del *.lock

echo.
echo すべての処理が完了しました！
pause