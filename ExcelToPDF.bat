@echo off
setlocal enabledelayedexpansion

echo �G�N�Z���V�[�g��PDF������o�b�`�t�@�C�������s���܂�
echo.
echo �����t�H���_���̂��ׂẴG�N�Z���t�@�C�����������܂�...
echo.

set outputfolder=%~dp0
echo PDF�ۑ���: %outputfolder%
echo.

echo �������J�n���܂�...

for %%f in (*.xls *.xlsx *.xlsm) do (
    if not exist "%%~nf.lock" (
        echo �t�@�C���u%%f�v��������...
        cscript //nologo "%~dp0ExcelToPDF.vbs" "%~dp0%%f" "%outputfolder%"
        echo.> "%%~nf.lock"
    )
)

rem ������Ƀ��b�N�t�@�C�����폜
del *.lock

echo.
echo ���ׂĂ̏������������܂����I
pause