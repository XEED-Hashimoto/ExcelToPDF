' ExcelToPDF.vbs
' �G�N�Z���t�@�C���̊e�V�[�g��PDF�ɕϊ�����X�N���v�g

' �R�}���h���C���������擾
If WScript.Arguments.Count < 1 Then
    WScript.Echo "�������s�����Ă��܂��B�G�N�Z���t�@�C���̃p�X���w�肵�Ă��������B"
    WScript.Quit
End If

excelFilePath = WScript.Arguments(0)
If WScript.Arguments.Count > 1 Then
    outputFolder = WScript.Arguments(1)
Else
    outputFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(excelFilePath)
End If

' �t�@�C�����i�g���q�Ȃ��j���擾
Set fso = CreateObject("Scripting.FileSystemObject")
fileName = fso.GetBaseName(excelFilePath)

' Excel�A�v���P�[�V�������N��
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Excel���N���ł��܂���ł����BExcel���C���X�g�[������Ă��邩�m�F���Ă��������B"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

' �G�N�Z���t�@�C�����J��
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(excelFilePath)
If Err.Number <> 0 Then
    WScript.Echo "�t�@�C���u" & excelFilePath & "�v���J���܂���ł����B"
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit
End If
On Error GoTo 0

' �e�V�[�g��PDF�Ƃ��ĕۑ�
For Each objSheet In objWorkbook.Sheets
    sheetName = objSheet.Name
    pdfPath = outputFolder & "\" & fileName & "_" & sheetName & ".pdf"
    
    WScript.Echo "  �V�[�g�u" & sheetName & "�v��PDF�����Ă��܂�..."
    
    ' ���݂̃V�[�g���A�N�e�B�u�ɂ���
    objSheet.Activate
    
    ' PDF�Ƃ��ĕۑ�
    On Error Resume Next
    objWorkbook.ActiveSheet.ExportAsFixedFormat 0, pdfPath, 0, 1, 0, , , 0
    If Err.Number <> 0 Then
        WScript.Echo "  �G���[: �V�[�g�u" & sheetName & "�v��PDF���Ɏ��s���܂����B"
    End If
    On Error GoTo 0
Next

' �G�N�Z�������
objWorkbook.Close False
objExcel.Quit

' �I�u�W�F�N�g�̉��
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "�t�@�C���u" & excelFilePath & "�v�̂��ׂẴV�[�g��PDF�����܂����B"
