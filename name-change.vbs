Option Explicit

Dim baseFileFolderName				'�摜�t�@�C���i�[�ꏊ��
Dim baseFileFolderObj				'�摜�t�@�C���i�[�ꏊ
Dim changeFileNameAfterFolder 		'�ϊ���摜�t�@�C���i�[��
Dim objFS							'�t�@�C���V�X�e���I�u�W�F�N�g
Dim baseName						'�u���O�̖��O
Dim changeName						'�u����̖��O
Dim file							'�t�@�C�����X�g������o������̃t�@�C����
Dim newFileName						'�V�t�@�C����
Dim counter							'�u���t�@�C����

baseFileFolderName = InputBox("�摜�t�@�C���i�[�ꏊ��������")

'===============================
' �t�@�C���V�X�e���I�u�W�F�N�g�쐬
'===============================
Set objFS = CreateObject("Scripting.FileSystemObject")

'===============================
' �t�H���_���݊m�F
'===============================
If objFS.FolderExists(baseFileFolderName) Then
	'===============================
	'�u���Ώە��������
	'===============================
	baseName = InputBox("�u���Ώۂ̕���������Ă�������")
	If Trim(baseName) = "" Then
		'���͒l�Ȃ��A�܂���0�̏ꍇ
		Set objFS = Nothing
		WScript.Echo "�Ȃ񂩂����"
		WScript.Quit
	End If

	'===============================
	'�u���㕶�������
	'===============================
	changeName = InputBox("�u����̕���������Ă�������")
	'If Trim(baseName) = ""  Then
		'���͒l�Ȃ��A�܂���0�̏ꍇ
		'Set objFS = Nothing
		'WScript.Echo "�Ȃ񂩂����"
		'WScript.Quit
	'End If
	'===============================
	' ���݂���ꍇ�̓t�H���_�I�u�W�F�N�g�擾
	'===============================
	Set baseFileFolderObj = objFS.GetFolder(baseFileFolderName)
	changeFileNameAfterFolder = baseFileFolderName & "\henkango"

	If objFS.FolderExists(changeFileNameAfterFolder) Then
	Else
		objFS.createFolder(changeFileNameAfterFolder)
	End If

	counter = 0

	'===============================
	' �u��
	'===============================
	' �t�H���_���̃t�@�C�������[�v
	For Each file In baseFileFolderObj.Files
		' �t�@�C������oldName���܂܂�Ă��邩�`�F�b�N
		If InStr(file.Name, baseName) > 0 Then
			' �V�����t�@�C�������쐬
			newFileName = Replace(file.Name, baseName, changeName)
			' �t�@�C�����R�s�[
			objFS.CopyFile file.Path, objFS.BuildPath(changeFileNameAfterFolder, newFileName)
			counter = counter + 1
		End If
	Next
    
    set baseFileFolderObj = Nothing

	if counter = 0 Then
		WScript.Echo "�u���Ώۃt�@�C��������܂���"
	else
		WScript.Echo "�u���Ώۃt�@�C���F" & counter & "��"
	End If

Else
	'�t�H���_�����݂��Ȃ��ꍇ�̏���
	WScript.Echo "�t�H���_���ˁ[��"
End If

set objFS = Nothing
