Option Explicit

Dim baseFileFolderName				'画像ファイル格納場所名
Dim baseFileFolderObj				'画像ファイル格納場所
Dim changeFileNameAfterFolder 		'変換後画像ファイル格納先
Dim objFS							'ファイルシステムオブジェクト
Dim baseName						'置換前の名前
Dim changeName						'置換後の名前
Dim file							'ファイルリストから取り出した一つのファイル名
Dim newFileName						'新ファイル名
Dim counter							'置換ファイル数

baseFileFolderName = InputBox("画像ファイル格納場所を教えて")

'===============================
' ファイルシステムオブジェクト作成
'===============================
Set objFS = CreateObject("Scripting.FileSystemObject")

'===============================
' フォルダ存在確認
'===============================
If objFS.FolderExists(baseFileFolderName) Then
	'===============================
	'置換対象文字列入力
	'===============================
	baseName = InputBox("置換対象の文字列を入れてください")
	If Trim(baseName) = "" Then
		'入力値なし、または0の場合
		Set objFS = Nothing
		WScript.Echo "なんかいれて"
		WScript.Quit
	End If

	'===============================
	'置換後文字列入力
	'===============================
	changeName = InputBox("置換後の文字列を入れてください")
	'If Trim(baseName) = ""  Then
		'入力値なし、または0の場合
		'Set objFS = Nothing
		'WScript.Echo "なんかいれて"
		'WScript.Quit
	'End If
	'===============================
	' 存在する場合はフォルダオブジェクト取得
	'===============================
	Set baseFileFolderObj = objFS.GetFolder(baseFileFolderName)
	changeFileNameAfterFolder = baseFileFolderName & "\henkango"

	If objFS.FolderExists(changeFileNameAfterFolder) Then
	Else
		objFS.createFolder(changeFileNameAfterFolder)
	End If

	counter = 0

	'===============================
	' 置換
	'===============================
	' フォルダ内のファイルをループ
	For Each file In baseFileFolderObj.Files
		' ファイル名にoldNameが含まれているかチェック
		If InStr(file.Name, baseName) > 0 Then
			' 新しいファイル名を作成
			newFileName = Replace(file.Name, baseName, changeName)
			' ファイルをコピー
			objFS.CopyFile file.Path, objFS.BuildPath(changeFileNameAfterFolder, newFileName)
			counter = counter + 1
		End If
	Next
    
    set baseFileFolderObj = Nothing

	if counter = 0 Then
		WScript.Echo "置換対象ファイルがありません"
	else
		WScript.Echo "置換対象ファイル：" & counter & "件"
	End If

Else
	'フォルダが存在しない場合の処理
	WScript.Echo "フォルダがねーぞ"
End If

set objFS = Nothing
