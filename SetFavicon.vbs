Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8
Const adTypeBinary=1, adSaveCreateOverWrite=2

'FileSystemオブジェクトを作成
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'引数を配列に格納
Dim arrArgs: Set arrArgs=WScript.Arguments

'引数がなければ終了
If arrArgs.Count=0 Then
	MsgBox "ファイルまたはフォルダをドロップしてください。", vbOkOnly+vbInformation, "SetFavicon"
	Wscript.Quit
End If

'faviconを保存するフォルダを設定
Dim faviconPath: faviconPath=objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\AppData\Roaming\favicon"
If Not objFS.FolderExists(faviconPath) Then objFS.CreateFolder(faviconPath)

'main()
	RecProcess arrArgs
	MsgBox "終了しました。", vbOkOnly+vbInformation, "SetFavicon"
	Set objShell=Nothing: Set objFS=Nothing
'end of main

'引数がフォルダの場合は再帰させる
Sub RecProcess(arrArgs)
	Dim path
	For Each path In arrArgs
		if objFS.FileExists(path) Then
			SetFavicon path
		ElseIf objFS.FolderExists(path) Then
			RecProcess objFS.GetFolder(path).Files
			RecProcess objFS.GetFolder(path).SubFolders
		Else
			MsgBox path & "は存在しません。" & vbCrLf & "次のファイルを処理します。", vbOKOnly+vbExclamation, "SetFavicon"
		End if
	Next
End Sub

'fabiconを取得してアイコンに反映させるプロシージャ
Sub SetFavicon(srcFilename)
'	出力テスト
	if MsgBox("srcFilename = " & srcFilename, vbOkCancel, "test")=vbCancel Then Wscript.Quit
	'拡張子を確認
	Dim dstFilename: dstFilename=""
	Select Case LCase(objFS.GetExtensionName(srcFilename))
		Case "url"
			srcFilename=dstFilename
		Case "website"
			Select Case MsgBox(srcFilename & "はInternet Explorerのピン止めショートカットです。" & vbCrLf & _
							   "インターネットショートカットに変換してよろしいですか？。", vbYesNo+vbQuestion, "SetFavicon")
				Case vbYes
					'新しい.urlファイル名を生成
					dstFilename=objFS.GetParentFolderName(srcFilename) & "\" & objFS.getBaseName(srcFilename) & ".url"
					If objFS.FileExists(dstFilename) Then
						Select Case MsgBox("インターネットショートカットが既に存在します。" & vbCrLf & "上書きしますか？", vbYesNo+vbQuestion, "SetFavicon")
							Case vbYes
								'何もしない
							Case vbNo
								MsgBox srcFilename & "は処理しませんでした。" & "次のファイルを処理します。", vbOkOnly+vbInformation, "SetFavicon"
								Exit Sub
							Case Else
								MsgBox "未定義のエラーです。" & vbCrLf & "次のファイルを処理します。", vbOkOnly+vbCritical, "SetFavicon"
								Exit Sub
						End Select
					End If
				Case vbNo
					MsgBox srcFilename & "は処理しませんでした。" & vbCrLf & "次のファイルを処理します。", vbOkOnly+vbInformation, "SetFavicon"
					Exit Sub
				Case Else
					MsgBox "未定義のエラーです。" & vbCrLf & "次のファイルを処理します。", vbOkOnly+vbCritical, "SetFavicon"
					Exit Sub
			End Select
		Case Else
			MsgBox srcFilename & "はインターネットショートカットではありません。" & vbCrLf & "次のファイルを処理します。", vbOkOnly+vbExclamation, "SetFavicon"
			Exit Sub
	End Select

	'.urlまたは.websiteファイルの内容を取得
	Dim strLine, strURL, strIconFile, strIconIndex, strHotKey, strIDList
	strURL="": strIconFile="": strIconIndex="0": strHotKey="0": strIDList="0"
	Dim objSrcFile: Set objSrcFile = objFS.OpenTextFile(srcFilename, ForReading)
	Do Until objSrcFile.AtEndOfStream
		strLine=objSrcFile.ReadLine
		Select Case Left(strLine, 7)
			Case "URL=htt"
				strURL=Mid(strLine, 5)
			Case "IconFil"
				strIconFile=Mid(strLine,10)
			Case "IconInd"
				strIconIndex=Mid(strLine,10)
			Case "HotKey="
				strHotKey=Mid(strLine, 8)
			Case "IDList="
				strIDList=Mid(strLine, 8)
		End Select
	Loop
	If strURL="" Then
		MsgBox srcFilename & "には、URLが設定されていません。" & vbCrLf & "次のファイルを処理します。", vbOkOnly+vbExclamation, "SetFavicon"
		Exit Sub
	End If
	objsrcFile.Close
'	出力テスト
	if MsgBox("strURL = " & strURL, vbOkCancel, "test")=vbCancel Then Wscript.Quit

	'faviconのURLを取得
	Dim faviconURL: Set faviconURL=""
	Dim objXmlHttp: Set objXmlHttp=CreateObject("MSXML2.XMLHTTP")
	objXmlHttp.Open "GET", strURL, False '同期処理
	objXmlHttp.Send
	Dim responseHeaders: responseHeaders=objXmlHttp.getAllResponseHeaders()
	DIm headerLine, headerLines: headerLines=Split(responseHeaders, vbCrLf)
	For Each headerLine In headerLines
		'HTMLファイルのContent-Typeを取得する
		If InStr(headerLine, "Content-Type:")>0 Then
			Dim contentType: contentType=Split(headerLine, ":")(1)
			'HTMLファイルの場合
			If InStr(contentType, "text/html")>0 Then
				Dim htmlText: htmlText=objXmlHttp.responseText
				Dim doc: Set doc=CreateObject("htmlfile")
				'HTMLのテキストを解析してDOMを作成
				doc.write htmlText
				Dim linkTag, linkTags: Set linkTags=doc.getElementsByTagName("link")
				For Each linkTag In linkTags
					'rel属性が"shortcut icon"または"icon"であるlinkタグを探す
					If (linkTag.getAttribute("rel")="shortcut icon" Or linkTag.getAttribute("rel")="icon") Then
						faviconURL=linkTag.getAttribute("href")
						Exit For
					End If
				Next
				'Content-Typeがtext/htmlであるため
				Exit For
			End If
		End If
	Next
	'faviconのURLが取得できなかった場合
	If faviconURL="" And strIconFile<>"" Then
		faviconURL=strIconFile
	Else 'IconFileが特定できなかった場合
		MsgBox srcFilename & "にはアイコンが設定されていません。" & vbCrLf & "次のファイルを処理します。", vbOkOnly+vbExclamation, "SetFavicon"
		Exit Sub
	End If
'	出力テスト
	if MsgBox("faviconURL = " & faviconURL, vbOkCancel, "test")=vbCancel Then Wscript.Quit

	'faviconを保存するファイル名を作成
	Dim faviconFilename
	If Right(faviconPath, 1)<>"\" Then faviconPath=faviconPath & "\"
	faviconURL=Mid(Replace(faviconURL, "favicon.", ""), InStr(faviconURL, "//")+2)
	faviconFilename=faviconPath & Replace(faviconURL, "/", ".")
'	出力テスト
	If MsgBox("faviconFilename = " & faviconFilename, vbOkCancel, "test")=vbCancel Then Wscript.Quit

	'faviconを取得して保存
	objXmlHttp.Open "GET", faviconURL, False '同期処理
	objXmlHttp.Send
	If objXmlHttp.Status=200 Then
		Dim objStream: Set objStream=CreateObject("ADODB.Stream")
		objStream.Type=adTypeBinary
		objStream.Open
		objStream.Write objXmlHttp.ResponseBody
		objStream.SaveToFile faviconFilename, adSaveCreateOverWrite
		objStream.Close
		Set objStream=Nothing
	Else
		MsgBox "アイコンのダウンロードに失敗しました。" & vbCrLf & faviconURL & vbCrLf,  "次のファイルを処理します。", vbOkOnly+vbExclamation, "SetFavicon"
		exit Sub
	End If
	Set objXmlHttp=Nothing
'	出力テスト
	If MsgBox("faviconFilename = " & faviconFilename, vbOkCancel, "test")=vbCancel Then Wscript.Quit

	'.urlファイルを再構成
	Dim objURLFile : Set objURLFile=objFS.CreateTextFile(dstFilename, ForWriting, True) 'ファイルが存在しない場合は作成する
	objURLFile.WriteLine("[InternetShortcut]")
	objURLFile.WriteLine("URL=" & strURL)
	objURLFile.WriteLine("IconFile=" & faviconFilename)
	objURLFile.WriteLine("IconIndex=" & StrIconIndex)
	objURLFile.WriteLine("HotKey=" & strHotKey)
	objURLFile.WriteLine("IDList=" & strIDList)
	objURLFile.Close
'	出力テスト
	If MsgBox("URL = " & strURL, vbOkCancel, "test")=vbCancel Then Wscript.Quit

	'変更後の.urlファイルのプロパティを書き換える '.urlファイルにはIconLocationプロパティがないため.lnkにリネーム
	Dim tmpFilename: tmpFilename=Left(dstFilename, InStrRev(dstFilename, ".")-1) & ".lnk"
	objShell.MoveFile dstFilename, tmpFilename
	' リネームした.lnkファイルのアイコンを変更
	Dim objLinkFile: Set obiLinkFile=objShell.CreateShortcut(tmpFilename)
	objLinkFile.IconLocation=faviconFilename
	objLinkFile.Save
	'リネームした.urlファイルを元に戻す
	objShell.MoveFile tmpfilename, dstFilename
'	出力テスト
	MsgBox "End of SubRoutine", vbOkOnly, "test"
End Sub
