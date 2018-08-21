function func_test(getMsg)
    wscript.echo(getMsg)
end function

'=========================================
'メール送信
'=========================================
function mailSend(fromAddress, toAddress, ccAddress, mailTitle, mailText, sendFileName1, sendFileName2, sendFileName3, sendFileName4, sendFileName5, sendFileName6)

	Dim MAIL_SERVER
	MAIL_SERVER = "mail.os.ksi.co.jp" 'KSIメールサーバ

	Dim objFSO      ' FileSystemObject
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	'メール生成
	Dim mailObj 
	Set mailObj      = CreateObject("CDO.Message")
	mailObj.From     = fromAddress
	mailObj.To       = toAddress
	mailObj.Cc       = ccAddress
	mailObj.Subject  = mailTitle
	mailObj.TextBody = mailText
	mailObj.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	mailObj.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MAIL_SERVER
	mailObj.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	'ファイル１がある場合は添付
	If (sendFileName1 <>"") AND (objFSO.FileExists(sendFileName1)) Then
		mailObj.AddAttachment sendFileName1
		objFSO.DeleteFile sendFileName1
	end If
	
	'ファイル2がある場合は添付
	If (sendFileName2 <>"") AND (objFSO.FileExists(sendFileName2)) Then
		mailObj.AddAttachment sendFileName2
		objFSO.DeleteFile sendFileName2
	end If
	
	'ファイル3がある場合は添付
	If (sendFileName3 <>"") AND (objFSO.FileExists(sendFileName3)) Then
		mailObj.AddAttachment sendFileName3
		objFSO.DeleteFile sendFileName3
	end If
	
	'ファイル4がある場合は添付
	If (sendFileName4 <>"") AND (objFSO.FileExists(sendFileName4)) Then
		mailObj.AddAttachment sendFileName4
		objFSO.DeleteFile sendFileName4
	end If
	
	'ファイル5がある場合は添付
	If (sendFileName5 <>"") AND (objFSO.FileExists(sendFileName5)) Then
		mailObj.AddAttachment sendFileName5
		objFSO.DeleteFile sendFileName5
	end If
	
	'ファイル6がある場合は添付
	If (sendFileName6 <>"") AND (objFSO.FileExists(sendFileName6)) Then
		mailObj.AddAttachment sendFileName6
		objFSO.DeleteFile sendFileName6
	end If

	mailObj.Configuration.Fields.Update

	'メール送信
	mailObj.Send

	'オブジェクト初期化
	Set mailObj = Nothing
	Set fileObj = Nothing

end function
