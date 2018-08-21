function func_test(getMsg)
    wscript.echo(getMsg)
end function

'=========================================
'���[�����M
'=========================================
function mailSend(fromAddress, toAddress, ccAddress, mailTitle, mailText, sendFileName1, sendFileName2, sendFileName3, sendFileName4, sendFileName5, sendFileName6)

	Dim MAIL_SERVER
	MAIL_SERVER = "mail.os.ksi.co.jp" 'KSI���[���T�[�o

	Dim objFSO      ' FileSystemObject
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	'���[������
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
	
	'�t�@�C���P������ꍇ�͓Y�t
	If (sendFileName1 <>"") AND (objFSO.FileExists(sendFileName1)) Then
		mailObj.AddAttachment sendFileName1
		objFSO.DeleteFile sendFileName1
	end If
	
	'�t�@�C��2������ꍇ�͓Y�t
	If (sendFileName2 <>"") AND (objFSO.FileExists(sendFileName2)) Then
		mailObj.AddAttachment sendFileName2
		objFSO.DeleteFile sendFileName2
	end If
	
	'�t�@�C��3������ꍇ�͓Y�t
	If (sendFileName3 <>"") AND (objFSO.FileExists(sendFileName3)) Then
		mailObj.AddAttachment sendFileName3
		objFSO.DeleteFile sendFileName3
	end If
	
	'�t�@�C��4������ꍇ�͓Y�t
	If (sendFileName4 <>"") AND (objFSO.FileExists(sendFileName4)) Then
		mailObj.AddAttachment sendFileName4
		objFSO.DeleteFile sendFileName4
	end If
	
	'�t�@�C��5������ꍇ�͓Y�t
	If (sendFileName5 <>"") AND (objFSO.FileExists(sendFileName5)) Then
		mailObj.AddAttachment sendFileName5
		objFSO.DeleteFile sendFileName5
	end If
	
	'�t�@�C��6������ꍇ�͓Y�t
	If (sendFileName6 <>"") AND (objFSO.FileExists(sendFileName6)) Then
		mailObj.AddAttachment sendFileName6
		objFSO.DeleteFile sendFileName6
	end If

	mailObj.Configuration.Fields.Update

	'���[�����M
	mailObj.Send

	'�I�u�W�F�N�g������
	Set mailObj = Nothing
	Set fileObj = Nothing

end function
