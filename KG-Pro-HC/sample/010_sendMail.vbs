Dim COMM_PATH
Dim MODU_PATH
Dim DATA_PATH
Dim EXEC_PATH

If WScript.Arguments.Count < 4 then
    '�����s���̓G���[
    WScript.echo("too few arguments.")
    WScript.Quit(-1)
    
Else
    COMM_PATH = WScript.Arguments(0)
    MODU_PATH = WScript.Arguments(1)
    DATA_PATH = WScript.Arguments(2)
    EXEC_PATH = WScript.Arguments(3)
End If

'���ʏ����̒�`
Dim FUNC_MAIL_SEND
FUNC_MAIL_SEND = COMM_PATH & "\010_mailSendFunction.vbs"

'�f�o�b�O
'Wscript.Echo FUNC_MAIL_SEND

'���t�����񐶐�
Dim strDate
strDate = Left(Replace(Replace(Replace(Replace(Now(), "/", "-"), ":", ""), " ", "_"), "-", ""), 8)

'===========================
'  ���[�����M����VBS��Ǎ�
'===========================
Dim objFSO_FUNC,objWSH_FUNC
Set objFSO_FUNC = CreateObject("Scripting.FileSystemObject")
Set objWSH_FUNC = objFSO_FUNC.OpenTextFile(FUNC_MAIL_SEND)
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close

'===========================
'  ���[�������p�̕ϐ�
'===========================
Dim fileExist
Dim titleTxt
Dim mailText
Dim fromAddress
Dim toAddress
Dim ccAddress
Dim attachPath1
Dim attachPath2
Dim attachPath3
Dim attachPath4

'===========================
'  ���[�����e����
'===========================
mailText    = "" _
& "�~�~����" & vbCr _
& "" & vbCr _
& "�����l�ł��A�����ł��B" & vbCr _
& "" & vbCr _
& "�e�X�g���M�ł�" & vbCr _
& "" & vbCr _
& "�ȏ�ł��A��낵�����肢�v���܂��B" & vbCr _
& "" & vbCr

'���M���A����ACC�A�^�C�g��
fromAddress = "takashi.aibara@kubota.com"
ccAddress   = "takashi.aibara@kubota.com"
toAddress   = "takashi.aibara@kubota.com"
titleTxt    = "�yTEST�z�e�X�g���[��(" & strDate & ")"

'===========================
'  �Y�t�t�@�C������
'===========================
attachPath1 = ""
attachPath2 = ""
attachPath3 = ""
attachPath4 = ""

'Wscript.Echo attachPath1

'�Y�t�t�@�C���̐���A�S���Ȃ��ꍇ�̓��[���Ȃ��A���镪�������M
Dim objFSO      ' FileSystemObject
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

fileExist = 0
If objFSO.FileExists(attachPath1) then fileExist = 1 else   attachPath1 = "" end if
If objFSO.FileExists(attachPath2) then fileExist = 1 else   attachPath2 = "" end if
If objFSO.FileExists(attachPath3) then fileExist = 1 else   attachPath3 = "" end if
If objFSO.FileExists(attachPath4) then fileExist = 1 else   attachPath4 = "" end if

'Wscript.Echo fileExist

'===========================
'  ���[�����t
'===========================
'If fileExist = 1 Then
   '���[�����t(from,to,cc,subject,text,attachment 
   Call mailSend(fromAddress, _
                 toAddress, _
                 ccAddress, _
                 titleTxt, _
                 mailText, _
                 attachPath1, attachPath2, attachPath3, attachPath4, "", "" )
'end If

'===========================
'  ��n��
'===========================
Set objFSO = Nothing
Set objWSH_FUNC = Nothing
Set objFSO_FUNC = Nothing