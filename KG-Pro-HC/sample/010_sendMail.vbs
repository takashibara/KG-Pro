Dim COMM_PATH
Dim MODU_PATH
Dim DATA_PATH
Dim EXEC_PATH

If WScript.Arguments.Count < 4 then
    '引数不足はエラー
    WScript.echo("too few arguments.")
    WScript.Quit(-1)
    
Else
    COMM_PATH = WScript.Arguments(0)
    MODU_PATH = WScript.Arguments(1)
    DATA_PATH = WScript.Arguments(2)
    EXEC_PATH = WScript.Arguments(3)
End If

'共通処理の定義
Dim FUNC_MAIL_SEND
FUNC_MAIL_SEND = COMM_PATH & "\010_mailSendFunction.vbs"

'デバッグ
'Wscript.Echo FUNC_MAIL_SEND

'日付文字列生成
Dim strDate
strDate = Left(Replace(Replace(Replace(Replace(Now(), "/", "-"), ":", ""), " ", "_"), "-", ""), 8)

'===========================
'  メール送信共通VBSを読込
'===========================
Dim objFSO_FUNC,objWSH_FUNC
Set objFSO_FUNC = CreateObject("Scripting.FileSystemObject")
Set objWSH_FUNC = objFSO_FUNC.OpenTextFile(FUNC_MAIL_SEND)
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close

'===========================
'  メール生成用の変数
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
'  メール内容生成
'===========================
mailText    = "" _
& "××さん" & vbCr _
& "" & vbCr _
& "お疲れ様です、相原です。" & vbCr _
& "" & vbCr _
& "テスト送信です" & vbCr _
& "" & vbCr _
& "以上です、よろしくお願い致します。" & vbCr _
& "" & vbCr

'送信元、宛先、CC、タイトル
fromAddress = "takashi.aibara@kubota.com"
ccAddress   = "takashi.aibara@kubota.com"
toAddress   = "takashi.aibara@kubota.com"
titleTxt    = "【TEST】テストメール(" & strDate & ")"

'===========================
'  添付ファイル制御
'===========================
attachPath1 = ""
attachPath2 = ""
attachPath3 = ""
attachPath4 = ""

'Wscript.Echo attachPath1

'添付ファイルの制御、全部ない場合はメールなし、ある分だけ送信
Dim objFSO      ' FileSystemObject
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

fileExist = 0
If objFSO.FileExists(attachPath1) then fileExist = 1 else   attachPath1 = "" end if
If objFSO.FileExists(attachPath2) then fileExist = 1 else   attachPath2 = "" end if
If objFSO.FileExists(attachPath3) then fileExist = 1 else   attachPath3 = "" end if
If objFSO.FileExists(attachPath4) then fileExist = 1 else   attachPath4 = "" end if

'Wscript.Echo fileExist

'===========================
'  メール送付
'===========================
'If fileExist = 1 Then
   'メール送付(from,to,cc,subject,text,attachment 
   Call mailSend(fromAddress, _
                 toAddress, _
                 ccAddress, _
                 titleTxt, _
                 mailText, _
                 attachPath1, attachPath2, attachPath3, attachPath4, "", "" )
'end If

'===========================
'  後始末
'===========================
Set objFSO = Nothing
Set objWSH_FUNC = Nothing
Set objFSO_FUNC = Nothing