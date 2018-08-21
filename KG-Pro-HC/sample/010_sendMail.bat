rem パス設定
call C:\pleiades\workspace\KG-Pro-HC\Common\000_Common.bat

rem VBScript実行
call %WSCR_PATH% %EXEC_PATH%\010_sendMail.vbs %COMM_PATH% %MODU_PATH% %DATA_PATH% %EXEC_PATH%
