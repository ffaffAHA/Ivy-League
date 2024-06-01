Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

objShell.Run "git add .", 0, True
commitMessage = InputBox("please input commit message:", "Commit Message")

objShell.Run "git commit -m """ & commitMessage & """", 0, True

' objShell.Run "git pull origin master --allow-unrelated-histories", 0, True

' 执行git push命令
objShell.Run "git push", 0, True
Set objShell = Nothing
' 这样，脚本将使用当前脚本所在的目录作为本地Git仓库的目录。保存脚本文件并运行它即可。