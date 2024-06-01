' `vbscript
Dim shell
Set shell = CreateObject("WScript.Shell")

' 切换到工作目录
' shell.CurrentDirectory = "C:\your\git\repository\path"

' 执行Git命令
shell.Run "git init"
' 初始化一个空的Git仓库
shell.Run "git add ."
' 将所有文件添加到暂存区
shell.Run "git commit -m ""Initial commit"""
' 提交暂存区的文件

' 可以添加更多的Git命令，如git pull, git push等

' 显示命令执行结果
Wscript.Echo "Git commands executed successfully."

' 释放对象
Set shell = Nothing
' "`

' 在脚本中，首先创建了一个WScript.Shell对象，然后使用其`Run`方法来执行Git命令。在示例中，我们执行了Git的初始化（git init）、添加文件到暂存区（git add .）、提交暂存区的文件（git commit -m "Initial commit"）等操作。你可以根据自己的需求添加更多的Git命令。

' 注意，你需要将示例脚本中的`C:\your\git\repository\path`替换为你的Git仓库的实际路径。

