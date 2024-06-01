Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")


' 基本思路：在bvs当前目录Set objShell 打开命令窗口，输入git add .命令, 输入本次更新内容，（默认不显示运行窗口“，0，”，且等待命令执行成功“True”.

' 这个地方没有选择目录位置，默认为vbs存放的目录

' ' 获取当前脚本所在目录
' currentDirectory = objShell.CurrentDirectory

' ' 进入当前脚本所在目录
' objShell.CurrentDirectory = currentDirectory



' 使用VPN，（git报错）？那么请第一次使用时将以下代码取消注释
' ***********************VPN***************************
' git config --global http.proxy 127.0.0.1:10809

' git config --global https.proxy 127.0.0.1:10809
' ****************************************************
' 执行git add命令
objShell.Run "git add .", 0, True

' 执行git commit命令
' 输入 commit message
commitMessage = InputBox("please input commit message:", "Commit Message")
objShell.Run "git commit -m """ & commitMessage & """", 0, True


' pull--allow-unrelated-histories可加可不加，报错就加试一下
' Git 通常不允许合并不相关历史的分支。可以通过添加 `--allow-unrelated-histories` 选项来允许合并不相关的历史。
objShell.Run "git pull origin master --allow-unrelated-histories", 0, True


' 执行git push命令
objShell.Run "git push", 0, True

' 释放对象
Set objShell = Nothing

' 这样，脚本将使用当前脚本所在的目录作为本地Git仓库的目录。保存脚本文件并运行它即可