Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Dim fld ' as object
dim Path ' As string
Path = fso.GetParentFolderName(WScript.ScriptFullName) '获取脚本所在文件夹字符串
Set fld=fso.GetFolder(Path) '通过路径字符串获取文件夹对象

Dim Sum,IsChooseDelete,ThisTime
Dim LogFile
Set LogFile= fso.opentextFile("log.txt",8,true)

Dim List
Set List= fso.opentextFile("ConvertFileList.txt",2,true)

Call LogOut("开始遍历文件") 
Call TreatSubFolder(fld) '调用该过程进行递归遍历该文件夹对象下的所有文件对象及子文件夹对象

Sub LogOut(msg)
    ThisTime=Now
    LogFile.WriteLine(year(ThisTime) & "-" & Month(ThisTime) & "-" & day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld) 
    Dim File
    Dim ts
    For Each File In fld.Files '遍历该文件夹对象下的所有文件对象
        If fso.GetExtensionName(File) ="doc" or fso.GetExtensionName(File)="docx" Then
            List.WriteLine(File.Path)
            Sum = Sum + 1
        End If
    Next
    Dim subfld
    For Each subfld In fld.SubFolders '递归遍历子文件夹对象
        TreatSubFolder subfld
    Next
End Sub
List.close

If MsgBox("文件遍历已完成，已找到" & Sum & "个word文档，详细列表在" & vbCrlf & "ConvertFileList.txt" & vbCrlf & "是否将这些文档转换为PDF？", vbYesNo + vbInformation, "文档遍历完成") = vbYes Then
    If MsgBox("是否在转换完毕后删除DOC文档?", vbYesNo+vbInformation, "是否在转换完毕后删除源文档?") = vbYes Then
        IsChooseDelete = MsgBox("请再次确认，是否在转换完毕后删除DOC文档?", vbYesNo + vbExclamation, "是否在转换完毕后删除源文档?")
    End If
else
    Msgbox("已取消转换操作")
    Wscript.Quit
End If
MsgBox "请在开始转换前退出所有Word文档避免文档占用错误发生", vbOKOnly + vbExclamation, "警告"


Const wdFormatPDF = 17
Set wdapp = CreateObject("Word.Application")'创建Word对象
wdapp.Visible=false '设置视图不可见

Dim Finished
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    Set objDoc = wdapp.Documents.Open(FilePath)
    objDoc.SaveAs Left(FilePath,InstrRev(FilePath,".")) & "pdf", wdFormatPDF '另存为PDF文档
    LogOut("文档" & FilePath & "已转换完成。(" & Finished & "/" & Sum & ")")
    wdapp.ActiveDocument.Close  
    Finished = Finished + 1
    If IsChooseDelete = vbYes Then
        fso.deleteFile FilePath
        LogOut("文件" & FilePath & "已被成功删除")
    End If
loop
'扫尾处理，ConvertFileList.txt和log.txt要自动删除的请去掉下面两行开头单引号
'fso.deleteFile "ConvertFileList.txt"
'fso.deleteFile "log.txt"
List.close
LogOut("文档转换已完成")
LogFile.close 
Set fso = nothing

Dim Msg
Msg = "已成功转换" & Finished & "个文件"
If IsChooseDelete = vbYes Then
    Msg=Msg + "并成功删除源文件"
MsgBox Msg
wdapp.Quit
Wscript.Quit
