Attribute VB_Name = "FileCreate"
Sub FileCreate02()。
'日付を基に連続でファイル生成を行う
Dim Filepath As String
Dim fbo, aan As Object
Set fbo = CreateObject("Scripting.FileSystemObject")
Dim Pname As String 'Projectname

Pname = "Date681"

Filepath = DesktopFilepath(Filepath) & Pname
Debug.Print Filepath

Filepath = fbo.createfolder(Filepath)

Dim mon As Integer

mon = 8

For i = 1 To 31
    Set aan = fbo.createtextfile(Filepath & "\" & NFormat(0) & "-" & mon & "-" & i & ".txt")
Next i

End Sub