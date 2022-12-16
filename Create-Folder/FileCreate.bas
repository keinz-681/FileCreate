Attribute VB_Name = "試作置き場0"
Sub FileCreate()
    ' テキストを行で読み込んでフォルダ生成を自動で行う。
    Dim Filepath As String
    Dim fbo, aan As Object
    Set fbo = CreateObject("Scripting.FileSystemObject")
    Dim Pname As String 'Projectname

    Pname = "おためし"

    Filepath = DesktopFilepath(Filepath) & Pname
    Debug.Print Filepath

    Filepath = fbo.createfolder(Filepath)

    Dim starray() As String
    Dim str As String

    str = "ant,bear,logn"

    starray = Split(str, ",")

    For i = 0 To UBound(starray)
        Set aan = fbo.createtextfile(Filepath & "\" & starray(i) & ".md")
    Next i

End Sub

Function DesktopFilepath(Filepath) As String

    DesktopFilepath = "C:\Users\" + Environ("Username") + "\Desktop\"

End Function
