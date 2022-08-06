Attribute VB_Name = "試作置き場0"
Sub FileCreate()
' テキストを行で読み込んでフォルダ生成を自動で行う。
Dim filepath As String
Dim fbo, aan As Object
Set fbo = CreateObject("Scripting.FileSystemObject")
Dim projectname As String

projectname = "to"

filepath = "C:\Users\" + Environ("Username") + "\Desktop\" & projectname
Debug.Print filepath


filepath = fbo.createfolder(filepath)

Dim starray() As String
Dim str As String

str = "ant,bear,logn"

starray = Split(str, ",")

Dim char As String

char = ".md"

For i = 0 To UBound(starray)
    Set aan = fbo.createtextfile(filepath & "\" & starray(i) & char)

Next i

End Sub
