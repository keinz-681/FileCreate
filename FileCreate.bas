Attribute VB_Name = "����u����0"
Sub FileCreate()
' �e�L�X�g���s�œǂݍ���Ńt�H���_�����������ōs���B
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
