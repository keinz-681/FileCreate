Attribute VB_Name = "CreateFile2"
Sub FileCreate2()

Dim Filepath As String
Dim fbo As Object
Set fbo = CreateObject("Scripting.FileSystemObject")
Dim pname As String

pname = "TryOne"

Filepath = DesktopFilepath & pname

Filepath = fbo.createfolder(Filepath)

Dim i As Integer

Dim fname As String
Dim Tpath As String

For i = 1 To 120 Step 1

    fname = Cells(i, 1)
Debug.Print fname
Tpath = Filepath & "\" & fname & ".txt"
Open Tpath For Output As #1
Close #1
Next i

End Sub
Private Function DesktopFilepath() As String

    DesktopFilepath = "C:\Users\" & Environ("Username") & "\Desktop\"
End Function
