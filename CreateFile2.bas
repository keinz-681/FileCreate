Attribute VB_Name = "CreateFile2"
Sub FileCreate2()

Dim filepath As String
Dim fbo As Object
Set fbo = CreateObject("Scripting.FileSystemObject")
Dim pname As String

pname = "TryOne"

filepath = DesktopFilepath & pname

Debug.Print Len(Dir(filepath))

'If (Len(Dir(Filepath)) = 0) Then
'Debug.Print "Folder already existed!!"
'Exit Sub
'Else
filepath = fbo.createfolder(filepath)

'End If

Dim i As Integer

Dim fname As String
Dim Tpath As String

For i = 1 To 120 Step 1

    fname = Cells(i, 1)
    Debug.Print fname
    Tpath = filepath & "\" & fname & ".html"
    Open Tpath For Output As #1
        Print #1, "<!DOCTYPE html>"
        Print #1, "<html><head>"
        Print #1, "<title>Non-Title</title>"
        Print #1, "</head>"
        Print #1, "<body>"
        Print #1, "<div>"
        Print #1, "<h1>Parent-Name:Example</h1>"
        Print #1, "<p>Parent-Description:Welcome this page!<br>this page is Example for the project.</p>"
        Print #1, "</div>"
        Print #1, "<!--Children-Pages-Links-->"
        Print #1, "<div>"
        Print #1, "<a href="">Link</a>"
        Print #1, "</div>"
        Print #1, "</body></html>"

    Close #1
Next i

End Sub
Private Function DesktopFilepath() As String

    DesktopFilepath = "C:\Users\" & Environ("Username") & "\Desktop\"
End Function
