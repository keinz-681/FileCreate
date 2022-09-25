Attribute VB_Name = "FileCreate"
Sub outTxtFile2()
  Dim strFilePath As String
  strFilePath = DesktopFilepath(Filepath) & "\test1.html" 'ファイルパス
 'Debug.Print strFilePath
  Open strFilePath For Output As #1
  Dim i, j As Integer
    Print #1, "<html>" & vbCrLf & "<head>" & vbCrLf & "</head>" & vbCrLf
    Print #1, "<body>" & vbCrLf & "<h1>こんにちは</h1>" & vbCrLf & "</body>" & vbCrLf & "</html>"
    For i = 1 To 50
    j = Cells(i, 1)
    Print #1, j & "<br>"
    Next i
  Close #1
  MsgBox "ファイルの作成が完了しました", vbInformation
End Sub
