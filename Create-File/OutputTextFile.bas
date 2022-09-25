Attribute VB_Name = "Module2"
Sub outTxtFile1()
  Dim strFilePath As String
  strFilePath = ActiveWorkbook.Path & "\test1.txt" 'ファイルパス
 'Debug.Print strFilePath
  Open strFilePath For Output As #1
  Dim i As Integer
  For i = 1 To 5
    Print #1, CStr(i)
  Next i
  Close #1
  MsgBox "ファイルの作成が完了しました", vbInformation
End Sub