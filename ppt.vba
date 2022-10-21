
Sub 合并ppt()

'合并当前目录下的所有ppt到一个ppt中

 Dim MyName, dic, i, MyFileName, FilePath
 Set dic = CreateObject("Scripting.Dictionary")
 On Error Resume Next
   FilePath = Application.ActivePresentation.Path
    dic.Add (FilePath & "\"), ""
    i = 0
    Do While i < dic.Counts
        ke = dic.keys
        MyName = Dir(ke(i), vbDirectory)
        Do While MyName <> ""
            If MyName <> "." And MyName <> ".." Then
                If (GetAttr(ke(i) & MyName) And vbDirectory) = vbDirectory Then
                    dic.Add (ke(i) & MyName & "\"), ""
                End If
            End If
            MyName = Dir
        Loop
        i = i + 1
    Loop
    For Each ke In dic.keys
        MyFileName = Dir(ke & "*.PPTX")
        Do While MyFileName <> ""
                Set pptInput = Presentations.Open(FilePath & "/" & MyFileName)
                Set pptoutput = Presentations.Open(FilePath & "/" & "111111.pptm") '合并后的文件名称
                For j = 1 To pptInput.Slides.Count
                    pptInput.Slides(j).Copy
                   pptoutput.Slides.Paste (pptoutput.Slides.Count)
                   Application.ActivePresentation.Save
         Next
                   
                   Application.ActivePresentation.Close
          Application.ActivePresentation.Save
          Presentations(FilePath & "\" & MyFileName).Close
            MyFileName = Dir
            Loop
    Next
    
   End Sub
