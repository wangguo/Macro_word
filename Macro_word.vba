Sub 图片居中()

' 在所有[img][/img]标记前后加上[align=center][/align]

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[img]"
        .Replacement.Text = "[align=center][img]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[/img]"
        .Replacement.Text = "[/img][/align]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 删除空白行()

'删除空行

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub 段首加空格()

'在每段段首加上4个半角空格

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p    "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub 段首删空格()

'删除每段段首的空格

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub 删图()

'删除Word文档中的所有图片

Dim pic As InlineShape
 
 For Each pic In ActiveDocument.InlineShapes
 
 If pic.Width <> 0 Then

pic.Select
 
 Selection.Delete
 
 End If


Next


End Sub


Sub 手动换行()

'将所有段落标记替换为手动换行标记


    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^l"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub 自动换行()

'将所有手动换行标记替换为段落标记

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 换HTML空格()

' 将所有HTML格式空格替换为半角空格

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
       
End Sub
Sub 自动缩放图()

'将Word文档中的可见图片调整为统一大小


Dim myis As InlineShape

For Each myis In ActiveDocument.InlineShapes
    
If myis.Width > CentimetersToPoints(2.5) Then
  
      
If myis.Width < CentimetersToPoints(0.5) Then GoTo 10
If myis.Height < CentimetersToPoints(0.5) Then GoTo 10
     
 myis.Reset
     
    ' myis.PictureFormat.ColorType = msoPictureGrayscale

   myis.LockAspectRatio = msoTrue
     
    
   myis.ScaleWidth = 99
    
  If myis.Width > CentimetersToPoints(1) Then myis.Width = CentimetersToPoints(3.5)
    
    myis.ScaleHeight = myis.ScaleWidth
         
      
  End If

10: Next myis
End Sub

Sub 图居中()

'居中Word文档中的所有可见图片

Dim myis As InlineShape

For Each myis In ActiveDocument.InlineShapes
    
  If myis.Width > 0 Then
  
  myis.Select
  
  
  Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      
        
  End If

Next myis
End Sub


Sub 换全角空格()

' 将所有全角空格替换为半角空格

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "　"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub 换空格()
 
  换HTML空格
  换全角空格

End Sub


Sub 添加行号()


'在选中的每个段落前加上1. 2. 3.……


Dim parag As Paragraph
Dim nLineNum: nLineNum = 0
Dim selRge As Range
Set selRge = Selection.Range
  
  For Each parag In Selection.Paragraphs
  nLineNum = nLineNum + 1
  
  
If nLineNum > 0 Then
   selRge.Paragraphs(nLineNum).Range.InsertBefore (nLineNum & ".  ")
 
 End If
  
  
'个位数前自动添加0
' If nLineNum < 10 And nLineNum > 0 Then
'    selRge.Paragraphs(nLineNum).Range.InsertBefore ("0" & nLineNum & "   ")
'  Else
'    selRge.Paragraphs(nLineNum).Range.InsertBefore (nLineNum & "   ")
'  End If
  
  
  
 Next

End Sub

Sub 加粗()

'在选中的文字前后加上[b][/b]
  
With Selection
    .InsertBefore "[b]"
End With

With Selection
    .InsertAfter "[/b]"
End With


End Sub


Sub 加链接()

  
  
With Selection
    .InsertBefore "[url]"
End With

With Selection
    .InsertAfter "[/url]"
End With


End Sub

Sub 加链接2()
  
  
With Selection
    .InsertBefore "[url=]"
End With

With Selection
    .InsertAfter "[/url]"
End With


End Sub

Sub 引用格式()

'在选中的文字前后加上[quote][/quote]
  
With Selection
    .InsertBefore "[quote]"
End With

With Selection
    .InsertAfter "[/quote]"
End With


End Sub


Sub 列表标签()

'选择区域首位加上[list][/list]

With Selection
    .InsertParagraphBefore
End With
  
With Selection
    .InsertBefore "[list]"
End With

With Selection
    .InsertAfter "[/list]"
End With


End Sub

Sub 列表段号()

'选择区域所有段落前加[*]

Dim parag As Paragraph
Dim nLineNum: nLineNum = 0
Dim selRge As Range
Set selRge = Selection.Range
  
  For Each parag In Selection.Paragraphs
  nLineNum = nLineNum + 1
  
  If nLineNum > 0 Then
    selRge.Paragraphs(nLineNum).Range.InsertBefore ("[*]")
  End If
  
 Next

End Sub

Sub 加列表()

列表段号
列表标签

End Sub


Sub 自动链接()

'识别链接，提取URL，在链接文本前后加上[URL]标记

For Each aHyperlink In ActiveDocument.Hyperlinks
        
   If InStr(LCase(aHyperlink.Address), "http") <> 0 Then
        
      aHyperlink.Range.Select
         
    With Selection
      .InsertBefore "[url=" & aHyperlink.Address & "]"
    End With
           
    With Selection
      .InsertAfter "[/url]"
    End With
    
    End If
        
Next aHyperlink


End Sub


Sub 清除格式()

   Selection.ClearFormatting
       
End Sub
Sub 去底纹()


    Selection.WholeStory
    
    去段落底纹
    去文字底纹
    
End Sub
Sub 去文字底纹()
    
    
    With Selection.Font
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(1).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub

Sub 去段落底纹()

  
    With Selection.ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub


Sub 标题样式加粗()


'如果段落样式为指定样式，则在首位加上[b][/b]

Dim cuti As Paragraph
 
  For Each cuti In ActiveDocument.Paragraphs
  
  If cuti.Style = ActiveDocument.Styles("标题 3") Then
  
  cuti.Range.Select
  
  With Selection
      .InsertBefore "[b]"
    End With
           
    With Selection
      .InsertAfter "[/b]"
    End With

  End If
  
 Next


End Sub

Sub 标题长度加粗()


' 要求用户设置长度值

Dim Message, Title, Default, MyValue

Message = "请输入限定的段落文本字/单词数"

Title = "限定长度"

Default = "10"

MyValue = InputBox(Message, Title, Default)

' 如果段落文字长度小于设定值，则在首位加上[b][/b]

Dim cuti As Paragraph
 
  For Each cuti In ActiveDocument.Paragraphs
  
      
  If cuti.Range.Words.Count < MyValue And cuti.Range.Words.Count > 1 Then
  
  
'  Range.Characters.Count < 20 Then
       
  cuti.Range.Select
     
  With Selection
      .InsertBefore "[b]"
    End With
        
   Selection.EndKey Unit:=wdLine
   Selection.TypeText Text:="[/b]"
   Selection.MoveRight Unit:=wdCharacter, Count:=1
      
    
   ' With Selection
   '   .InsertAfter "[/b]"
  '  End With

  End If
   
 Next

End Sub
Sub 清除加粗()

' 清除所有的加粗标记[b][/b]

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[b]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "[/b]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 修复分段()
'
' 文中有不正确的分段标记，该宏可以修复此类问题
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "aaabbbccc"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ".aaabbbccc"
        .Replacement.Text = ".^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "aaabbbccc"
        .Replacement.Text = "   "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub 删空行()



Dim kong As Paragraph
 
  For Each kong In ActiveDocument.Paragraphs
  
      
  If kong.Range.Characters.Count = 1 Then
  
         
  kong.Range.Select
   
  
  Selection.Delete
        
  
  End If
   
 Next

'段首删空格


End Sub
Sub 检查链接()
'
' 检查“[url=”和“http://”中是否有空格，有则删除
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[url= http://"
        .Replacement.Text = "[url=http://"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    With Selection.Find
        .Text = "[url= https://"
        .Replacement.Text = "[url=https://"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub


Sub 取消所有超链接()

'清除所有的超链接


Dim oField As Field

For Each oField In ActiveDocument.Fields
 If oField.Type = wdFieldHyperlink Then
   oField.Unlink
 End If
   
Next
   Set oField = Nothing
End Sub


Sub 选择部分手动换行()

'将选择部分的段落标记替换为手动换行标记

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^l"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub




Sub 周报链接()

'Markup语法（写周报用）：识别链接，提取URL，加上#

For Each aHyperlink In ActiveDocument.Hyperlinks
        
   If InStr(LCase(aHyperlink.Address), "http") <> 0 Then
        
      aHyperlink.Range.Select
         
    With Selection
      .InsertBefore "#[" & aHyperlink.Address & " "
    End With
           
    With Selection
      .InsertAfter "]"
    End With
    
    End If
        
Next aHyperlink


End Sub


Sub 加代码()
  
  
With Selection
    .InsertBefore "[code=""]"
End With

With Selection
    .InsertAfter "[/code]"
End With


End Sub



Sub 超级替换()

'把常见的确实可以自动替换的错别字进行自动替换。
'第一个参数是错别字，第二个参数是正确的字


替换常用错别字 "惟一", "唯一"
替换常用错别字 "帐号", "账号"
替换常用错别字 "图象", "图像"
替换常用错别字 "登陆", "登录"
替换常用错别字 "其它", "其他"
替换常用错别字 "按装", "安装"
替换常用错别字 "按纽", "按钮"
替换常用错别字 "成份", "成分"
替换常用错别字 "题纲", "提纲"
替换常用错别字 "煤体", "媒体"
替换常用错别字 "存贮", "存储"
替换常用错别字 "一桢", "一帧"
替换常用错别字 "好象", "好像"
替换常用错别字 "对像", "对象"


End Sub

Sub 替换常用错别字(strWrong As String, strRight)

'此过程仅供程序调用，不要人手工使用
'
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = strWrong
        .Replacement.Text = strRight
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub 表格转换()

'将表格转换成bbcode表格格式

换表格
每段加竖线
首尾加table


End Sub



Sub 换表格()

' 将表格换为文本


    Application.DefaultTableSeparator = "|"
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, _
        NestedTables:=True
End Sub

Sub 首尾加table()

'选择区域首位加上[table][/table]

With Selection
    .InsertParagraphBefore
End With
  
With Selection
    .InsertBefore "[table]"
End With

With Selection
    .InsertAfter "[/table]"
End With


End Sub


Sub 每段加竖线()

'选择区域所有段落前加|

Dim parag As Paragraph
Dim nLineNum: nLineNum = 0
Dim selRge As Range
Set selRge = Selection.Range
  
  For Each parag In Selection.Paragraphs
  
 
  nLineNum = nLineNum + 1
  
  
  If nLineNum > 0 Then
  

    selRge.Paragraphs(nLineNum).Range.InsertBefore ("|")
        
    Set myrange = selRge.Paragraphs(nLineNum).Range
        
    myrange.End = myrange.End - 1
    
    myrange.InsertAfter ("|")


  End If
  
 Next

End Sub





Sub 段间加空行()

'在段落间加上空行，[list]列表之间不加空行

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p[*]"
        .Replacement.Text = "[*]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[/list]^p^p"
        .Replacement.Text = "[/list]^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
End Sub



Sub 字体红色()
  
  
With Selection
    .InsertBefore "[color=red]"
End With

With Selection
    .InsertAfter "[/color]"
End With


End Sub


Sub 选中手动换行()
'
' 将选中的内容手动换行
'

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^l"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub




Sub 自动链接123()

'识别链接，提取URL，在链接文本前后加上[URL]标记

Dim link

For Each aHyperlink In ActiveDocument.Hyperlinks
        
   If InStr(LCase(aHyperlink.Address), "http") <> 0 Then
        
      aHyperlink.Range.Select
      
      link = aHyperlink.Address
         
     Selection.MoveRight Unit:=wdCell
     
     Selection.TypeText Text:=link

     Selection.MoveRight Unit:=wdCell
    End If
        
Next aHyperlink


End Sub
Sub Macro1()
'
' Macro1 Macro
' 宏在 2014-5-23 由 wg 录制
'
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=10
    Selection.MoveRight Unit:=wdCell
End Sub
Sub Macro2()
'
' Macro2 Macro
' 宏在 2014-5-23 由 wg 录制
'
    Selection.TypeText Text:="123123123"
End Sub


Sub Macro3()
'
' Macro3 Macro
' 宏在 2014-7-9 由 wg 录制
'
    Selection.InlineShapes(1).Fill.Visible = msoFalse
    Selection.InlineShapes(1).Fill.Solid
    Selection.InlineShapes(1).Fill.Transparency = 0#
    Selection.InlineShapes(1).Line.Weight = 0.75
    Selection.InlineShapes(1).Line.Transparency = 0#
    Selection.InlineShapes(1).Line.Visible = msoFalse
    Selection.InlineShapes(1).LockAspectRatio = msoTrue
    Selection.InlineShapes(1).Height = 133.8
   
    Selection.InlineShapes(1).PictureFormat.Brightness = 0.5
    Selection.InlineShapes(1).PictureFormat.Contrast = 0.5
    Selection.InlineShapes(1).PictureFormat.ColorType = msoPictureAutomatic
    Selection.InlineShapes(1).PictureFormat.CropLeft = 56.41
    Selection.InlineShapes(1).PictureFormat.CropRight = 33.45
    Selection.InlineShapes(1).PictureFormat.CropTop = 34.87
    Selection.InlineShapes(1).PictureFormat.CropBottom = 22.68
End Sub




Sub 调整图片大小()


Dim pic As InlineShape

Dim n
 
 For Each pic In ActiveDocument.InlineShapes
 
 If pic.Width <> 0 Then

    pic.Select
       
 Selection.InlineShapes(1).LockAspectRatio = msoTrue
  
 
   pic.Width = 99.2
    

n = pic.ScaleWidth

      
pic.ScaleHeight = n

pic.ScaleWidth = n

      

 
 End If


Next


End Sub
Sub Macro4()
'
' Macro4 Macro
' 宏在 2014-7-9 由 wg 录制
'
    Selection.InlineShapes(1).Fill.Visible = msoFalse
    Selection.InlineShapes(1).Fill.Solid
    Selection.InlineShapes(1).Fill.Transparency = 0#
    Selection.InlineShapes(1).Line.Weight = 0.75
    Selection.InlineShapes(1).Line.Transparency = 0#
    Selection.InlineShapes(1).Line.Visible = msoFalse
    Selection.InlineShapes(1).Width = 99.2
    Selection.InlineShapes(1).PictureFormat.Brightness = 0.5
    Selection.InlineShapes(1).PictureFormat.Contrast = 0.5
    Selection.InlineShapes(1).PictureFormat.ColorType = msoPictureAutomatic
    Selection.InlineShapes(1).PictureFormat.CropLeft = 56.41
    Selection.InlineShapes(1).PictureFormat.CropRight = 33.45
    Selection.InlineShapes(1).PictureFormat.CropTop = 34.87
    Selection.InlineShapes(1).PictureFormat.CropBottom = 22.68
End Sub
