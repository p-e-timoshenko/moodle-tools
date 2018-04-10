Attribute VB_Name = "MoodleGIFT"
' ������ VBA ��� ����������� � ������ Moodle GIFT
'
' �����: �.�. ���������
' ����: 2018-03-28
'
' �������� ��� ������� ���������������� �� �������� "��� ����",
' � ������������ � ��������� GNU GPL v3
' https://www.gnu.org/licenses/quick-guide-gplv3.ru.html
'
' ��������
'
' ������ ������������ ��� �������������� ������, ��������������
'   � ������� ��������� Microsoft Word 2016.
'
' ������ ����������� �������� �������, �������������� � ���������
'  ��������.
'
' - ������ ������ ��������� ��������� � ������� "1." ��� ���� �������
'   ������ �������
' - ����� ������ ��������� ��������� � ������� "a.", "a)"
'   ��� ����� ��� ������ ������ ������ "~", "=".
'   ���������� ����� ���������, ���� �� �������� ������ "="
'     ��� ������������ � ������� ������ ��� ������ �������.
'
' �������:
'
' 1. ������
' �. ������������ �����,
'    ��������� �� 2 �������
' �) ������������ �����,
'    ��������� �� 2 �������
' �. ���������� ����� 2, ���������� ������ ��� ���������� ������.
'
' ������������ ������, ����� ������� ����� ������ ������ ��� �� ������� ������ �������
' ������������ ����� 1
' ������������ ����� 2
' ���������� �����, ���������� ������
' ������������ ����� 3, ����� �������� ����� ������ ������
'
' MOODLE GIFT ������ �������� �������
'
' // ����� ���� ������ ������ ����� ����� �������� �����������
' // ������� ��������� ����� � ������� $CATEGORY:
' $CATEGORY: $module$ / ��������� / ������������
'
' // �������� ������, �������������� ����� ������ ��� ����������
' // ���������� ������� ����� ��������� ������:
' ::��������� �������
' ::���������� �������
' {
' = ���������� �����
' ~ ������������ �����
' ~ ������������ �����
' ~ ������������ �����
' }
'
' ::�01
' ::���� ����� ����� 21/2?
' {
' ~ 10
' = 10.5
' ~ 11
' ~ \=3\\2 + 2\\3
' }
'
' // ����� ��������� ��������� ������ ���� ������ ������.
' // ���� � ������� ��� �������� ������ ���������� ���� ��
' // ����������� �������� "~", "=", "#", "{", "}", and ":", �� �� ������ ���� ����������� "\".
' // ����� ����� ��������� ����������� ���� "\", ����� ������� ��� �� �������
' // �� ���������� ���������.
'
' // ������ ����� ������������ ���� � ����������� ��������.
' // � ������� ���� ������������ ��������� ���� 2 ���������� ��������.
' ::�01
' ::21/2={=10.5 =10,5}.
'
' // ���� ������ �������� ����������� (sub) ��� ����������� (sup) ����� ��� ������� (img),
' // �� ���������� ������� � ����� [html] (��. htmlbook.ru/html/)
' // ��� "p" - ������ ("<p>") � ����� ("</p>") ������
' // ��� "img" - �����������, ��� @@PLUGINFILE@@/images/25.png ��������, ���
' //   � �������� ��� ���������� ��������� ���� ("@@PLUGINFILE@@") ����
' //   ���������� "images", � ������� ��������� ���� "25.png",
' //   ���������� ����������� � ������� PNG.
' ::�02
' ::[html]<p>���� ����� ����� 3<sup>2</sup> + 4<sup>2</sup><p>
' {
' ~24
' =[html]<p>5<sup>2</sup></p>
' =[html]<p><img scr\="@@PLUGINFILE@@/images/25.png" /></p>
' }
'
' ���������
'  ��� ��������� ������ ���������� � Microsoft Word 2016 ���������
'    ��������� ��������:
'  1. �������� ������� "�����������"
'  - ������� ���������� ���� "��������� Word" (���� "����" - "���������")
'  - � ������ ������� ����� ������� "��������� �����"
'  - �� ������� "��������� ����� � ��������� ������" � ������ ������
'    "�������� �������" ��������� ������� �������� ������ "�����������"
'  - ����� �� ������ "��", ������� ����
'  2. �������� ��������� ��������� ������������ � ������� � �������� VBA
'  - ������� � ���� "�����������"
'  - � ����� "���" ������ �� ������ "������������ ��������" � �������
'    � ���������� ���� "����� ���������� �������������".
'  - � ������ ������� ����� ������� "��������� ��������".
'  - � ����������� ������� ������� �����
'    "�������� ��� ������� (�� �������������, �������� ������ ������� ���������)",
'    ��������� ������� ����� � ������� "�������� ������ � ��������� ������ �������� VBA"
'    �, ����� �� ������ "��", ������� ����.
'    ��������! ����� ����� � ����� ������ �� ������������� ����
'      �� ���������� ������ � �������� ��� ������� � ��������� ������������
'      ������� � �������������� ���������!
'  - ������ �� ������ "Visual Basic" (���� "�����������")
'  3. �������� ������ MoodleGIFT
'  - � ���������� ������� "Project" (���� "View" - "Project Explorer", Ctrl+R)
'    ������� � ������� "Normal" ����� "Modules".
'  - �� ��������� �������� ������� ����������� ����, ����� ������ ������� ����,
'    � ������� "Insert" - "Module". ����� ���������� �������� ������������� ������
'    �������, ������� ��� � ������ � ����������� ���� ����� "Remove MoodleGIFT...".
'  - ����������� ����� ������� ������������� � "MoodleGIFT" ����� �
'    ���� ������� "Properties" (���� "View" - "Properties Window", F4).
'  - �������� ��� ���������� � ���� ��������� ���� ������ "MoodleGIFT",
'    ������ ��� ������� ������� ����� ������� ���� �� �������� "MoodleGIFT"
'    � ���������� �������.
'  4. �������� ��������� ���������� ���������
'  - � ���� "Tools" ������� "References..."
'  - � ����������� ���������� ���� � ������ "Available References" ���������� �������
'  - �������� ������ "Microsoft VBScript Regular Expressions 5.5". ������������� �����
'    ��������� ������� ������� �������� "Visual Basic For Applications", "OLE Automation",
'    "Microsoft Office 16.0 Object Library" � "Microsoft Word 16.0 Object Library".
'
'
' �������������
'
' ��� ���������� �������� ���������� � ����� "���" ������ "�����������"
'  ������ �� ������ "�������" � �������  � ���������� ���� ���������� ��������.
'  � ������ ���������� ����������� ������ � ���������� ������ "���������".
'  �� ���������� ������� ����� ������������� ��������������� ����� (�� 5 ���),
'   � ������� �������� Word ����� "������".
'
' ������������� ������������� ���� ���������, ����� ����� �������� ���������
'  ������ ������� ����������� ��������, �����, ����� � ���� �������������.
'
' ��� �������������� ������ � ������ Moodle GIFT ��������� ���������
'  � Microsoft Word 2016 �� ����������� ��������� ���������,
'  �������� �� ������ ���� ���������� �� ����������:
'
' - PrepareForConversionToMoodleGIFT
'   ��� 1. ���������� � ����������� � ������ Moodle GIFT.
'   � ������ ��������� ����������� ����������
'     ���������� ��������� ����� ��������.
'   ����� ���������� ����� ���� ��� �������, ������� � �������
'   ����� ������������ � ���� PNG �����������, ������� �������, ��������,
'   ����������� �������� ����� ����������� ���������� ����.
'   ��������� ����� ������������� � ������������ � ������������ �������:
'     - Test[Category] - ����������� ������ ��� ����������� �������� ���������
'       � ������� "������������ ��������� / ��������� ��������� / ��������� ������������".
'       ����� ������� ������ ������ �������.
'     - Test[Question] - �����, ���������� �������� ������. ���� �������� ������ ��������
'       ��������� �������, �� ��������� ������ ���������� ������ Test[Continue].
'     - Test[RightAnswer] - �����, ���������� ���������� �����. ���� ����� ��������
'       ��������� �������, �� ��������� ������ ���������� ������ Test[Continue].
'     - Test[WrongAnswer] - �����, ���������� ������������ �����. ���� ����� ��������
'       ��������� �������, �� ��������� ������ ���������� ������ Test[Continue].
'     - Test[Continue] - �����, ���������� ������������ ����������� �������� ��������� �������.
'   ����� ����� ��������� � ������� ��������� ���������, ������ �� ����������,
'   �� �������� �� �������� �� �������������, �.�. ��� ����� � �������� �����������.
'
' - ExportPictures
'   ��� 2. ������� ����������� � ��������� ������� � ������� PNG
'   � ������ ���������� �������� ����������
'     ���������� ��������� �������� ������.
'   ���� ��� ����� ����������, ���� �������������� ��������
'     �� �������� �����������.
'
' - ConvertToHtml
'   ��� 3. ����������� ����������� ��������� � ������ HTML.
'   � ����� ��������� ��������� ����������� � ����������� �����,
'     ������� � ������� "&", "<", ">".
'   �� ���� ���� ���������� ���������� � ��������� ��������� � ������� HTML
'   ������� "[html]", ����� ������ ("<p>") � ��������� ("</p>") ������.
'
' - EscapeSpecialChars
'   ��� 4. ������������� ����������� �������� "\", "~", "=", "{", "}", "#", ":".
'   ����� ����� ��������� ����������� ���� "\", ����� ������� ��� �� �������
'   �� ���������� ���������.
'
' - ConvertToMoodleGIFT
'   ��� 5. ����������� � ������ Moodle GIFT.
'   ���������� ����������� ����� ��������� � ��������� ����� � ���������,
'     ����������� �������� ���������.
'
' - ConvertToMoodleGIFTWithMedia [�� �����������]
'   ��� 6. ���������. ���� ��� ����������� � ������, ���� ���� �����������.
'   ������������ ��������� ���� � ������� � �������������
'
' ��������! ���� � ���������� ���������� ���������� ���� ���������� ����������
'  ���������� ���������, �� ����� ��� ����������� ���������� ���������� ������
'  � �����, �������� ������������ ������ � ��� �������.
'

Option Explicit
Option Compare Text

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ��� 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub PrepareForConversionToMoodleGIFT()
' ���������� � ����������� � ������ Moodle GIFT
'
' �������� ����������� �� ������ ����
'
' ������: PrepareForConversionToMoodleGIFT
' �����: �.�. ���������
' ����: 2018-03-28
    
    
    ' �������� ���������� ������
    Dim oldScrUpd As Boolean
    oldScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' ����������� ������������� ������
    Application.UndoRecord.StartCustomRecord "PrepareForConversionToMoodleGIFT"
    
    On Error GoTo ErrorHandler
    
    Dim oDoc As Word.Document
    Set oDoc = ActiveDocument
        
    oDoc.Activate
    
    InlineAllShapes oDoc
    ConvertShapesAndTablesToPNG oDoc
    ConvertHyperlinksToPlainText oDoc
    
    CleanDocument oDoc
    FixParagraphSymbolsFormatting oDoc
    
    InitStyles oDoc
    
    ProcessParagraphsAndApplyStyles oDoc
    DeleteEmptyParagraphs oDoc
    
    Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    
Finish:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = oldScrUpd
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        Dim Msg As String
        Msg = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
    
    'Resume Next
    GoTo Finish
End Sub

Private Sub InlineAllShapes(oDoc As Document)
' �������������� ���� ��������� �������� �� ���������������

    Dim oShape As Shape
    For Each oShape In oDoc.Shapes
        oShape.ConvertToInlineShape
    Next
End Sub

Private Sub ConvertShapesAndTablesToPNG(oDoc As Document)
' �������������� ���� ��������� � �����������
'
' ��������� ��������� ��������� ������������������ ��������:
' 1. �������������� ������, ������, ����� � ���������� ����� � EMF ������
' 2. �������������� EMF ������� � PNG (DataType = 14)
'
    If oDoc.Shapes.Count = 0 _
        And oDoc.InlineShapes.Count = 0 _
        And oDoc.Tables.Count = 0 _
        And oDoc.OMaths.Count = 0 Then
        Exit Sub
    End If
    
    Dim nObj As Integer
    
    For nObj = oDoc.Tables.Count To 1 Step -1
        oDoc.Tables(nObj).Select
        
        Selection.Cut
        Selection.PasteSpecial _
            link:=False, _
            DataType:=wdPasteEnhancedMetafile, _
            Placement:=wdInLine, _
            DisplayAsIcon:=False
    Next

    For nObj = oDoc.OMaths.Count To 1 Step -1
        oDoc.OMaths(nObj).Range.Select
        
        Selection.Cut
        Selection.PasteSpecial _
            link:=False, _
            DataType:=wdPasteEnhancedMetafile, _
            Placement:=wdInLine, _
            DisplayAsIcon:=False
    Next
    
    For nObj = oDoc.Shapes.Count To 1 Step -1
        oDoc.Shapes(nObj).Select
        
        Selection.Cut
        Selection.PasteSpecial _
            link:=False, _
            DataType:=wdPasteEnhancedMetafile, _
            Placement:=wdInLine, _
            DisplayAsIcon:=False
    Next

    For nObj = oDoc.InlineShapes.Count To 1 Step -1
        oDoc.InlineShapes(nObj).Select
        
        Selection.Cut
        Selection.PasteSpecial _
            link:=False, _
            DataType:=wdPasteEnhancedMetafile, _
            Placement:=wdInLine, _
            DisplayAsIcon:=False
    Next
    
    For nObj = oDoc.InlineShapes.Count To 1 Step -1
        oDoc.InlineShapes(nObj).Select
        
        Selection.Cut
        Selection.PasteSpecial _
            link:=False, _
            DataType:=14, _
            Placement:=wdInLine, _
            DisplayAsIcon:=False
    Next
End Sub

Sub ConvertHyperlinksToPlainText(oDoc As Document)
' �������������� ����������� � �����
    
    Dim nHyperlink As Long
    
    For nHyperlink = ActiveDocument.Hyperlinks.Count To 1 Step -1
        oDoc.Hyperlinks(nHyperlink).Delete
    Next
End Sub

Private Sub CleanDocument(oDoc As Word.Document)
' ������ ����������� ���������
        
    Dim oFind As Word.Find
    'Set oFind = oDoc.Content.Find
    
    'Fix the skipped blank Header/Footer problem
    Dim lngJunk As WdStoryType
    lngJunk = oDoc.Sections(1).Headers(1).Range.StoryType
    Set oFind = oDoc.StoryRanges(wdMainTextStory).Find
    
    Dim pr As Variant
  
    '  ������ ������ ��������� ��������� ��������:
    '  - �������� ������ ���������,
    '  - ������ ����������� �������� ��������
    '  - ������ �������� ��������� ���������,
    '  - ������ �������� ������� ������ ������,
    '  - �������� ������������� ��������,
    '  - �������� ������� ������� � ����� ������
    For Each pr In Array( _
        Array("^-", "", False), _
        Array(Chr(160), " ", False), _
        Array("^t", " ", False), _
        Array("^l", "^p", False), _
        Array("[ ]{2;}", " ", True), _
        Array(" ^p", "^p", False))
        
        With oFind
            .ClearFormatting
            .Replacement.ClearFormatting
        
            .Text = pr(0)
            .Replacement.Text = pr(1)
            
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = pr(2)
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            .Execute Replace:=wdReplaceAll
        End With
    Next
End Sub

Private Sub FixParagraphSymbolsFormatting(oDoc As Word.Document)
' ���������� �������������� �������� ����� ������

    Dim oFind As Word.Find
    'Set oFind = oDoc.Content.Find
    
    'Fix the skipped blank Header/Footer problem
    Dim lngJunk As WdStoryType
    lngJunk = oDoc.Sections(1).Headers(1).Range.StoryType
    Set oFind = oDoc.StoryRanges(wdMainTextStory).Find
    
    With oFind
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = "^p"
        .Replacement.Text = "^p"
        With .Replacement.Font
            .Bold = False
            .Italic = False
            .Color = wdColorAutomatic
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
            
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function FindStyle(oStyles As Styles, Name As String) As Style
' ����� ����� �� ��������
    Dim oStyle As Style
    Set FindStyle = Nothing
    
    For Each oStyle In oStyles
        If oStyle.NameLocal = Name Then
            Set FindStyle = oStyle
            Exit For
        End If
    Next
End Function

Private Function FindListTemplate(oListTemplates As ListTemplates, Name As String) As ListTemplate
' ����� ����� �� ��������
    Dim lt As ListTemplate
    Set FindListTemplate = Nothing
    
    For Each lt In oListTemplates
        If lt.Name = Name Then
            Set FindListTemplate = lt
            Exit For
        End If
    Next
End Function

Private Sub InitStyles(oDoc As Document)
' ��������� �����

    Dim oStyles As Styles
    Set oStyles = oDoc.Styles
    
    CreateTestBasicStyle oStyles
    CreateTestWrongAnswerStyle oStyles
    CreateTestRightAnswerStyle oStyles
    CreateTestQuestionStyle oStyles
    CreateTestCategoryStyle oStyles
    CreateTestContinueStyle oStyles
    
    CreateTestQuestionListStyle oDoc
    CreateTestRightAnswerListStyle oDoc
    CreateTestWrongAnswerListStyle oDoc
End Sub

Private Function CreateTestBasicStyle(oStyles As Styles, Optional StyleName As String)
' ������� ������� �������� �����

    If IsMissing(StyleName) Or Len(StyleName) = 0 Then
        StyleName = "Test[Basic]"
    End If
    
    Set CreateTestBasicStyle = Nothing
    If Not (FindStyle(oStyles, StyleName) Is Nothing) Then
        Exit Function
    End If
 
    oStyles.Add Name:=StyleName, Type:=wdStyleTypeParagraph
    
    Set CreateTestBasicStyle = oStyles(StyleName)
    With CreateTestBasicStyle
        .AutomaticallyUpdate = False
        
        If StyleName <> "Test[Basic]" Then
            .BaseStyle = "Test[Basic]"
            .UnhideWhenUsed = True
            .QuickStyle = True
        Else
            .UnhideWhenUsed = False
            .QuickStyle = False
        End If
        .NextParagraphStyle = StyleName
        
        With .Font
            .Name = "Times New Roman"
            .Size = 12
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            .UnderlineColor = wdColorAutomatic
            .StrikeThrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .SmallCaps = False
            .AllCaps = False
            .Color = wdColorAutomatic
            .Engrave = False
            .Superscript = False
            .Subscript = False
            .Scaling = 100
            .Kerning = 0
            .Animation = wdAnimationNone
            .Ligatures = wdLigaturesNone
            .NumberSpacing = wdNumberSpacingDefault
            .NumberForm = wdNumberFormDefault
            .StylisticSet = wdStylisticSetDefault
            .ContextualAlternates = 0
        End With
        
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphJustify
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = False
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            .CollapsedByDefault = False
            
            .TabStops.ClearAll
            
            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
        
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            
            With .Borders
                .DistanceFromTop = 1
                .DistanceFromLeft = 4
                .DistanceFromBottom = 1
                .DistanceFromRight = 4
                .Shadow = False
            End With
        End With
    
        .NoSpaceBetweenParagraphsOfSameStyle = True
        .NoProofing = False
        
        .LinkToListTemplate ListTemplate:=Nothing
        
        .Frame.Delete
    End With
End Function

Private Sub CreateTestCategoryStyle(oStyles As Styles)
' ������� ����� �������� ���������

    Dim oStyle As Style
    Set oStyle = CreateTestBasicStyle(oStyles, "Test[Category]")
    If oStyle Is Nothing Then
        Exit Sub
    End If
    
    With oStyle
        .AutomaticallyUpdate = False
        .NextParagraphStyle = "Test[Question]"
        
        With .Font
            .Name = "Arial"
        End With
        
        With .ParagraphFormat
            .SpaceBefore = 12
            .SpaceBeforeAuto = False
            .SpaceAfter = 12
            .SpaceAfterAuto = False
            .Alignment = wdAlignParagraphCenter
        End With
        
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .Color = wdColorAutomatic
            .LineWidth = wdLineWidth025pt
        End With
        
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .Color = wdColorAutomatic
            .LineWidth = wdLineWidth025pt
        End With
    End With
End Sub

Private Sub CreateTestQuestionStyle(oStyles As Styles)
' ������� ����� ��������� �������

    Dim oStyle As Style
    Set oStyle = CreateTestBasicStyle(oStyles, "Test[Question]")
    If oStyle Is Nothing Then
        Exit Sub
    End If
    
    With oStyle
        .AutomaticallyUpdate = False
        .NextParagraphStyle = "Test[RightAnswer]"
        
        With .ParagraphFormat
            .SpaceBefore = 24
            .SpaceAfter = 6
        End With
    End With
    
End Sub

Private Sub CreateTestRightAnswerStyle(oStyles As Styles)
' ������� ����� ����������� ������ �� �������� ������

    Dim oStyle As Style
    Set oStyle = CreateTestBasicStyle(oStyles, "Test[RightAnswer]")
    If oStyle Is Nothing Then
        Exit Sub
    End If
    
    With oStyle
        .AutomaticallyUpdate = False
        .NextParagraphStyle = "Test[WrongAnswer]"
            
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(1)
            .RightIndent = CentimetersToPoints(1)
        End With
    End With
End Sub

Private Sub CreateTestContinueStyle(oStyles As Styles)
' ������� ����� ���������� ������� ��� ������

    Dim oStyle As Style
    Set oStyle = CreateTestBasicStyle(oStyles, "Test[Continue]")
    If oStyle Is Nothing Then
        Exit Sub
    End If
    
    With oStyle
        .AutomaticallyUpdate = False
        .NextParagraphStyle = "Test[Continue]"
        
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 6
            
            .LeftIndent = CentimetersToPoints(1)
            .RightIndent = CentimetersToPoints(1)
            
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            
            With .Borders
                .DistanceFromLeft = CentimetersToPoints(0.5)
            End With
        End With
    End With
    
End Sub

Private Sub CreateTestWrongAnswerStyle(oStyles As Styles)
' ������� ����� ������������� ������ �� �������� ������
   
    Dim oStyle As Style
    Set oStyle = CreateTestBasicStyle(oStyles, "Test[WrongAnswer]")
    If oStyle Is Nothing Then
        Exit Sub
    End If
    
    With oStyle
        .AutomaticallyUpdate = False
            
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(1)
            .RightIndent = CentimetersToPoints(1)
        End With
    End With
End Sub

Private Function CreateTestListStyle(oDoc As Document, ListName As String, LinkedStyleName As String) As ListTemplate
' ������� ����� ������� ������������� ������ ��� ��������
    Dim oStyles As Styles
    Set oStyles = oDoc.Styles
    
    Set CreateTestListStyle = Nothing
    If Not (FindStyle(oStyles, ListName) Is Nothing) Then
        Exit Function
    End If
    
    Dim lt As ListTemplate
    Set lt = FindListTemplate(oDoc.ListTemplates, ListName)
    If lt Is Nothing Then
        Set lt = oDoc.ListTemplates.Add(False)
        lt.Name = ListName
    End If

    oStyles.Add ListName, wdStyleTypeList
    oStyles(ListName).LinkToListTemplate lt, 1
    
    With lt.ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        If Len(LinkedStyleName) > 0 Then
            .LinkedStyle = LinkedStyleName
        End If
    End With
    
    Set CreateTestListStyle = lt
End Function

Private Sub CreateTestQuestionListStyle(oDoc As Document)
' ������� ����� ������� ������������� ������ ��� ��������
    CreateTestListStyle oDoc, "Test[QuestionNumbering]", "Test[Question]"
End Sub

Private Sub CreateTestRightAnswerListStyle(oDoc As Document)
' ������� ����� ������� ������ ��� ���������� �������
    Dim lt As ListTemplate
    Set lt = CreateTestListStyle(oDoc, "Test[RightAnswerSymbol]", "Test[RightAnswer]")
    
    If lt Is Nothing Then
        Exit Sub
    End If
    
    With lt.ListLevels(1)
        .NumberFormat = "="
        .NumberStyle = wdListNumberStyleOrdinalText
        
        With .Font
            .Color = wdColorGreen
        End With
    End With
End Sub

Private Sub CreateTestWrongAnswerListStyle(oDoc As Document)
' ������� ����� ������� ������ ��� ������������ �������
    Dim lt As ListTemplate
    Set lt = CreateTestListStyle(oDoc, "Test[WrongAnswerSymbol]", "Test[WrongAnswer]")
    
    If lt Is Nothing Then
        Exit Sub
    End If
    
    With lt.ListLevels(1)
        .NumberFormat = "~"
        .NumberStyle = wdListNumberStyleOrdinalText
        
        With .Font
            .Color = wdColorRed
        End With
    End With
End Sub

Private Function GetParagraphType(oParagraph As Paragraph) As String
' ��� ������
    
    GetParagraphType = ""
    
    If Len(oParagraph.Range.Text) = 1 Then
        GetParagraphType = "empty"
        Exit Function
    End If
        
    If Not IsEmpty(oParagraph.Style) Then
        Select Case oParagraph.Style
            Case "Test[Category]": GetParagraphType = "styled_category"
            Case "Test[Question]": GetParagraphType = "styled_question"
            Case "Test[RightAnswer]": GetParagraphType = "styled_answer"
            Case "Test[WrongAnswer]": GetParagraphType = "styled_answer"
            Case "Test[Continue]": GetParagraphType = "styled_continue"
        End Select
        
        If Len(GetParagraphType) > 0 Then
            Exit Function
        End If
    End If
    
    If ProcessOrderedQuestionParagraph(oParagraph, ApplyChanges:=False) Then
        GetParagraphType = "ordered_question"
        Exit Function
    End If
    
    If ProcessOrderedAnswerParagraph(oParagraph, ApplyChanges:=False) Then
        GetParagraphType = "ordered_answer"
        Exit Function
    End If
    
    If oParagraph.Range.Font.Bold Then
        GetParagraphType = "highlighted_question"
        Exit Function
    End If
    
    If oParagraph.Range.Font.Color <> wdColorAutomatic _
        Or oParagraph.Range.HighlightColorIndex <> wdAuto Then
        GetParagraphType = "highlighted_answer"
        Exit Function
    End If
    
End Function

Private Function ProcessOrderedQuestionParagraph( _
    oParagraph As Paragraph, _
    Optional ApplyChanges As Boolean = True) As Boolean
'��������� ������������ �������� � ������� "1. ..."

    Dim re As New RegExp
    With re
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^([\n\r]*)(\s*\d+\.)"
    End With
        
    Dim m As Object
    Set m = re.Execute(oParagraph.Range.Text)
    
    ProcessOrderedQuestionParagraph = m.Count = 1
    If ProcessOrderedQuestionParagraph Then
        ProcessOrderedQuestionParagraph = m(0).SubMatches.Count = 2
    End If
    
    If ProcessOrderedQuestionParagraph And ApplyChanges Then
        Dim Offset, Count As Long
        Offset = Len(m(0).SubMatches(0)) + 1
        Count = Len(m(0).SubMatches(1))
        
        oParagraph.Range.Characters(Offset).Delete wdCharacter, Count
        While oParagraph.Range.Characters(Offset).Text = " "
            oParagraph.Range.Characters(Offset).Delete wdCharacter, 1
        Wend
    End If
End Function

Private Function ProcessOrderedAnswerParagraph( _
    oParagraph As Paragraph, _
    Optional ApplyChanges As Boolean = True) As Boolean
'��������� �������, �������������� � ��������� �����:
' "a) ...", "a. ...", "1) ... ", "~ ...", "= ..."
    
    Dim re As New RegExp
    With re
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^([\n\r]*)(\s*(?:[A-Za-z�-��-�]{1}[\)\.]|\d+\)|[~=]))"
    End With
    
    Dim m As Object
    Set m = re.Execute(oParagraph.Range.Text)
    
    ProcessOrderedAnswerParagraph = m.Count = 1
    If ProcessOrderedAnswerParagraph Then
        ProcessOrderedAnswerParagraph = m(0).SubMatches.Count = 2
    End If
    
    If ProcessOrderedAnswerParagraph And ApplyChanges Then
        Dim Offset, Count As Long
        Offset = Len(m(0).SubMatches(0)) + 1
        Count = Len(m(0).SubMatches(1))
        
        oParagraph.Range.Characters(Offset).Delete wdCharacter, Count
        While oParagraph.Range.Characters(Offset).Text = " "
            oParagraph.Range.Characters(Offset).Delete wdCharacter, 1
        Wend
        
        RemoveLastDelimiterInAnswerParagraph oParagraph
        
        ProcessOrderedAnswerParagraph = Trim(m(0).SubMatches(1)) = "="
    End If
End Function

Private Sub RemoveLastDelimiterInAnswerParagraph(oParagraph As Paragraph)
'�������� ������ "." ��� ";" � ����� ������
    
    Dim re As New RegExp
    With re
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "([;,\.][ \t]*)([\n\r]*)$"
    End With
    
    Dim m As Object
    Set m = re.Execute(oParagraph.Range.Text)
    
    Dim ProcessAnswer As Boolean
    ProcessAnswer = m.Count = 1
    If ProcessAnswer Then
        ProcessAnswer = m(0).SubMatches.Count = 2
    End If
    
    If ProcessAnswer Then
        Dim Offset, Count As Long
        Offset = oParagraph.Range.Characters.Count - Len(m(0).SubMatches(1))
        Count = Len(m(0).SubMatches(0))
        
        oParagraph.Range.Characters(Offset).Delete wdCharacter, Count
    End If
End Sub


Private Sub ProcessParagraphsAndApplyStyles(oDoc As Word.Document)
' ��������� ������� � ���������� ������

    Dim nParagraph As Integer
    Dim oParagraph As Word.Paragraph
    
    Dim PreviousParagraphType, ParagraphType As String
    ParagraphType = "empty"
    
    'TestTask
    Dim PreviousQuestionType As String
    Dim QuestionType As String
    Dim AnswerType As String
    Dim StartParagraphIndex As Integer
    Dim EndParagraphIndex As Integer

    'Init TestTask
    PreviousQuestionType = ""
    QuestionType = ""
    AnswerType = ""
    StartParagraphIndex = -1
    EndParagraphIndex = -1
    
    For nParagraph = 1 To oDoc.Paragraphs.Count
        PreviousParagraphType = ParagraphType
    Do
        Set oParagraph = oDoc.Paragraphs(nParagraph)
        ParagraphType = GetParagraphType(oParagraph)
        Select Case ParagraphType
            Case "empty": Exit Do
            
            Case "styled_category"
                If StartParagraphIndex >= 0 Then
                    EndParagraphIndex = nParagraph - 1
                End If
                
                ProcessTaskParagraphsAndApplyStyles oDoc, _
                    PreviousQuestionType, QuestionType, AnswerType, _
                    StartParagraphIndex, EndParagraphIndex
                
                PreviousQuestionType = QuestionType
                QuestionType = ""
                AnswerType = ""
                StartParagraphIndex = -1
                EndParagraphIndex = -1
                Exit Do
                
            Case Is = "styled_question", Is = "ordered_question", Is = "highlighted_question"
                If StartParagraphIndex >= 0 Then
                    EndParagraphIndex = nParagraph - 1
                End If
                
                ProcessTaskParagraphsAndApplyStyles oDoc, _
                    PreviousQuestionType, QuestionType, AnswerType, _
                    StartParagraphIndex, EndParagraphIndex
                
                PreviousQuestionType = QuestionType
                QuestionType = ParagraphType
                AnswerType = ""
                StartParagraphIndex = nParagraph
                EndParagraphIndex = -1
                Exit Do
                
            Case "styled_answer"
                If Len(AnswerType) = 0 Then
                    AnswerType = ParagraphType
                End If
                Exit Do
                
            Case "ordered_answer"
                Select Case AnswerType
                    Case Is = "", _
                         Is = "styled_answer", _
                         Is = "highlighted_answer"
                        AnswerType = ParagraphType
                End Select
                Exit Do
                
            Case "highlighted_answer"
                Select Case AnswerType
                    Case Is = "", Is = "styled_answer"
                    
                    AnswerType = ParagraphType
                End Select
                Exit Do
                
            Case "styled_continue"
                Exit Do
                
            Case ""
                '������� ����������� ������ �������
                If (Len(PreviousQuestionType) = 0 Or PreviousQuestionType = "empty") _
                And PreviousParagraphType = "empty" Then
                    If StartParagraphIndex >= 0 Then
                        QuestionType = "empty"
                        EndParagraphIndex = nParagraph - 1
                    End If
                    
                    ProcessTaskParagraphsAndApplyStyles oDoc, _
                        PreviousQuestionType, QuestionType, AnswerType, _
                        StartParagraphIndex, EndParagraphIndex
                
                    PreviousQuestionType = QuestionType
                    QuestionType = "empty"
                    AnswerType = ""
                    StartParagraphIndex = nParagraph
                    EndParagraphIndex = -1
                    Exit Do
                End If
                
                Exit Do
        End Select
    Loop While False
    
    '���������� ��������� ������
    If (nParagraph = oDoc.Paragraphs.Count) And (StartParagraphIndex >= 0) Then
        EndParagraphIndex = nParagraph
        ProcessTaskParagraphsAndApplyStyles oDoc, _
            PreviousQuestionType, QuestionType, AnswerType, _
            StartParagraphIndex, EndParagraphIndex
    End If
    Next
End Sub

Private Sub ProcessTaskParagraphsAndApplyStyles(oDoc As Word.Document, _
    PreviousQuestionType As String, _
    QuestionType As String, _
    AnswerType As String, _
    StartParagraphIndex As Integer, _
    EndParagraphIndex As Integer)
' ��������� ������� ��������� �������� ���������� ������

    If Len(QuestionType) = 0 _
    Or StartParagraphIndex < 0 _
    Or EndParagraphIndex < 0 Then
        Exit Sub
    End If
    
    Dim nParagraph As Integer
    Dim oParagraph As Word.Paragraph
    Dim ParagraphType As String
    Dim IsHighlighted As Boolean
            
    For nParagraph = StartParagraphIndex To EndParagraphIndex
        Set oParagraph = oDoc.Paragraphs(nParagraph)
        ParagraphType = GetParagraphType(oParagraph)
    Do
        Select Case ParagraphType
            Case Is = "empty", _
                Is = "styled_category", _
                Is = "styled_question", _
                Is = "styled_answer", _
                Is = "styled_continue": Exit Do
                
            Case Is = "ordered_question", _
                Is = "highlighted_question"
                
                With oParagraph.Range
                    .Font.Bold = False
                    .Font.Color = wdColorAutomatic
                    .HighlightColorIndex = wdAuto
                End With
                
                If ParagraphType = "ordered_question" Then
                    ProcessOrderedQuestionParagraph oParagraph, ApplyChanges:=True
                End If
                
                oParagraph.Style = "Test[Question]"
                Exit Do
                
            Case Is = "ordered_answer", _
                Is = "highlighted_answer"
                                
                With oParagraph.Range
                    IsHighlighted = .Font.Color <> wdColorAutomatic _
                        Or .HighlightColorIndex <> wdAuto _
                        Or (ParagraphType = "ordered_answer" And .Font.Bold)
                    
                    .Font.Bold = False
                    .Font.Color = wdColorAutomatic
                    .HighlightColorIndex = wdAuto
                End With
                
                If ParagraphType = "ordered_answer" Then
                    IsHighlighted = IsHighlighted _
                        Or ProcessOrderedAnswerParagraph(oParagraph, ApplyChanges:=True)
                End If
                
                If IsHighlighted Then
                    oParagraph.Style = "Test[RightAnswer]"
                Else
                    oParagraph.Style = "Test[WrongAnswer]"
                End If
                
                Exit Do
            Case Is = ""
                Select Case QuestionType
                Case Is = "empty"
                    If nParagraph = StartParagraphIndex Then
                        ParagraphType = "highlighted_question"
                    Else
                        ParagraphType = "highlighted_answer"
                    End If
                    
                Case Is = "highlighted_question"
                    If AnswerType = "highlighted_answer" Then
                        ParagraphType = "highlighted_answer"
                    Else
                        oParagraph.Style = "Test[Continue]"
                        Exit Do
                    End If
                Case Else
                    oParagraph.Style = "Test[Continue]"
                    Exit Do
                End Select
        End Select
    Loop While True: Next
End Sub

Private Sub DeleteEmptyParagraphs(oDoc As Word.Document)
' �������� ������ �������

    Dim oParagraph As Word.Paragraph
    
    For Each oParagraph In oDoc.Paragraphs
        If Len(oParagraph.Range.Text) = 1 Then
            oParagraph.Range.Delete
        End If
    Next
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ��� 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ExportPictures()
'�������������� �����������
'
' �������� ����������� �� ������ ����
'   ����� �������� ����������� ���������� ���������.
' �������� ����������� �������� ������ �������� �����������
'   ����� ����������� ���� ��������.
'
' ������: ExportPictures
' �����: �.�. ���������
' ����: 2018-03-28
    
    ' �������� ���������� ������
    Dim oldScrUpd As Boolean
    oldScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    
    oDoc.Activate

    Dim alertStatus As WdAlertLevel
    alertStatus = Application.DisplayAlerts
    Application.DisplayAlerts = wdAlertsNone

    Dim docPath As String
    Dim htmPath As String

    docPath = oDoc.FullName
    htmPath = docPath & ".htm"

    oDoc.Save
    
    Dim ViewType As WdViewType
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ViewType = ActiveWindow.ActivePane.View.Type
    Else
        ViewType = ActiveWindow.View.Type
    End If
    
    If Len(Dir(docPath & ".files", vbDirectory)) > 0 Then
        Kill docPath & ".files\*"
        RmDir docPath & ".files"
    End If
    
    If Len(Dir(htmPath)) > 0 Then
        Kill htmPath
    End If

    oDoc.SaveAs2 FileName:=htmPath, _
        FileFormat:=wdFormatFilteredHTML, _
        LockComments:=False, _
        Password:="", _
        AddToRecentFiles:=False, _
        WritePassword:="", _
        ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, _
        SaveFormsData:=False, _
        SaveAsAOCELetter:=False
        
    oDoc.Close
    
    If Dir(docPath & ".files", vbDirectory) <> "" Then
        Kill docPath & ".files\*.xml"
        Kill docPath & ".files\*.html"
        Kill docPath & ".files\*.thmx"
    End If
    
    If Len(Dir(htmPath)) > 0 Then
        Kill htmPath
    End If
    
    Word.Documents.Open docPath
    
    Application.DisplayAlerts = alertStatus
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = ViewType
    Else
        ActiveWindow.View.Type = ViewType
    End If
    
    Word.Application.Visible = True
    Word.Application.Activate
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ��� 3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ConvertToHtml()
' ����������� ����������� ��������� � ������ HTML
'
' �������� ����������� �� ������� ����
'   ����� �������� ����������� ���������� ���������
'   � �������� �����������.
'
' ������: ConvertToHtml
' �����: �.�. ���������
' ����: 2018-03-28

    
    
    ' �������� ���������� ������
    Dim oldScrUpd As Boolean
    oldScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' ����������� ������������� ������
    Application.UndoRecord.StartCustomRecord "ConvertToHtml"
    
    On Error GoTo ErrorHandler
    
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    
    oDoc.Activate
    
    ConvertEntitiesToHtml oDoc
    ConvertPNGsToImgTag oDoc
    ProcessStyledParagraphsContainedHTML oDoc
    
Finish:
    ' ����������� ������������� ������
    Application.UndoRecord.EndCustomRecord
    
    ' ��������������� ���������� ������
    Application.ScreenUpdating = oldScrUpd
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        Dim Msg As String
        Msg = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
    
    'Resume Next
    GoTo Finish
End Sub

Private Sub ConvertEntitiesToHtml(oDoc As Word.Document)
' ���������� ���������
    
    Dim oFind As Word.Find
    'Set oFind = oDoc.Content.Find
    
    'Fix the skipped blank Header/Footer problem
    Dim lngJunk As WdStoryType
    lngJunk = oDoc.Sections(1).Headers(1).Range.StoryType
    Set oFind = oDoc.StoryRanges(wdMainTextStory).Find
    
    Dim pr As Variant
    
    '������ ����������� ��������
    For Each pr In Array( _
        Array("&", "&amp;", False), _
        Array("<", "&lt;", False), _
        Array(">", "&gt;", False))
        
        With oFind
            .ClearFormatting
            .Replacement.ClearFormatting
        
            .Text = pr(0)
            .Replacement.Text = pr(1)
            
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = pr(2)
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            .Execute Replace:=wdReplaceAll
        End With
    Next
    
    With oFind
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = ""
        .Font.Superscript = True
        
        .Replacement.Text = "<sup>^&</sup>"
        .Replacement.Font.Superscript = False
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
            
        .Execute Replace:=wdReplaceAll
    End With
    
    With oFind
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = ""
        .Font.Subscript = True
        
        .Replacement.Text = "<sub>^&</sub>"
        .Replacement.Font.Subscript = False
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
            
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub ConvertPNGsToImgTag(oDoc As Document)
' �������������� ����������� � ����

    If oDoc.InlineShapes.Count = 0 Then
        Exit Sub
    End If
        
    Dim nObj As Integer
    Dim oShp As InlineShape
    
    Dim Width, Height As Single
    For nObj = oDoc.InlineShapes.Count To 1 Step -1
        Set oShp = oDoc.InlineShapes(nObj)
        Width = oShp.Width
        Height = oShp.Height
        
        'Find image with pattern
        oDoc.Range(oShp.Range.Start, oShp.Range.End).Text = "<img " & _
            "src=""" & GetImagePath(oDoc, nObj) & """ " & _
            "width=""" & Single2Str(Width) & "pt"" " & _
            "height=""" & Single2Str(Height) & "pt"" " & _
            "/>"
    Next
End Sub

Private Function Single2Str(ByVal Value As Single)
'�������������� ����� � ������ ��� ����������� �� ��������

    Dim SepDecimal, SepMiles As String
    
    If CStr(0.5) = "0.5" Then
        SepDecimal = "."
        SepMiles = ","
    Else
        SepDecimal = ","
        SepMiles = "."
    End If
    
    Single2Str = CStr(Value)
    Single2Str = Replace(Single2Str, " ", "")
    Single2Str = Replace(Single2Str, SepMiles, "")
    Single2Str = Replace(Single2Str, SepDecimal, ".")
End Function

Public Function GetRelativeDirectoryFromPathFilename(ByVal path As String) As String
    If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    If Right(path, 1) = "/" Then path = Left(path, Len(path) - 1)
    
    Dim pos, posAlt As Integer
    pos = InStrRev(path, "\")
    posAlt = InStrRev(path, "/")
    
    If pos < posAlt Then
        pos = posAlt
    End If
    
    If pos > 0 Then
        GetRelativeDirectoryFromPathFilename = "@@PLUGINFILE@@/" & Right(path, Len(path) - pos)
    Else
        GetRelativeDirectoryFromPathFilename = "@@PLUGINFILE@@"
    End If
End Function

Private Function GetImagePath(oDoc As Document, ByVal ImageNumber As Integer)
'����� �������� ����������� � �������� �����������

    Dim BaseFolder As String
    BaseFolder = oDoc.FullName & ".files\"
        
    GetImagePath = ""
    If Len(Dir(BaseFolder, vbDirectory)) > 0 Then
        Dim nZeros As Integer
        For nZeros = 5 To 0 Step -1
            GetImagePath = "image" & String(nZeros, "0") & CStr(ImageNumber) & ".png"
            If Len(Dir(BaseFolder & GetImagePath)) > 0 Then
                Exit For
            End If
        Next
        
        BaseFolder = GetRelativeDirectoryFromPathFilename(BaseFolder) & "/"
    Else
        BaseFolder = "@@PLUGINFILE@@/"
    End If
    
    If Len(GetImagePath) = 0 Then
        GetImagePath = "image" & CStr(ImageNumber) & ".png"
    End If
    
    GetImagePath = BaseFolder & GetImagePath
    
    GetImagePath = Replace(GetImagePath, "&", "&amp;")
End Function

Private Sub ProcessStyledParagraphsContainedHTML(oDoc As Word.Document)
' ��������� ����������� �������, ���������� HTML

    Dim nParagraph As Integer
    Dim oParagraph As Word.Paragraph
    
    Dim nStartParagraph, nEndParagraph As Integer
    nStartParagraph = -1
    nEndParagraph = -1
    Dim HasHTML As Boolean: HasHTML = False
    
    For nParagraph = 1 To oDoc.Paragraphs.Count
        Set oParagraph = oDoc.Paragraphs(nParagraph)
        
        Select Case GetParagraphType(oParagraph)
            Case "styled_question", "styled_answer"
                If HasHTML And nStartParagraph >= 0 Then
                    DoProcessStyledParagraphsContainedHTML oDoc, nStartParagraph, nEndParagraph
                End If
                nStartParagraph = nParagraph
                nEndParagraph = nParagraph
                HasHTML = HasHTMLContent(oParagraph.Range.Text)
            Case "styled_continue"
                nEndParagraph = nParagraph
                HasHTML = HasHTML Or HasHTMLContent(oParagraph.Range.Text)
            Case Else
                If HasHTML And nStartParagraph >= 0 Then
                    DoProcessStyledParagraphsContainedHTML oDoc, nStartParagraph, nEndParagraph
                End If
                nStartParagraph = -1
                nEndParagraph = -1
                HasHTML = False
        End Select
   
        If (nParagraph = oDoc.Paragraphs.Count) _
        And HasHTML And nStartParagraph >= 0 Then
            nEndParagraph = nParagraph
            DoProcessStyledParagraphsContainedHTML oDoc, nStartParagraph, nEndParagraph
        End If
    Next
End Sub

Private Sub DoProcessStyledParagraphsContainedHTML( _
    oDoc As Word.Document, _
    ByVal StartParagraph As Integer, _
    ByVal EndParagraph As Integer)
' ��������� ��������� �������, ����������� HTML

    Dim nParagraph As Integer
    Dim oParagraph As Word.Paragraph
        
    For nParagraph = StartParagraph To EndParagraph
        Set oParagraph = oDoc.Paragraphs(nParagraph)
        
        If Not HasHTMLParagraphs(oParagraph.Range.Text) Then
            oParagraph.Range.InsertBefore "<p>"
            With oParagraph.Range
                .Collapse wdCollapseEnd
                .Move wdCharacter, -1
                .InsertAfter "</p>"
            End With
        End If
        
        If nParagraph = StartParagraph Then
            oParagraph.Range.InsertBefore "[html]"
        End If
    Next
End Sub

Private Function HasHTMLContent(ByVal Text As String) As Boolean
    Dim re As New RegExp
    With re
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "&[A-Za-z]+;|&#x?[A-Za-z0-9]+;|</?\w+(?:(?:\s+\w+(?:\s*\\?=\s*(?:"".*?""|'.*?'|[\^'"">\s]+))?)+\s*|\s*)/?>"
    End With
        
    Dim m As Object
    Set m = re.Execute(Text)
    
    HasHTMLContent = m.Count = 1
End Function

Private Function HasHTMLParagraphs(ByVal Text As String) As Boolean
    Dim re As New RegExp
    With re
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "</?[Pp](?:(?:\s+\w+(?:\s*\\?=\s*(?:"".*?""|'.*?'|[\^'"">\s]+))?)+\s*|\s*)/?>"
    End With
        
    Dim m As Object
    Set m = re.Execute(Text)
    
    HasHTMLParagraphs = m.Count = 1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ��� 4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub EscapeSpecialChars()
' ������������� ����������� �������� "~", "=", "#", "{", "}", and ":".
'   ����� ����� ��������� ����������� ���� "\", ����� ������� ��� �� �������
'   �� ���������� ���������.
'
' �������� ����������� �� ��������� ����
'   ����� �������� ����������� ���������� ���������, �������� �����������
'   � ����������� ����������� ��������� � ������ HTML.
'
' ������: EscapeSpecialChars
' �����: �.�. ���������
' ����: 2018-03-28

    
    
    ' �������� ���������� ������
    Dim oldScrUpd As Boolean
    oldScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' ����������� ������������� ������
    Application.UndoRecord.StartCustomRecord "EscapeSpecialChars"
    
    On Error GoTo ErrorHandler
    
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    
    oDoc.Activate
    
    DoEscapingSpecialChars oDoc
    
Finish:
    ' ����������� ������������� ������
    Application.UndoRecord.EndCustomRecord
    
    ' ��������������� ���������� ������
    Application.ScreenUpdating = oldScrUpd
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        Dim Msg As String
        Msg = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
    
    'Resume Next
    GoTo Finish
End Sub

Private Sub DoEscapingSpecialChars(oDoc As Word.Document)
' ������ ����������� ��������
    
    Dim oFind As Word.Find
    'Set oFind = oDoc.Content.Find
    
    'Fix the skipped blank Header/Footer problem
    Dim lngJunk As WdStoryType
    lngJunk = oDoc.Sections(1).Headers(1).Range.StoryType
    Set oFind = oDoc.StoryRanges(wdMainTextStory).Find
    
    '������ ����������� ��������
    With oFind
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = "[\\~=\{\}#:]"
        .Replacement.Text = "^92^&"
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
            
        .Execute Replace:=wdReplaceAll
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ��� 5
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ConvertToMoodleGIFT()
' ����������� � ������ Moodle GIFT
'
' ������: ConvertToMoodleGIFT
' �����: �.�. ���������
' ����: 2018-03-28

    
    
    ' �������� ���������� ������
    Dim oldScrUpd As Boolean
    oldScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
        
    On Error GoTo ErrorHandler
    
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    
    oDoc.Activate
    
    SaveToMoodleGIFTFileStyledParagraphs oDoc
    
    'TODO: Open utf-8 text
    'Dim FName As String: FName = oDoc.FullName & ".txt"
    'oDoc.Close
    'Documents.Open FileName:=FName, Format:=wdFormatText, Encoding:=65001
    
Finish:
    ' ��������������� ���������� ������
    Application.ScreenUpdating = oldScrUpd
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        Dim Msg As String
        Msg = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
    
    'Resume Next
    GoTo Finish
End Sub

Private Sub SaveToMoodleGIFTFileStyledParagraphs(oDoc As Word.Document)
' �������������� ����������� ������� � ������ Moodle GIFT � ���������� ����������� � ����

    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    
    On Error GoTo Finish
       
    Dim fso As Object
    Set fso = CreateObject("ADODB.Stream")
    With fso
        .Type = adTypeText
        .Charset = "UTF-8"
        .Open
    End With
    
    Dim PreviousParagraphType, ParagraphType As String
    ParagraphType = "empty"
    PreviousParagraphType = ParagraphType
    
    Dim nParagraph As Integer
    Dim oParagraph As Word.Paragraph
    
    Dim FirstLineQuestion As Variant
        
    Dim ParagraphText As String
    
    oDoc.ConvertNumbersToText
    
    For nParagraph = 1 To oDoc.Paragraphs.Count: Do
    
        Set oParagraph = oDoc.Paragraphs(nParagraph)
        
        ParagraphText = Replace(ParagraphText, Chr(9), " ")
        ParagraphText = Replace(ParagraphText, Chr(10), " ")
        ParagraphText = Replace(ParagraphText, Chr(13), " ")
        ParagraphText = Trim(oParagraph.Range.Text)
                
        ParagraphType = GetParagraphType(oParagraph)

        Select Case ParagraphType
            Case "styled_category"
                If PreviousParagraphType = "styled_answer" Then
                    Exit Do
                End If

                fso.WriteText vbCrLf
                fso.WriteText "$CATEGORY: $module$ / " & ParagraphText & vbCrLf
                fso.WriteText vbCrLf
                
            Case "styled_question"
                If PreviousParagraphType = "styled_answer" Then
                    fso.WriteText "}" & vbCrLf
                End If

                fso.WriteText vbCrLf
                FirstLineQuestion = ProcessStyledQuestionParagraph(oParagraph)
                If Len(FirstLineQuestion(1)) > 0 Then
                    fso.WriteText "::Q" & FirstLineQuestion(1) & vbCrLf
                    fso.WriteText "::" & FirstLineQuestion(2) & vbCrLf
                Else
                    fso.WriteText FirstLineQuestion(2) & vbCrLf
                End If
                
            Case "styled_answer"
                If PreviousParagraphType = "styled_question" Then
                    fso.WriteText "{" & vbCrLf
                End If
                
                If PreviousParagraphType = "styled_question" Or PreviousParagraphType = "styled_answer" Then
                    fso.WriteText ParagraphText ' & vbCrLf
                End If
                
            Case "styled_continue"
                If PreviousParagraphType = "styled_question" Or PreviousParagraphType = "styled_answer" Then
                    fso.WriteText ParagraphText ' & vbCrLf
                End If

                Exit Do
                
            Case Else: Exit Do
        End Select

        PreviousParagraphType = ParagraphType
        Loop While False
        
        If nParagraph = oDoc.Paragraphs.Count Then
            If ParagraphType = "styled_answer" Then
                fso.WriteText "}" & vbCrLf
            End If
                
            fso.WriteText vbCrLf
        End If
    Next
    
    oDoc.Undo ' oDoc.ConvertNumbersToText
    
    fso.SaveToFile oDoc.FullName & ".txt", adSaveCreateOverWrite
    
Finish:
    Set fso = Nothing
End Sub

Private Function ProcessStyledQuestionParagraph(oParagraph As Paragraph) As String()
' ��������� ������ ������� � ��� �����������

    Dim Results(1 To 2) As String
    
    Dim re As New RegExp
    With re
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\s*(\d+)\.\s*(.*?)\s*$"
    End With
        
    Dim m As Object
    Set m = re.Execute(oParagraph.Range.Text)
    
    Dim Processed As Boolean
    Processed = m.Count = 1
    If Processed Then
        Processed = m(0).SubMatches.Count = 2
    End If
    
    If Processed Then
        Results(1) = Trim(m(0).SubMatches(0))
        Results(2) = Trim(m(0).SubMatches(1))
    Else
        Results(1) = ""
        Results(2) = Trim(oParagraph.Range.Text)
    End If
    
    ProcessStyledQuestionParagraph = Results
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ��� 6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ConvertToMoodleGIFTWithMedia()
' ����������� � ������ Moodle GIFT with media
'
' ������: ConvertToMoodleGIFTWithMedia
' �����: �.�. ���������
' ����: 2018-03-28
    On Error GoTo ErrorHandler
    
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    
    oDoc.Activate
    'TODO: �� ����������
    'https://accessexperts.com/blog/2012/02/06/zipandunzipfrommicrosoftvba/
    'https://excelpoweruser.blogspot.ru/2011/07/how-to-zip-file-by-vba.html
    'https://stackoverflow.com/questions/35717193/unzip-file-through-excel-vba-code
    'https://www.rondebruin.nl/win/s7/win001.htm

Finish:
    ' ��������������� ���������� ������
    Application.ScreenUpdating = oldScrUpd
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        Dim Msg As String
        Msg = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    End If
    
    'Resume Next
    GoTo Finish
End Sub
