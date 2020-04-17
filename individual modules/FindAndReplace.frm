VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindAndReplace 
   Caption         =   "Find and Replace"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FindAndReplace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindAndReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub FileTypeListBox_Click()

End Sub

Private Sub UserForm_Initialize()
    FileTypesListBox.AddItem ("docx")
End Sub

Private Sub NewFileType_Click()
     FileTypesListBox.AddItem (InputBox("What is the file extension of the new file type you want to add?" & vbNewLine & "You should make sure the file type is supported by Microsoft Word", "Add new file type"))
End Sub

Sub legacyFindAndReplaceInFolder(folderPath As String, findText As String, replaceText As String, fileType As String) 'legacy function (no textboxes)
  Dim objDoc As Document
  Dim strFile As String
  Dim strFolder As String
  Dim strFindText As String
  Dim strReplaceText As String
 
  '  Pop up input boxes for user to enter folder path, the finding and replacing texts.
  
  strFile = Dir(folderPath & "\" & "*." & fileType, vbNormal)
  strFindText = findText
  strReplaceText = replaceText
 
  '  Open each file in the folder to search and replace texts. Save and close the file after the action.
  While strFile <> ""
  CurrentDocLabel.Caption = "Currently editing word document: " & strFile
    Set objDoc = Documents.Open(FileName:=folderPath & "\" & strFile)
    With objDoc
      With Selection
        .HomeKey Unit:=wdStory
        With Selection.Find
          .Text = findText
          .Replacement.Text = replaceText
          .Forward = ForwardCheckBox.Value
          .Wrap = wdFindContinue
          .Format = FormatCheckBox.Value
          .MatchCase = CaseSensitiveCheckBox.Value
          .MatchWholeWord = WholeWorldCheckBox.Value
          .MatchWildcards = WildcardsCheckBox.Value
          .MatchSoundsLike = SoundsLikeCheckBox.Value
          .MatchAllWordForms = AllWordFormsCheckBox.Value
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
      End With
      objDoc.Save
      objDoc.Close
      strFile = Dir()
      
    End With
  Wend
  CurrentDocLabel.Caption = ""
  MsgBox ("Legacy replaced all instances of '" & CStr(FindBox.Text) & "' with '" & CStr(ReplaceBox.Text) & "' in file type " & fileType)
End Sub

'https://wordmvp.com/FAQs/MacrosVBA/FindReplaceAllWithVBA.htm
Sub FindAndReplaceFirstStoryOfEachType(folderPath As String, findText As String, replaceText As String, fileType As String)
  Dim objDoc As Document
    
  strFile = Dir(folderPath & "\" & "*." & fileType, vbNormal)
 
  '  Open each file in the folder to search and replace texts. Save and close the file after the action.
  ActiveDocument.BuiltInDocumentProperties("Author") = "Alexander Dobson-Pleming"
  While strFile <> ""
  CurrentDocLabel.Caption = "Currently editing word document: " & strFile
    Set objDoc = Documents.Open(FileName:=folderPath & "\" & strFile)
    With objDoc
    
        Dim myStoryRange As Range
    
        For Each myStoryRange In objDoc.StoryRanges
        With myStoryRange.Find
            .Text = findText
            .Replacement.Text = replaceText
            .Forward = ForwardCheckBox.Value
            .Wrap = wdFindContinue
            .Format = FormatCheckBox.Value
            .MatchCase = CaseSensitiveCheckBox.Value
            .MatchWholeWord = WholeWorldCheckBox.Value
            .MatchWildcards = WildcardsCheckBox.Value
            .MatchSoundsLike = SoundsLikeCheckBox.Value
            .MatchAllWordForms = AllWordFormsCheckBox.Value
            .Execute Replace:=wdReplaceAll
        End With
        Do While Not (myStoryRange.NextStoryRange Is Nothing)
            Set myStoryRange = myStoryRange.NextStoryRange
            With myStoryRange.Find
                .Text = findText
                .Replacement.Text = replaceText
                .Forward = ForwardCheckBox.Value
                .Wrap = wdFindContinue
                .Format = FormatCheckBox.Value
                .MatchCase = CaseSensitiveCheckBox.Value
                .MatchWholeWord = WholeWorldCheckBox.Value
                .MatchWildcards = WildcardsCheckBox.Value
                .MatchSoundsLike = SoundsLikeCheckBox.Value
                .MatchAllWordForms = AllWordFormsCheckBox.Value
                .Execute Replace:=wdReplaceAll
            End With
        Loop
    Next myStoryRange
    
    objDoc.Save
    objDoc.Close
    strFile = Dir()
    End With
    
  Wend
  CurrentDocLabel.Caption = ""
  MsgBox ("Replaced all instances of '" & CStr(FindBox.Text) & "' with '" & CStr(ReplaceBox.Text) & "' in file type " & fileType)
  
End Sub



Private Sub OpenButton_Click()
    FolderBox.Text = CommonObjects.GetFolder()
End Sub

Private Sub ReplaceBox_Change()

End Sub

Private Sub StartButton_Click()
    For i = 0 To FileTypesListBox.ListCount - 1
        If LegacyFindReplace.Value = True Then
            Call legacyFindAndReplaceInFolder(CStr(FolderBox.Text), CStr(FindBox.Text), CStr(ReplaceBox.Text), CStr(FileTypesListBox(i)))
        Else
            Call FindAndReplaceFirstStoryOfEachType(CStr(FolderBox.Text), CStr(FindBox.Text), CStr(ReplaceBox.Text), CStr(FileTypesListBox.List(i)))
        End If
    Next
    If FileTypesListBox.ListCount > 1 Then
        MsgBox ("Finished replacements for all file types")
    End If
    Me.hide
End Sub

Private Sub UserForm_Click()

End Sub
