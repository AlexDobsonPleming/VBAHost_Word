Attribute VB_Name = "CommonObjects"
Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = Left$(ActiveDocument.Path, InStrRev(Path, "\"))
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function GetFile() As String
    Dim fileDiag As FileDialog
    Dim sItem As String
    Set fileDiag = Application.FileDialog(msoFileDialogOpen)
    With fileDiag
        .Title = "Select a file"
        .AllowMultiSelect = False
        .InitialFileName = ActiveDocument.Path 'Left$(ActiveDocument.Path, InStrRev(Path, "\"))
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFile = sItem
    Set fldr = Nothing
End Function

Function ListBoxContains(listBoxName, searchItem) As Boolean
    For i = 0 To listBoxName.ListCount - 1
        'MsgBox (searchItem & vbNewLine & listBoxName.List(i))
        If InStr(listBoxName.List(i), searchItem) = True Or searchItem = listBoxName.List(i) Then
            ListBoxContains = True
            Exit Function
        End If
    Next
    ListBoxContains = False
End Function

'an excerpt from https://www.thespreadsheetguru.com/the-code-vault/microsoft-word-vba-to-save-document-as-a-pdf-in-same-folder
Sub WordToPDF(folderPath As String, FileName As String, exportPath As String)
    
    
    Dim objDoc As Document
    Set objDoc = Documents.Open(FileName:=folderPath & "\" & FileName)
    On Error GoTo ProblemSaving
        objDoc.ExportAsFixedFormat _
        OutputFileName:=exportPath & "\" & FileName & ".pdf", _
        ExportFormat:=wdExportFormatPDF
    On Error GoTo 0
    
    objDoc.Save
    objDoc.Close
    
    Exit Sub
    
    'Error Handlers
    
ProblemSaving:
    MsgBox "There was a problem saving your PDF. This is most commonly caused" & _
   " by the original PDF file already being open."
    Exit Sub
  
End Sub






Sub Word_ExportPDF()
'PURPOSE: Generate A PDF Document From Current Word Document
'NOTES: PDF Will Be Saved To Same Folder As Word Document File
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim CurrentFolder As String
'Dim fileName As String
Dim myPath As String
Dim UniqueName As Boolean

UniqueName = False

'Store Information About Word File
  myPath = ActiveDocument.FullName
  CurrentFolder = ActiveDocument.Path & "\"
  FileName = Mid(myPath, InStrRev(myPath, "\") + 1, _
   InStrRev(myPath, ".") - InStrRev(myPath, "\") - 1)

'Does File Already Exist?
  Do While UniqueName = False
    DirFile = CurrentFolder & FileName & ".pdf"
    If Len(Dir(DirFile)) <> 0 Then
      UserAnswer = MsgBox("File Already Exists! Click " & _
       "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
      
      If UserAnswer = vbYes Then
        UniqueName = True
      ElseIf UserAnswer = vbNo Then
        Do
          'Retrieve New File Name
            FileName = InputBox("Provide New File Name " & _
             "(will ask again if you provide an invalid file name)", _
             "Enter File Name", FileName)
          
          'Exit if User Wants To
            If FileName = "False" Or FileName = "" Then Exit Sub
        Loop While ValidFileName(FileName) = False
      Else
        Exit Sub 'Cancel
      End If
    Else
      UniqueName = True
    End If
  Loop
  
'Save As PDF Document
  On Error GoTo ProblemSaving
    ActiveDocument.ExportAsFixedFormat _
     OutputFileName:=CurrentFolder & FileName & ".pdf", _
     ExportFormat:=wdExportFormatPDF
  On Error GoTo 0

'Confirm Save To User
  With ActiveDocument
    FolderName = Mid(.Path, InStrRev(.Path, "\") + 1, Len(.Path) - InStrRev(.Path, "\"))
  End With
  
  MsgBox "PDF Saved in the Folder: " & FolderName

Exit Sub

'Error Handlers
ProblemSaving:
  MsgBox "There was a problem saving your PDF. This is most commonly caused" & _
   " by the original PDF file already being open."
  Exit Sub

End Sub


Function ValidFileName(FileName As String) As Boolean
'PURPOSE: Determine If A Given Word Document File Name Is Valid
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim TempPath As String
Dim doc As Document

'Determine Folder Where Temporary Files Are Stored
  TempPath = Environ("TEMP")

'Create a Temporary XLS file (XLS in case there are macros)
  On Error GoTo InvalidFileName
    Set doc = ActiveDocument.SaveAs2(ActiveDocument.TempPath & _
     "\" & FileName & ".doc", wdFormatDocument)
  On Error Resume Next

'Delete Temp File
  Kill doc.FullName

'File Name is Valid
  ValidFileName = True

Exit Function

'ERROR HANDLERS
InvalidFileName:
'File Name is Invalid
  ValidFileName = False

End Function
