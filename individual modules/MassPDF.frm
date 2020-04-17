VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MassPDF 
   Caption         =   "Mass PDF Conversion"
   ClientHeight    =   9195.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "MassPDF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MassPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddDocButton_Click()
    DocListBox.AddItem (CommonObjects.GetFile)
End Sub

Private Sub AddFileType_Click()
    Dim fileTypeToAdd As String
    fileTypeToAdd = InputBox("File types should be compatible with Microsoft Word", "Add File Type")
    For i = 0 To FileTypesListBox.ListCount - 1
        If fileTypeToAdd = FileTypesListBox.List(i) Then
            MsgBox ("File type is already in directory")
            Exit Sub
        End If
    Next
    FileTypesListBox.AddItem (fileTypeToAdd)
    
    Call PopulateDocList
End Sub




Private Sub OpenButton2_Click()
    SavePathBox.Text = CommonObjects.GetFolder()
End Sub

Private Sub RemoveDocButton_Click()
    For i = 0 To DocListBox.ListCount - 1
        If DocListBox.Selected(i) = True Then
            DocListBox.RemoveItem (i)
        End If
    Next
End Sub

Private Sub RemoveFileTypeButton_Click()
    For i = 0 To FileTypesListBox.ListCount - 1
        If FileTypesListBox.Selected(i) = True Then
            FileTypesListBox.RemoveItem (i)
        End If
    Next
    
    Call PopulateDocList
End Sub

Private Sub FolderBox_Change()
    Call PopulateDocList
    
End Sub

Private Sub OpenButton_Click()
    FolderBox.Text = CommonObjects.GetFolder()
End Sub

Sub ClearDocList()
    DocListBox.Clear
    'DocListBox.AddItem
    'DocListBox.List(0) = "Document Name"
    'DocListBox.List(0, 1) = "File Type"
End Sub

Sub TestFileIteration()
    Dim i
    i = 0
    'FileTypesListBox.List(i)
    strFile = Dir("C:\Users\Alexander\Documents\Work\VBAHost\Test Documents" & "\" & "*.docx", vbNormal)
    While strFile <> ""
        strFile = Dir()
        MsgBox (strFile)
    Wend
End Sub


Sub PopulateDocList()
    Call ClearDocList
    For i = 0 To FileTypesListBox.ListCount - 1
        strFile = Dir(FolderBox.Value & "\" & "*." & FileTypesListBox.List(i), vbNormal)
        While strFile <> ""
            
            If CommonObjects.ListBoxContains(DocListBox, strFile) = False Then
                DocListBox.AddItem
                DocListBox.List(DocListBox.ListCount - 1) = strFile
                'DocListBox.List(DocListBox.ListCount - 1, 1) = FileTypesListBox.List(i)
            End If
            strFile = Dir()
        Wend
    Next
End Sub



Private Sub StartExportButton_Click()
    If FolderBox.Text = "" Or SavePathBox.Text = "" Then
        MsgBox ("One of the the input folder or export folder boxes are empty" & vbNewLine & "Put something in one of the boxes and try again")
        Exit Sub
    End If

    For i = 0 To DocListBox.ListCount - 1
        Call CommonObjects.WordToPDF(FolderBox.Text, DocListBox.List(i), SavePathBox.Text)
    Next
    
    MsgBox ("Converted " & DocListBox.ListCount & " files to PDF" & vbNewLine & vbNewLine & "An explorer window will now open with your files")
    Me.hide
    Shell "C:\WINDOWS\explorer.exe """ & SavePathBox.Text & "", vbNormalFocus
End Sub

Private Sub UserForm_Initialize()
    Call ClearDocList   'add headers to DocListBox
    
    FileTypesListBox.AddItem ("docx")
    FileTypesListBox.AddItem ("doc")
    FileTypesListBox.AddItem ("rtf")
    FileTypesListBox.AddItem ("txt")
    
    'Call CommonObjects.WordToPDF
End Sub


