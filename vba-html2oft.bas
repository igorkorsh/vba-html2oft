Public xlApp As New Excel.Application
Public fs As New Scripting.FileSystemObject
Public fd As Office.FileDialog
Public fileName As String
Public saveFolder As String

Sub CreateEmailFromFile()

    Dim filePath As Variant
    Dim fileContent As String

    xlApp.Visible = False
    Set fd = xlApp.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Convert"
        .Filters.Clear
        .Filters.Add "HTML", "*.html; *.htm"
        .Title = "Select a file"

         If .Show = -1 Then
            For Each filePath In .SelectedItems
                strFileContent = fReadFile(filePath)
                Call fCreateEmail(strFileContent)
                fileName = fs.GetBaseName(filePath)
            Next
         End If
    End With

    Set fd = Nothing
    xlApp.Quit
    Set xlApp = Nothing

End Sub

Sub SaveEmail()

    Dim thisMail As Outlook.MailItem
    Dim objItem As Object
    
    Call fGetFolder
    
    Set objItem = ActiveInspector.CurrentItem
    Set thisMail = objItem
    'Debug.Print saveFolder & "\" & fileName
    If saveFolder <> "" Then
        thisMail.SaveAs saveFolder & "\" & fileName & ".oft", OlSaveAsType.olTemplate
        thisMail.Close olDiscard
    Else 
        Exit Sub
    End If
    
    filename = ""
    Set thisMail = Nothing

End Sub

Function fReadFile(file) As String

    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Open
        .Charset = "UTF-8"
        .LoadFromFile file
            fReadFile = .ReadText
        .Close
    End With

End Function

Function fCreateEmail(content)

    Dim newMail As Outlook.MailItem

    Set newMail = Application.CreateItem(olMailItem)
    With newMail
       .BodyFormat = olFormatHTML
       .HTMLBody = content
       .Display
    End With

End Function

Function fGetFolder() As String

    xlApp.Visible = False
    Set fd = xlApp.FileDialog(msoFileDialogFolderPicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .Title = "Select a Folder"
        .InitialFileName = saveFolder

         If .Show = -1 Then
            saveFolder = .SelectedItems(1)
         Else
            saveFolder = ""
            Exit Function
         End If
    End With

    Set fd = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    
End Function
