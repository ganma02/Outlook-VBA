Sub UpdateContacts()
Dim olApp As Object
Dim olNs As Object
Dim cFolder As Object
Dim groupFolder As Object
Dim strGroup As String
Dim myContacts As Object
Dim myItem As Object
Dim strAddress As String
Dim strFix As String
Dim blnCreated As Boolean
Dim i As Integer
Dim FoldersArray As Variant

    Sheets("sheet_name").Select

    On Error Resume Next
   Set olApp = CreateObject("Outlook.Application")
   Set olNs = olApp.GetNamespace("MAPI")
    
            FolderPath = "anas@example.com\Contacts\folder"
   FoldersArray = Split(FolderPath, "\")
    Set cFolder = olApp.Session.Folders.Item(FoldersArray(0))
    If Not cFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Object
            Set SubFolders = cFolder.Folders
            Set cFolder = SubFolders.Item(FoldersArray(i))
            If cFolder Is Nothing Then
                MsgBox "Folder Not Found"
            End If
        Next
    End If
 Debug.Print cFolder
    i = 2
    Set myContacts = cFolder.items
    Do Until Trim(Cells(i, 1).Value) = ""
strAddress = Cells(i, 1)
Set myItem = myContacts.Find("[Email1Address]='" & strAddress & "'")
If TypeName(myItem) = "ContactItem" Then
If Not TypeName(myItem) = "Nothing" Then
myItem.categories = "Delete"
myItem.Save
End If
End If

i = i + 1
Loop

Set olApp = Nothing

End Sub

