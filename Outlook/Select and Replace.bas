Option Explicit

Public Sub DoSomethingSelection()
    Dim Session As Outlook.NameSpace
    Dim currentExplorer As Explorer
    Dim Selection As Selection
    Dim strOld As String
    Dim strNew As String
    
    
    Dim obj As Object

    Set currentExplorer = Application.ActiveExplorer
    Set Selection = currentExplorer.Selection

    strOld = InputBox("What word(s) do you want to replace?")
    strNew = InputBox("What is the replacement word or phrase?")
    For Each obj In Selection
    With obj
    '.BusinessAddressCountry = Replace(.BusinessAddressCountry, strOld, strNew)
    '.FullName = Replace(.FullName, strOld, strNew)
    .UserProperties("genre") = Replace(.UserProperties("genre"), strOld, strNew)
    .Save
    End With

    Next

    Set Session = Nothing
    Set currentExplorer = Nothing
    Set obj = Nothing
    Set Selection = Nothing

End Sub
