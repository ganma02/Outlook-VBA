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

    For Each obj In Selection
    With obj
    .UserProperties("anything") = True
    .Save
    End With

    Next

    Set Session = Nothing
    Set currentExplorer = Nothing
    Set obj = Nothing
    Set Selection = Nothing

End Sub

