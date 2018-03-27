Function GetEmailAddress(Sin As String) As String
  Dim X As Long, AtSign As Long, AtSign2 As Long, StartAt As Long, S As String, subS As String
  Dim Locale As String, Domain As String
  Locale = "[A-Za-z0-9.!#$%&'*/=?^_`{|}~+-]"
  Domain = "[A-Za-z0-9._-]"
  StartAt = 1
  Do
    S = Mid(Sin, StartAt)
    AtSign = InStr(StartAt, S, "@")
    If AtSign < 2 Then Exit Do
    If Mid(S, AtSign - 1, 1) Like Locale Then
      For X = AtSign To 1 Step -1
        If Not Mid(" " & S, X, 1) Like Locale Then
          subS = Mid(S, X)
          If Left(subS, 1) = "." Then subS = Mid(subS, 2)
          Exit For
        End If
      Next
      AtSign2 = InStr(subS, "@")
      For X = AtSign2 + 1 To Len(subS) + 1
        If Not Mid(subS & " ", X, 1) Like Domain Then
          subS = Left(subS, X - 1)
          If Right(subS, 1) = "." Then subS = Left(subS, Len(subS) - 1)
          GetEmailAddress = GetEmailAddress & ", " & subS
          Exit For
        End If
      Next
    End If
    StartAt = AtSign + 1
  Loop
  GetEmailAddress = Mid(GetEmailAddress, 3)
End Function
