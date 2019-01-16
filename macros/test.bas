Sub AutoTextCongress()

Dim oAutoText As AutoTextEntry

Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries.Add(Name:="shorter", Range:=Selection.Range)
    oAutoText.Value = "CMON TRY AGAIN"

Set oAutoText = Nothing

End Sub