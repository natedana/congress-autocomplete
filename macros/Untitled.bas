Attribute VB_Name = "Module3"
' Attribute VB_Name = "AutoCongress"

Sub AutoCongress()
' Attribute AutoCongress.VB_Description = "Auto Complete for 2018 Feb. House and Senate"
' Attribute AutoCongress.VB_ProcData.VB_Invoke_Func = "Nates_Templates.NewMacros.AutoCongress"

Call TestMacro

End Sub

Private Sub TestMacro()

Dim oAutoText As AutoTextEntry

Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Alexander, Lamar", Range:=Selection.Range)
    oAutoText.Value = "Senator Lamar Alexander (R-TN)"

Set oAutoText = Nothing

End Sub
