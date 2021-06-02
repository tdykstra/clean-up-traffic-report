Attribute VB_Name = "Module1"
Sub EveryThing()
    Call DeleteTop2RowsIf
    Call HideColumns
    Call RemoveText
    Call DeleteTop2RowsIf
    Call RenameTitleText
    Call DeleteTop2RowsIf
    Call ShortenLongColumns
    Call PinTopRow
    Call RenameMoreColms
    Call WidenColum
    Call WidenColum
    'Call Macro999
End Sub

Sub DeleteHide()
    Call DeleteTop2RowsIf
    Call HideColumns
End Sub

Sub DeleteTop2RowsIf()

Set rRng = Sheet1.Range("A2")
If IsEmpty(rRng.Value) Then
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
End If
    
End Sub

Sub HideColumns()
Columns("A").Hidden = True ' Topic type
'Columns("C").Hidden = True
Columns("H:K").Hidden = True
Columns("G").Hidden = True
Columns("N:W").Hidden = True
Columns("Y:AO").Hidden = True
End Sub

Sub RemoveText()
'
    Range("Table1[[#Headers],[Title]]").Select
    Cells.Replace What:=" in ASP.NET Core", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub RenameTitleText()
    Rows("1:1").Select
    Range("Table1[[#Headers],[Title]]").Activate
    Selection.Replace What:="Sum of ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub RenameMoreColms()

    Rows("1:1").Select
    Range("Table1[[#Headers],[Title]]").Activate
    Selection.Replace What:="BounceRate", Replacement:="Bounce", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
       Rows("1:1").Select
    Range("Table1[[#Headers],[Title]]").Activate
    Selection.Replace What:="CSATHelpfulRate", Replacement:="CSAT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub ShortenLongColumns() ' ASP.NET Core specific
    Range("Table1[[#Headers],[Title]]").Select
    Cells.Replace What:="Secure an ASP.NET Core", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub PinTopRow()

Set rRng = Sheet1.Range("A2")
 
If IsEmpty(rRng.Value) Then
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
End If

ActiveWindow.SplitRow = 1
ActiveWindow.FreezePanes = True

End Sub


Sub WidenColum()
     Columns("D:F").EntireColumn.AutoFit
     Columns("L:M").EntireColumn.AutoFit
     Columns("X:X").EntireColumn.AutoFit
      Columns("B:B").ColumnWidth = 50
End Sub



