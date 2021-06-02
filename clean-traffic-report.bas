Attribute VB_Name = "Module1"
Sub EveryThing_ASPNET()
    Call DeleteTop2RowsIf
    Call HideColumns
    Call RenameTitleText
    Call PinTopRow
    Call RenameMoreColms
    Call WidenColum
    Call MakeHyperLinkColumn
    Call AllAspNetCore_ASPNET
End Sub

Sub AllAspNetCore_ASPNET()
    Call RemoveText_ASPNET
    Call ShortenLongColumns_ASPNET
End Sub

Sub MakeHyperLinkColumn()
    Call MakeHyperLinks
    Call HyperLinkColumnName
End Sub

Sub HyperLinkColumnName()
    Range("Table1[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "Link"
    Range("AP2").Select
End Sub
    
Sub MakeHyperLinks()
    Range("AP2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=HYPERLINK([@LiveUrl])"
    Range("AP3").Select
End Sub

Sub DeleteHide()
    Call DeleteTop2RowsIf
    Call HideColumns
End Sub

Sub DeleteTop2RowsIf()

    Set rRng = Sheet1.Range("A2")
    ' If row 2 is empty, remove the top two rows.
    ' Row 1 shows the selected filters, row 2 is blank.
    If IsEmpty(rRng.Value) Then
       Rows("1:1").Select
       Selection.Delete Shift:=xlUp
       Selection.Delete Shift:=xlUp
    End If
    
End Sub

Sub HideColumns()
Columns("A").Hidden = True ' Topic type
Columns("C").Hidden = True ' Live URL
Columns("G").Hidden = True ' Search referrals
Columns("H:K").Hidden = True 'KPI rank, KPI rank change, CTR, CopyTryScroll
Columns("N:W").Hidden = True ' Organic search through Dwell rate
Columns("Y:AO").Hidden = True ' CSAT response rate through end
End Sub

Sub RemoveText_ASPNET()
' Remove "in ASP.NET Core"
    Range("Table1[[#Headers],[Title]]").Select
    Cells.Replace What:=" in ASP.NET Core", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub RenameTitleText()
' Remove "Sum of" from heading cells.
    Rows("1:1").Select
    Range("Table1[[#Headers],[Title]]").Activate
    Selection.Replace What:="Sum of ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub RenameMoreColms()
' Shorten long column header text
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

Sub ShortenLongColumns_ASPNET()
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
     Columns("B:B").ColumnWidth = 50 ' Title
     Columns("D:F").EntireColumn.AutoFit ' Page views, PV MoM, Visitors
     Columns("L:M").EntireColumn.AutoFit ' Bounce rate, Exit rate
     Columns("X:X").EntireColumn.AutoFit ' CSAT rate
End Sub
