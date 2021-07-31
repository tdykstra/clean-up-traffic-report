Sub Everything_ASPNET()
    Call EverythingCommonToAll
    Call RemoveText_ASPNET
    Call ShortenLongColumns_ASPNET
    Call HideColumns_ASPNET
End Sub

Sub EveryThing_DOTNET()
    Call EverythingCommonToAll
    Call RemoveText_DOTNET
    Call ShortenLongColumns_DOTNET
    Call HideColumns_DOTNET
End Sub

Sub Everything_EF()
    Call EverythingCommonToAll
    Call RemoveText_EF
    Call HideColumns_ASPNET
End Sub

Sub EverythingCommonToAll()
    Call DeleteTop2RowsIf
    Call PinTopRow
    Call RemoveSumOf
    Call ShortenLongColumnNames
    Call WidenColumns
    Call MakeHyperLinkColumn
    Call ScrollToBeginning
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

Sub RemoveSumOf()
    Rows("1:1").Select
    Range("Table1[[#Headers],[Title]]").Activate
    Selection.Replace What:="Sum of ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub ShortenLongColumnNames()
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

Sub WidenColumns()
     Columns("B:B").ColumnWidth = 50 ' Title
     Columns("D:F").EntireColumn.AutoFit ' Page views, PV MoM, Visitors
     Columns("L:M").EntireColumn.AutoFit ' Bounce rate, Exit rate
     Columns("X:X").EntireColumn.AutoFit ' CSAT rate
End Sub

Sub ScrollToBeginning()
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollRow = 1
End Sub

Sub HideColumns_ASPNET()
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

Sub RemoveText_EF()
' Remove " - EF Core"
    Range("Table1[[#Headers],[Title]]").Select
    Cells.Replace What:=" - EF Core", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub ShortenLongColumns_ASPNET()
    Range("Table1[[#Headers],[Title]]").Select
    Cells.Replace What:="Secure an ASP.NET Core", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Sub HideColumns_DOTNET()
Columns("A").Hidden = True ' Topic type
Columns("C").Hidden = True ' Live URL
Columns("E").Hidden = True ' PV MoM
Columns("G").Hidden = True ' Search referrals
Columns("N").Hidden = True ' Organic search
Columns("N:W").Hidden = True ' Organic search through Dwell rate
Columns("Y:Z").Hidden = True ' CSAT response rate, CSAT helpful responses
Columns("AB:AO").Hidden = True ' CSAT rating verbatims through end
End Sub

Sub RemoveText_DOTNET()
' Nothing to remove at this time.
End Sub

Sub ShortenLongColumns_DOTNET()
' Nothing to shorten at this time.
End Sub

Sub MakeHyperLinkColumn()
    Call MakeHyperLinks
    Call HyperLinkColumnName
End Sub

Sub MakeHyperLinks()
    Range("AP2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=HYPERLINK([@LiveUrl])"
    Range("AP3").Select
End Sub

Sub HyperLinkColumnName()
    Range("Table1[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "Link"
    Range("AP2").Select
End Sub
