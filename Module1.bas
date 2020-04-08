Attribute VB_Name = "Module1"
Sub Macro_Make_Data_Table()
Attribute Macro_Make_Data_Table.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro_Make_Data_Table Macro
'

'
        
    Sheets(1).Activate
    Application.Goto Reference:="R12C1"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$12:$R$7556"), , xlYes).Name _
        = "Table_Data"

End Sub
Sub Macro_Make_Bins_Table()
Attribute Macro_Make_Bins_Table.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro_Make_Bins_Table Macro
'

'
    Sheets(1).Activate
    Sheets.Add After:=ActiveSheet
    ActiveCell.FormulaR1C1 = "Start Duration"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "End Duration"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Bin"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=5/1440"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Less than 5 minutes"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=5/1440"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=10/1440"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "5-10 minutes"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=10/1440"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=15/1440"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "10-15 minutes"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "=15/1440"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=20/1440"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "15-20 minutes"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "=20/1440"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "20+ minutes"
    Range("A1:C6").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$6"), , xlYes).Name = _
        "Table2"
    Range("Table2[#All]").Select
    ActiveSheet.ListObjects("Table2").Name = "Table_Bins"
    ActiveSheet.Name = "Call Length Bins"
End Sub
Sub Macro_Add_Data_Columns()
Attribute Macro_Add_Data_Columns.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro_Add_Data_Columns Macro
'

'
    Sheets(1).Activate
    Range("Table_Data[[#Headers],[Call Start Time]]").Activate
    Selection.End(xlToRight).Select
    Range("Table_Data[[#Headers],[Call Start Time]]").Offset(0, Range("Table_Data").Columns.Count).Value = "Bin"
    Range("Table_Data[[#Headers],[Call Start Time]]").Offset(0, Range("Table_Data").Columns.Count).Value = "Date"
    Range("Table_Data[[#Headers],[Call Start Time]]").Offset(0, Range("Table_Data").Columns.Count).Value = "Day of Week"
    Range("Table_Data[[#Headers],[Call Start Time]]").Offset(0, Range("Table_Data").Columns.Count).Value = "Time"
    Range("Table_Data[[#Headers],[Call Start Time]]").Offset(0, Range("Table_Data").Columns.Count).Value = "During GG Registration"
    Range("Table_Data[[#Headers],[Call Start Time]]").Offset(0, Range("Table_Data").Columns.Count).Value = "Under 1 Minute"
    Range("Table_Data[Bin]").Formula = "=IF(ISBLANK([@[Talk Time]]),"""",VLOOKUP([@[Talk Time]],Table_Bins,3,TRUE))"
    With Range("Table_Data[Date]")
        .Formula = "=DATEVALUE(LEFT(Table_Data[@[Call End Time]],10))"
        .NumberFormat = "m/d/yyyy"
    End With
    Range("Table_Data[Day of Week]").Formula = "=TEXT([@Date],""dddd"")"
    With Range("Table_Data[Time]")
        .Formula = "=TIMEVALUE(RIGHT([@[Call End Time]],11))"
        .NumberFormat = "[$-en-US]h:mm AM/PM;@"
    End With
    Range("Table_Data[During GG Registration]").FormulaR1C1 = _
        "=IFERROR(" & Chr(10) & "IF(OR(" & Chr(10) & "AND(Table_Data[@[Day of Week]]=""Monday"",Table_Data[@Time]>=TIMEVALUE(""13:45""),Table_Data[@Time]<TIMEVALUE(""16:00""))," & Chr(10) & "AND(Table_Data[@[Day of Week]]=""Tuesday"",Table_Data[@Time]>=TIMEVALUE(""08:45""),Table_Data[@Time]<TIMEVALUE(""10:45""))," & Chr(10) & "AND(Table_Data[@[Day of Week]]=""Wednesday"",Table_Data[@Time]>=TIMEVALUE(""13:45""),Table_Data[@Time]<TIMEVALUE(""16:00""))," & Chr(10) & "AND(Table_Data[@[D" & _
        "ay of Week]]=""Thursday"",Table_Data[@Time]>=TIMEVALUE(""10:45""),Table_Data[@Time]<TIMEVALUE(""13:00""))," & Chr(10) & "AND(Table_Data[@[Day of Week]]=""Thursday"",Table_Data[@Time]>=TIMEVALUE(""13:45""),Table_Data[@Time]<TIMEVALUE(""16:00"")))," & Chr(10) & """Yes"",""No"")," & Chr(10) & """"")" & _
        ""
    Range("Table_Data[Under 1 Minute]").Formula = "=IF([@[Talk Time]]<TIMEVALUE(""0:01""),""Yes"",""No"")"
    
End Sub

Sub Macro_Create_Call_Summary()
'
' Macro_Create_Call_Summary Macro
'

'
    Sheets(1).Activate
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "NewReport"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table_Data").CreatePivotTable TableDestination:="NewReport!R7C1" _
        , TableName:="PivotTable3"
    Sheets("NewReport").Select
    Cells(7, 1).Select
    With ActiveSheet.PivotTables("PivotTable3")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Call Result")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields( _
        "During GG Registration")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Under 1 Minute")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Bin"), "Count of Bin", xlCount
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Bin")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Call Result").CurrentPage _
        = "Answered"
    ActiveSheet.PivotTables("PivotTable3").PivotFields("During GG Registration"). _
        CurrentPage = "No"
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Under 1 Minute"). _
        CurrentPage = "No"

    ActiveSheet.Name = "Answered Calls by Length"

End Sub


Sub Generate_Reports()

Application.Run "Macro_Make_Data_Table"
Application.Run "Macro_Make_Bins_Table"
Application.Run "Macro_Add_Data_Columns"
Application.Run "Macro_Create_Call_Summary"

End Sub
