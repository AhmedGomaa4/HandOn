Option Explicit
Sub Post_CitTo_SC()

'Its Temp Macro for the month 13
    
    Dim NewNamesArr As Object
    Dim xdialog, xworkbook As Object
    Dim MyCIT, MyUpdate, myDate, dadFdr, newDR, UC, wbName, Lcol As String
    Dim oneWorkBook, item As Variant
    Dim FirstEntry, LR, MaxMonth, Lines, NArows, SCLR As Long
    
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    Dim shName As String, myrange As String
    Set NewNamesArr = CreateObject("System.Collections.ArrayList")
    MyCIT = ActiveWorkbook.Name
    myrange = Range("A3:S" & ActiveSheet.UsedRange.Rows.Count).Address
    
    Set xdialog = Application.FileDialog(msoFileDialogFilePicker)
    xdialog.AllowMultiSelect = True
    If xdialog.Show = 0 Then Exit Sub
    myDate = Format(Date - 1, "DD-MM-YYYY")
    MyUpdate = InputBox("enter till date", , myDate)
    
'' Creating New Folder
    ChDir Mid(xdialog.SelectedItems(1), 1, InStrRev(xdialog.SelectedItems(1), "\") - 1)
    ChDir ".."
    dadFdr = CurDir
    MkDir dadFdr & "\SC " & MyUpdate
    newDR = dadFdr & "\SC " & MyUpdate


    ' Choosing stock control files
For Each oneWorkBook In xdialog.SelectedItems
    Set xworkbook = Workbooks.Open(oneWorkBook, UpdateLinks:=0)
    
        Sheets("Purchase").Select
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
        
        If InStr(LCase(oneWorkBook), "com") > 0 Then shName = "com"
        If InStr(LCase(oneWorkBook), "bmob") > 0 Then shName = "bmob"
        If InStr(LCase(oneWorkBook), "dom") > 0 Then shName = "dom"
        If InStr(LCase(oneWorkBook), "marketp") > 0 Then shName = "marketp"
        If InStr(LCase(oneWorkBook), "outlet") > 0 Then shName = "Outlet"
        
        'chck if it's commercial WB or not
        If shName = "com" Then
            'Deleting old purchase data from Commercial WB
            If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
            If Sheets("Stock control").FilterMode = True Then Sheets("Stock control").ShowAllData
            Range("S2:S" & ActiveSheet.UsedRange.Rows.Count).ClearContents
            Cells(Rows.Count, 1).End(xlUp).CurrentRegion.EntireRow.Delete
            Cells(Rows.Count, 1).End(xlUp).Offset(2).Select
        Else
            'Deleting old purchase data from Dom&Bmob WB
            If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
            If Sheets("Stock control").FilterMode = True Then Sheets("Stock control").ShowAllData
            Range("S2:S" & ActiveSheet.UsedRange.Rows.Count).ClearContents
            Cells(1, 1).CurrentRegion.Offset(2).EntireRow.Delete
            Cells(3, 1).Select
        End If
    'filter and select data from CIT
    Windows(MyCIT).Activate
    
    ActiveSheet.Range(myrange).AutoFilter field:=19, Criteria1:="" & shName & "*", Operator:=xlOr
    If Range("3:3").EntireRow.Hidden = False Then
        Range("A3").Select
    Else
        'ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(2).Cells(1, 1).Select
        FirstEntry = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(FirstEntry, "A").Select
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Resize(, 18).Copy
    'paste data into purchase sheet
    Windows(xworkbook.Name).Activate
    ActiveSheet.Paste
    LR = Sheets("Purchase").Cells(Rows.Count, 1).End(xlUp).Row
'    MaxMonth = (Format(WorksheetFunction.Max(Range("M:M")), "M") - 1) * 13 + 18

    MaxMonth = (34 - 1) * 13 + 18       '(34 Means Oct
    
    Selection.Resize(, 1).Offset(, 18).Select
    Selection.FormulaR1C1 = "=VLOOKUP(RC[-18],'Stock control'!C[-18],1,0)"
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    
        ActiveSheet.Range("A2:S" & LR).AutoFilter field:=19, Criteria1:="=#N/A" _
        , Operator:=xlAnd
    
    Lines = Application.WorksheetFunction.CountA(Range("B:B").SpecialCells(xlCellTypeVisible))
   
   If Lines = 1 Then GoTo noNA '''''''''''''no na
        
        If Lines > 2 Then ' there is mor than 1 na so fill down
        'ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(2)(1, 1).Resize(, 9).Select
        FirstEntry = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(FirstEntry, "A").Resize(, 9).Select
        Range(Selection, Selection.End(xlDown)).Copy Cells(LR + 10, 1)
        Else                ' there is Only One NA line
        'ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(2)(1, 1).Resize(, 9).Select
        FirstEntry = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(FirstEntry, "A").Resize(, 9).Select
        Selection.Copy Cells(LR + 10, 1)
        End If
                
    ActiveSheet.Range("A" & LR + 10).CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3, _
        4, 5, 6, 7, 8, 9), Header:=xlNo
    
    Range("A" & LR + 10).CurrentRegion.Copy
    NArows = Range("A" & LR + 10).CurrentRegion.Rows.Count
    Sheets("Stock control").Select
    
    Range("A4").End(xlDown).Offset(1).Select
    ActiveSheet.Paste
    ActiveCell.Offset(-1).EntireRow.Copy
    ActiveCell.Resize(NArows).EntireRow.PasteSpecial xlPasteFormats
    Sheets("Purchase").Cells(LR + 10, 1).CurrentRegion.EntireRow.Delete
    
noNA:
    
    Sheets("stock control").Select
    Sheets("Purchase").ShowAllData
    Sheets("Purchase").Range("A1").Value = MyUpdate
    SCLR = Sheets("stock control").Cells(Rows.Count, 1).End(xlUp).Row
      
    'hide and show grouping & Paste Formulas
    
    Lcol = Split(Range("A4").End(xlToRight).Address(1, 0), "$")(0)
        Lcol = Split(Range("A4").End(xlToRight).Address(1, 0), "$")(0)
    Sheets("stock control").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    On Error Resume Next: Sheets("stock control").ShowAllData: On Error GoTo 0

'custom version copy formulas starting from sep 2022
    Range("JM2:" & Lcol & "2").Copy Range("JM5:" & Lcol & SCLR)
    Range("JM5:" & Lcol & SCLR).Copy
    Range("JM5:" & Lcol & SCLR).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    

    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    

    Cells(1, MaxMonth).Offset(, -5).EntireColumn.Select
    Range(Selection, Selection.Offset(, 11)).Select
    Selection.EntireColumn.Hidden = False
    ActiveCell.Offset(3, 5).Select
    
    UC = ActiveCell.Offset(1).Address
    wbName = Split(ActiveWorkbook.Name, "ill")(0) & "ill " & MyUpdate & ".xlsb"
    ActiveWorkbook.SaveAs newDR & "\" & wbName, FileFormat:=xlExcel12
    NewNamesArr.Add wbName
Next

'''''''''''''''''' Done Making Each SC WB ''''''''''''''''''''''''''


'''''''''''''''''' Now Creating Uni SC ''''''''''''''''''''''''''

    Dim UniSC As Workbook
    Set UniSC = Workbooks.Add
    UniSC.Activate
    [A3].Value = "ITem Code": [B3].Value = "ITem Desc": [C3].Value = "Cost"
    [B1].Value = "S.C Av Cost " & MyUpdate
    Range("A:B").ColumnWidth = 28
    
'    For i = 1 To xdialog.SelectedItems.Count
'       mYbook = Mid(xdialog.SelectedItems(i), InStrRev(xdialog.SelectedItems(i), "\") + 1, 100)
For Each item In NewNamesArr

        Workbooks(item).Activate
        Range(Range("A5"), Range("B5").End(xlDown)).Copy
        UniSC.Activate
        Cells(Rows.Count, 1).End(xlUp).Offset(1).Select
        ActiveSheet.Paste
        Workbooks(item).Activate
        Range(Range(UC), Range(UC).End(xlDown)).Copy
        UniSC.Activate
        Cells(Rows.Count, 3).End(xlUp).Offset(1).Select
        ActiveSheet.Paste
        
    Next
    [A1].Select
    ActiveWorkbook.SaveAs newDR & "\" & [B1].Value & ".xlsb", FileFormat:=xlExcel12
'MsgBox "Done" & vbNewLine & "Data PosTed To Stock Control Successfully :)" & vbNewLine & _
        "Please Check Then SAVE"
  SecondsElapsed = Round(Timer - StartTime, 2)
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

' Free Up Memory
    Set xdialog = Nothing
    Set xworkbook = Nothing
    oneWorkBook = Empty
    item = Empty
    MyCIT = Empty
    MyUpdate = Empty
    myDate = Empty
    dadFdr = Empty
    newDR = Empty
    UC = Empty
    wbName = Empty
    Lcol = Empty
    FirstEntry = Empty
    LR = Empty
    MaxMonth = Empty
    Lines = Empty
    NArows = Empty
    SCLR = Empty
    StartTime = Empty
    SecondsElapsed = Empty
    shName = Empty
    myrange = Empty

End Sub





