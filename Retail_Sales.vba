Private myrange As String, LR As Long 'To Bypass variables to another sub

Sub Ret_SalesV3()

Dim xUser As String

    xUser = InputBox("Enter your First Letter of your name", , "m")
    Application.ScreenUpdating = False
    Cells.WrapText = False
    ActiveSheet.AutoFilterMode = False
    
   ' Removing First Empty rows
    Do While FRowCount < 5
        FRowCount = WorksheetFunction.CountA(Range("1:1"))
        If FRowCount < 5 Then Range("1:1").Delete
    Loop
    
    
    'Check if There is SN col
        If [A1].Value = "TRX_Store_Name" Then
            HaveSN = 0
            Else
            HaveSN = 1
        End If

    Range("A1").CurrentRegion.Select
    Selection.Cut
    
    If HaveSN = 1 Then
            Range("CA1").Select
            ActiveSheet.Paste
        Else
            Range("CB1").Select
            ActiveSheet.Paste
    End If


    ' List of Columns that should be in the data
    MyArr = Array("transaction_type", "sold_date", "customer_type", "customer_name", "department_name", _
    "category_name", "item_code", "description", "quantity", "net_sales", "rms_settlementtranskey", "tech_club", _
    "trx_store_name", "source_store_name", "order_store", "ref__", "class", "budget_channel")
    

  Dim c As Long, sd As Object
  Set sd = CreateObject("Scripting.Dictionary")
    For c = 0 To Selection.Columns.Count - 1
      sd.item(LCase(Cells(1, ActiveCell.Column + c).Value)) = Cells(1, ActiveCell.Column + c).Column
'      Debug.Print sd.Item(1)
     Next
    ''''''''''''''''''''''''''''''''''''
   
  '' To Read Sd Data to check
'   For Each k In sd.Keys
'      Print Key And Value
'      Debug.Print k, sd(k)
'    Next
     
     
     
   ''''''''''''''''''''''''''''''''''''
     For M = 0 To 15
        If Not sd.Exists(MyArr(M)) Then
            MsgBox (MyArr(M) & " Not Found")
            ErrCount = ErrCount + 1
        End If
    Next M
      
    If Not ErrCount >= 1 Then GoTo NoErr
EditOrNo = MsgBox("Continue Or Let U Edit First ???" & Chr(10) & "Press No To Edit Columns ", vbYesNo)
  If EditOrNo = vbNo Then
    Selection.Cut Range("A1")
    Exit Sub
  End If
  '''''''''
NoErr:

Columns(sd.item("trx_store_name")).Cut:         Range("B1").Select:       ActiveSheet.Paste
Columns(sd.item("source_store_name")).Cut:      Range("C1").Select:       ActiveSheet.Paste
Columns(sd.item("order_store")).Cut:            Range("D1").Select:       ActiveSheet.Paste
Columns(sd.item("ref__")).Cut:                  Range("E1").Select:       ActiveSheet.Paste
                                                Range("E1").Value = "Invoice.NO"
                                                
If sd.Exists("class") Then
Columns(sd.item("class")).Cut:                 Range("F1").Select:       ActiveSheet.Paste: End If
If sd.Exists("budget_channel") Then
Columns(sd.item("budget_channel")).Cut:         Range("G1").Select:       ActiveSheet.Paste: End If

                                                Range("F1").Value = "Class"
                                                Range("G1").Value = "Bud.Channel"
                                                Range("H1").Value = "Month.Class"
Columns(sd.item("transaction_type")).Cut:       Range("I1").Select:       ActiveSheet.Paste
Columns(sd.item("sold_date")).Cut:              Range("J1").Select:       ActiveSheet.Paste
                                                Range("K1").Value = "Week.Class"
Columns(sd.item("customer_type")).Cut:          Range("L1").Select:       ActiveSheet.Paste
Columns(sd.item("customer_name")).Cut:          Range("M1").Select:       ActiveSheet.Paste
                                                Range("N1").Value = "Bus.Line"
                                                Range("O1").Value = "Bud.Brand"
                                                Range("P1").Value = "Bud.Cat"
                                                Range("Q1").Value = "Item.Group"
Columns(sd.item("department_name")).Cut:        Range("R1").Select:       ActiveSheet.Paste
Columns(sd.item("category_name")).Cut:          Range("S1").Select:       ActiveSheet.Paste
Columns(sd.item("item_code")).Cut:              Range("T1").Select:       ActiveSheet.Paste
Columns(sd.item("description")).Cut:            Range("U1").Select:       ActiveSheet.Paste
Columns(sd.item("quantity")).Cut:               Range("V1").Select:       ActiveSheet.Paste
Columns(sd.item("net_sales")).Cut:              Range("W1").Select:       ActiveSheet.Paste
                                                Range("X1").Value = "Unit.Cost.(SC)"
                                                Range("Y1").Value = "T-Cost"
                                                Range("Z1").Value = "Domestic"
                                                Range("AA1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AB1").Value = "Bmob"
                                                Range("AC1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AD1").Value = "STK"
                                                Range("AE1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AF1").Value = "Other"
                                                Range("AG1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AH1").Value = "Comm"
                                                Range("AI1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AJ1").Value = "Decrease"
                                                Range("AK1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AL1").Value = "Miele"
                                                Range("AM1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AN1").Value = "X1"
                                                Range("AO1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AP1").Value = "X2"
                                                Range("AQ1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AR1").Value = "X3"
                                                Range("AS1").FormulaR1C1 = "=""TTL ""&RC[-1]"
                                                Range("AT1").Value = "TTL.Bonus"
                                                Range("AU1").Value = "Net Cost ÊßáÝÉ ãíÒÇä"
                                                Range("AV1").Value = "1st GP"
                                                Range("AW1").Value = "1st GP%"
                                                Range("AX1").Value = "Outlet Used Pro"
                                                Range("AY1").Value = "TTL Online comp"
                                                Range("AZ1").Value = "Noon Inv Error"
                                                Range("BA1").Value = "Ultra-Arkan-LG Ince"
                                                Range("BB1").Value = "Global Service"
                                                Range("BC1").Value = "Live Chat"
                                                Range("BD1").Value = "B2B Allow"
                                                Range("BE1").Value = "T-Forex"
                                                Range("BF1").Value = "Net.GP"
                                                Range("BG1").Value = "Net GP%"
                                                Range("BH1").Value = "Sub Cat"
Columns(sd.item("customer_phone")).Cut:              Range("BI1").Select:      ActiveSheet.Paste
Columns(sd.item("rms_settlementtranskey")).Cut: Range("BJ1").Select:      ActiveSheet.Paste
Columns(sd.item("refrence_number")).Cut:        Range("BK1").Select:      ActiveSheet.Paste
Columns(sd.item("full_discount")).Cut:              Range("BL1").Select:      ActiveSheet.Paste
Columns(sd.item("sold_price")).Cut:              Range("BM1").Select:      ActiveSheet.Paste
Columns(sd.item("installment")).Cut:              Range("BN1").Select:      ActiveSheet.Paste
Columns(sd.item("installmentwaydesc")).Cut:              Range("BO1").Select:      ActiveSheet.Paste
Columns(sd.item("trx_salesrep_number")).Cut:              Range("BP1").Select:      ActiveSheet.Paste
Columns(sd.item("web_number")).Cut:              Range("BQ1").Select:      ActiveSheet.Paste
Columns(sd.item("accountnumber")).Cut:              Range("BR1").Select:      ActiveSheet.Paste

    Rows(1).Insert
    

    
    LR = Cells(Rows.Count, 2).End(xlUp).Row

''' Make Serial Column

    If HaveSN = 0 Then
        [a2].Value = "S"
        [A3].Value = 1
        [A4].FormulaR1C1 = "=R[-1]c+1"
        Range("A4:A" & LR).Select
        Selection.FillDown
        Selection.Copy
        Selection.PasteSpecial xlPasteValues
    Else
         Columns("CA").Cut:     Range("A1").Select:       ActiveSheet.Paste
    End If
    
    ActiveSheet.Name = "Data"
    
''''Create Sheet Names
    If LCase(xUser) = "g" Then
        ActiveWorkbook.Names.Add Name:="Unii", RefersToR1C1:="=@Data!C16&@Data!C18"
        ActiveWorkbook.Names.Add Name:="Chnl", RefersToR1C1:="=@Data!C7"
        ActiveWorkbook.Names.Add Name:="Item", RefersToR1C1:="=@Data!C20"
        ActiveWorkbook.Names.Add Name:="Qty", RefersToR1C1:="=@Data!C22"
        ActiveWorkbook.Names.Add Name:="Sls", RefersToR1C1:="=@Data!C23"
    End If
            
    [A1].Select
    ' Get Columns numbers
     Set sd = Nothing
     Set sd = CreateObject("Scripting.Dictionary")
    
            For c = 0 To Range("B2").CurrentRegion.Columns.Count - 1
              sd.item(LCase(Cells(2, 1 + c).Value)) = Cells(2, 1 + c).Column
             Next
            
    '''''''''''''''''''''''' Start Fillind Data ''''''''''''''''''''
        
    Cells(3, sd.item("item.group")).Select
    ActiveCell.FormulaR1C1 = _
            "=VLOOKUP(RC[3],'D:\Gomich folder\[Help Cost.xlsb]RMS'!C1:C15,12,0)"
        
    Cells(3, sd.item("bus.line")).Select
    ActiveCell.FormulaR1C1 = _
        "=iferror(VLOOKUP(RC[3]&RC[4]&RC[5],'D:\Gomich folder\[Help Cost.xlsb]Brand'!C1:C8,6,0),0)"
    
    Cells(3, sd.item("bud.brand")).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[5],'D:\Gomich folder\[Help Cost.xlsb]RMS'!C1:C15,13,0)"
    
    Cells(3, sd.item("bud.cat")).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[4],'D:\Gomich folder\[Help Cost.xlsb]RMS'!C1:C15,14,0)"
    
    Cells(3, sd.item("department_name")).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[2],'D:\Gomich folder\[Help Cost.xlsb]RMS'!C1:C15,3,0)"
        
    Cells(3, sd.item("category_name")).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[1],'D:\Gomich folder\[Help Cost.xlsb]RMS'!C1:C15,4,0)"
    
    Range("N3:S3").Copy Range("N1")
    Range(Cells(3, sd.item("bus.line")), Cells(LR, sd.item("category_name"))).Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    
    Range("N1:S1").NumberFormat = ";;;"
    If Evaluate("=countif(N3:N300000,""#N/A"")") = 0 Then Range("N1:S1").Replace What:="=", Replacement:="X="
    
    Range([X3], Range("X" & LR)).Select
    Selection.FormulaR1C1 = _
         "=VLOOKUP(RC[-4],'[Help Cost.xlsb]AV Cost'!C1:C3,3,0)"
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
       
    myrange = Range("A2").CurrentRegion.Address
    
    Rows(2).AutoFilter
    

   
    Call ChnlFillRight
    
    'Total Cost
    Range("Y3").FormulaR1C1 = "=RC[-1]*RC[-3]"
    
    'Total Bonus
    Range("AT3").FormulaR1C1 = _
    "=RC[-1]+RC[-3]+RC[-5]+RC[-7]+RC[-9]+RC[-11]+RC[-13]+RC[-15]+RC[-17]+RC[-19]"
    'Net Cost - Trial
    Range("AU3").FormulaR1C1 = "=RC[-22]-RC[-1]"
    ' GP
    Range("AV3").FormulaR1C1 = "=RC[-25]-RC[-1]"
    
    'Subtotals
    Range("V1").Select
    Selection.Value = "=subtotal(9,V2:V" & LR & ")"
    Selection.Copy
    Range("W1,Y1,AA1,AC1,AE1,AG1,AI1,AK1,AM1,AO1,AT1:AV1,AX1:BF1,AQ1,AS1").Select
    ActiveSheet.Paste
    
    ' Grouping
    Range("X1,Z1,AB1,AD1,AF1,AH1,AJ1,AL1,AN1,AP1,AR1,AU1:BE1").Select
    Dim cll As Range
    For Each cll In Selection
        cll.EntireColumn.Group:    Next cll
        
    
    Range("C:F").Group
    Range("H:N").Group
    Range("BK:BR").Group
    
        
'''' Filter with SalesW-Cost
'   Fill Columns from Brand to Category with Outlet

    On Error Resume Next
    
    ActiveSheet.Range(myrange).AutoFilter field:=2, Criteria1:= _
        "=Sale With Cost W.H", Operator:=xlAnd
    Lines = Application.WorksheetFunction.CountA(Range("B:B").SpecialCells(xlCellTypeVisible))
    If Lines = 1 Then GoTo skip0
        
    If Range("3:3").EntireRow.Hidden = False Then
            Range("N3:O3").Select
            Selection.Value = "Sale With Cost"
            If Lines > 2 Then Range(Selection, Selection.End(xlDown)).FillDown
        Else
            ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(2)(1, 14).Resize(, 2).Select
            Selection.Value = "Sale With Cost"
            If Lines > 2 Then Range(Selection, Selection.End(xlDown)).FillDown
            Selection.Offset(, 3).Select
            Selection.Value = "Sale With Cost"
            If Lines > 2 Then Range(Selection, Selection.End(xlDown)).FillDown

    End If
    
skip0:
    On Error GoTo 0
    
    ActiveSheet.ShowAllData
    
    
GoTo skip1  ' TO always skip this stage
' Deleted Code
skip1:

    Range("F:H,K:K,N:Q,T:T,W:W").EntireColumn.AutoFit
    Columns("A:A").ColumnWidth = 6.29
    Columns("B:D").ColumnWidth = 16.43
    Columns("J:J").ColumnWidth = 15
    Columns("L:L").ColumnWidth = 17.14
    Columns("M:M").ColumnWidth = 28.14
    Columns("R:S").ColumnWidth = 19.71
    Columns("U:U").ColumnWidth = 27.71
    Columns("U:U").ColumnWidth = 33.29
    Columns("BJ:BJ").ColumnWidth = 26.57
    Columns("P:P").ColumnWidth = 10.86
    Rows(2).WrapText = True
    
    Range("A2:A" & LR).Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection
            .Borders.Value = 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With


    
''''' Cond Format if columns contains #N/A

    Range("N2:Q2").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=COUNTIF(N3:N300000,#N/A)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Font.Bold = True
        .Font.Color = -16776961
        .NumberFormat = """## ""@"
    End With
    
    
    'move useless data
    Range("A2").End(xlToRight).Offset(0, 1).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).EntireColumn.Select
    Selection.Copy
    Sheets.Add
    Range("B1").Select
    ActiveSheet.Paste
    ActiveSheet.Name = "more"
    Sheets("Data").Range("A:A").Copy
    Sheets("More").Range("A1").Select
    ActiveSheet.Paste
        For i = 1 To ActiveSheet.UsedRange.Columns.Count
            If Cells(2, i - deleted).Value = Empty Then
                Columns(i - deleted).Delete
                deleted = deleted + 1
            End If
        Next i
    
    Sheets("Data").Select
    Selection.Delete
    
    ActiveWindow.SplitRow = 2
    ActiveWindow.FreezePanes = True
    
    
'Replace "installment" & "cash commercial"
        Range("L:L").Replace What:="installment", Replacement:="Credit"
        Range("L:L").Replace What:="cash commercial", Replacement:="Cash"
    
''''''''''''''''    '''''''''''''''''   ''''''''''''''' ''''''''''''''''    ''''''''''''''''''

    'New Version Edit       ''' adding provision and arakan allowance
    Range("AX:AX").Insert
    [AX2].Value = "Provision"
    [AX3].Value = "=.0025*" & "W3"
    [AY1].Copy: [AX1].Select: ActiveSheet.Paste: Application.CutCopyMode = False
    
    'Miele MDA Formula
    Range("AL1").FormulaR1C1 = _
            "=AND(RC[-23]=""Miele"",RC[-22]=""MDA"",RC[-31]<>""Arkan"")"

    
    [BG3].Value = "=W3-Y3+AT3-AX3+AY3+AZ3+BA3-" & "BB3-BC3-BD3-BE3-BF3"    ' Edit NetGP Formula

    
'''' Test for 1 time then Delete if it run without errors
    
    'Range("Y3:CA3").Replace What:="@Qty", Replacement:="V3", FormulaVersion:=xlReplaceFormula2
    'Range("Y3:CA3").Replace What:="@Sls", Replacement:="W3", FormulaVersion:=xlReplaceFormula2
'''''''''
    
    '[AT3].FormulaR1C1 = "=IF(OR(RC[-31]=""Ariston"",RC[-31]=""Indesit"",RC[-31]=""Braun""),RC[-23]*0.05,0)"
    '[AT2].Value = "Agency Decrease"
        'Bonus Formula
    [AT3].FormulaR1C1 = "=RC[-1]+RC[-3]+RC[-5]+RC[-7]-RC[-9]+RC[-11]+RC[-13]+RC[-15]+RC[-17]+RC[-19]"
        'Forex
    [BF3].Value = "x=IF(G3<>""Outlet"",IFERROR(VLOOKUP(O3,{""Braun"",0.1;""Ariston"",0.1;""Indesit"",0.1;""Miele"",0.05},2,0)*W3,0)+IFERROR(VLOOKUP(T3,xcomp!S:U,3,0)*V3,0),0)"
        'AC Incentive
    '[BC3].Value = "x=IF(AND(G3<>""Deel"",OR(P3=""AC Domestic"",P3=""Ultra AC"")),w3*0,0)"
    
    'Commercial Compensation
    [AI3].Value = "x=IFERROR(VLOOKUP(T3,xComp!CS:CU,3,0)*V3,0)"
    
    'Noon Inv Error
    [BA3].Value = "x=IFERROR(VLOOKUP(BK3&T3,xComp!CR:CU,4,0),0)"
    
    'Outlet Used Pro
    [AY3].Value = "x=IF(G3=""Outlet"",W3*xComp!CH$1,0)+IFERROR(VLOOKUP(T3,xComp!CC:CE,3,0)*V3,0)"
    
    'B2B Allow Formula
    [BE3].Value = "x=IF(G3=""B2B BR"",W3*xComp!AU$1,0)"
    
    'Ultra & Arkan incentive Formula
    [BB3].Value = "x=IFERROR(IF(G3=""Arkan"",W3*xComp!CN$1,IF(P3=""Ultra LED"",W3*xComp!AP$1,VLOOKUP(T3,xComp!AL:AQ,3,0)*V3)),0)"
                 
    'Global Service
    [Bc3].Value = "x=IF(AND(G3=""Branches"",P3=""Laptop""),W3*xComp!BP$1,0)"
    
    'Live Chat
    [BD3].Value = "x=IFERROR(VLOOKUP(BK3,xComp!Y:AB,4,0)*W3,0)"
    
    'SubCategory
    '[BI3].Value = "x=IF(OR(O3=""Sale With Cost"",G3=""Outlet""),0,VLOOKUP(R3&S3,'D:\Gomich Folder\[Help Cost.xlsb]Brand'!B:Z,15,0))"    'MDA only ... Below MDA & Accessories
    [Bi3].Value = "x=IF(OR(O3=""Sale With Cost"",G3=""Outlet""),0,IF(P3=""Accessories"",IFERROR(VLOOKUP(T3,'D:\Gomich folder\[Help Cost.xlsb]Brand'!$BL:$BN,3,0),0),VLOOKUP(R3&S3,'D:\Gomich folder\[Help Cost.xlsb]Brand'!B:Z,15,0)))"
    
    'Online Comp
    [AZ3].Value = "x=IF(G3=""online BR"",IFERROR(VLOOKUP(Item,xComp!BT:BX,5,0),0)*v3,0)"
    
    'Unit Cost
'    [X1].Value = "x=IFERROR(VLOOKUP(T1,'[All Sep.xlsb]Retail'!$E:$O,VLOOKUP(J1,'[All Sep.xlsb]Retail'!$A:$B,2,0),0),VLOOKUP(T1,'D:\Gomich folder\[Help Cost.xlsb]AV Cost'!$A:$C,3,0))"
    [X1].Value = "x=IF(AND(G1=""Deel"",NOT(ISNUMBER(SEARCH(""RS"",BK1))))," & _
                    "IFERROR(VLOOKUP(T1,'[All Sep.xlsb]DEEL'!$E:$O,VLOOKUP(J1,'[All Sep.xlsb]DEEL'!$A:$B,2,0),0),VLOOKUP(T1,'[Help Cost.xlsb]AV Cost'!$A:$C,3,0))," & _
                    "IFERROR(VLOOKUP(T1,'[All Sep.xlsb]Retail'!$E:$O,VLOOKUP(J1,'[All Sep.xlsb]Retail'!$A:$B,2,0),0),VLOOKUP(T1,'[Help Cost.xlsb]AV Cost'!$A:$C,3,0)))"


    'ToTal Domestic
    [AA3].Value = "x=IFERROR(VLOOKUP(P3&R3&G3,xComp!A:K,11,0)*W3,0)"
    
    'ToTal Bmob
    [AC3].Value = "x=IFERROR(VLOOKUP(T3,xComp!AF:AH,3,0)*V3,0)"
    
 ''' Coloring Header Row and format
 
    Range("Q2:V2,H2:N2,A2:F2").Interior.Color = 10921638
    Range("G2").Interior.Color = 10086143
    Range("Z2:AS2,O2:P2").Interior.Color = 14348258
    
    Range("W2,Y2,BG2").Interior.Color = 14083324
    Range("AT2").Interior.Color = 15652797
    Range("AX2:BF2").Interior.Color = 13431551
    
    Range("Y2,AL2,BB2:BF2,AX2,Y1,AL1,BB1:BF1,AX1").Font.Color = RGB(255, 0, 0)
    Range("AY1:BA1,AT1,W1").Font.Color = RGB(0, 180, 0)
    
    Range("W:BG").NumberFormat = "#,##0"
    Range("BG1,AU1:AW1,Y1,W1").NumberFormat = "#,##0, K"
    
    [A1].Select: MsgBox ("Done")

    
End Sub

Private Sub ChnlFillRight()

    'formulat to determin if chnl need to be renamed
    Range([Z3], Range("Z" & LR)).Select
    Selection.Formula = "=OR(F3=""Btech Mini"",F3=""Jumia"",F3=""Noon"")"
    
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    
'Filter with these chnl then fillRight
    On Error Resume Next
    
    ActiveSheet.Range(myrange).AutoFilter field:=26, Criteria1:="True"
    
    Lines = Application.WorksheetFunction.CountA(Range("B:B").SpecialCells(xlCellTypeVisible))
    
            If Lines = 1 Then GoTo chkOutlet
    If Range("3:3").EntireRow.Hidden = False Then
            Range("G3").Select
            Selection.FillRight
                If Lines > 2 Then Range(Selection, Selection.End(xlDown)).FillRight
    Else
            ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(2)(1, 7).Select
            Selection.FillRight
                If Lines > 2 Then Range(Selection, Selection.End(xlDown)).Offset(, -1).Resize(, 2).FillRight

    End If
    
chkOutlet:

    ActiveSheet.ShowAllData
    ActiveSheet.Range(myrange).AutoFilter field:=7, Criteria1:="Outlet"
    
    Lines = Application.WorksheetFunction.CountA(Range("B:B").SpecialCells(xlCellTypeVisible))
    
            If Lines = 1 Then Exit Sub
    If Range("3:3").EntireRow.Hidden = False Then
            Range("N3:S3").Select
            Selection.Value = "Outlet"
                If Lines > 2 Then Range(Selection, Selection.End(xlDown)).FillDown
    Else
            ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(2)(1, 14).Resize(, 6).Select
            Selection.Value = "Outlet"
                If Lines > 2 Then Range(Selection, Selection.End(xlDown)).FillDown

    End If

    ActiveSheet.ShowAllData
    Range([Z3], Range("Z" & LR)).ClearContents
    On Error GoTo 0

End Sub






