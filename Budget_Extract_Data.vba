Private HelpBK, CurrBK As Workbook, sh As Variant
Sub MainOne()
    
    Application.ScreenUpdating = False
'check if xGoDesign exist
    On Error GoTo GoCreatexGoDesign
    Sheets("xGoDesign").Select
    On Error GoTo 0

    Dim SDAarr, MDAarr, RTLarr As Variant
    Dim TargetCol, LR, mns As Long
    Set CurrBK = ActiveWorkbook
    Set HelpBK = Workbooks.Open("D:\Gomich Folder\Help Cost.xlsb")
    'Set HelpBK = Workbooks.Open("C:\Users\ahmed.abbady\Desktop\Help\Help Cost.xlsb")
    Sheets("Budget").Select
    CurrBK.Activate


    MDAarr = Array("Cairo MDA", "Alex MDA", "Delta 1 MDA ", "Delta 2 MDA", "Upper Egy MDA", "Chains MDA", "Miele-Arkan")
    
    
    For Each sh In MDAarr
        
        Sheets(sh).Select
        Range("1:1").Insert
        Range("B630").Value = "Inv Dis % :"
        Call HelperColumns
        Range(Cells(2, 1), Cells.SpecialCells(xlLastCell)).AutoFilter
        
        For mns = 6 To 17
            TargetCol = 6
            Range("C6:D30").Copy
            LR = Sheets("xGoDesign").Cells(Rows.Count, "D").End(xlUp).Row + 1
            Sheets("xGoDesign").Cells(LR, "D").PasteSpecial xlPasteValues


            For Each cll In Sheets("xGoDesign").Range("F2:Q2")
                If cll.Value = "Cash" Or cll.Value = "Credit" Then GoTo DDD:
                xfind = Range("E:E").Find(cll).Row
                Cells(xfind, mns).Offset(2).Resize(25).Copy
                Sheets("xGoDesign").Cells(LR, TargetCol).PasteSpecial xlPasteValues

DDD:
                TargetCol = TargetCol + 1
            Next cll

            Sheets("xGoDesign").Cells(LR, "C").Resize(25).Value = ActiveSheet.Name
            Sheets("xGoDesign").Cells(LR, "B").Resize(25).Value = Cells(3, mns).Value

       Next mns

    Next sh

        'Msgbox "MDA Done"

'>>>>End OF MDA''''''''''''''''''



    SDAarr = Array("CAIRO SDA", "ALEX SDA", "DELTA 1 SDA", "DELTA 2 SDA", "UPPER EGY. SDA", "CHAINS SDA")

    For Each sh In SDAarr
        Sheets(sh).Select
        Range("1:1").Insert
        Range("B550").Value = "Inv Dis % :"
        Call HelperColumns
        Range(Cells(2, 1), Cells.SpecialCells(xlLastCell)).AutoFilter

        For mns = 6 To 17
            TargetCol = 6
            Range("C5:D19").Copy
            LR = Sheets("xGoDesign").Cells(Rows.Count, "D").End(xlUp).Row + 1
            Sheets("xGoDesign").Cells(LR, "D").PasteSpecial xlPasteValues


            For Each cll In Sheets("xGoDesign").Range("F2:Q2")
                If cll.Value = "Cash" Or cll.Value = "Credit" Then GoTo BBB:

                xfind = Range("E:E").Find(cll).Row
                Cells(xfind, mns).Offset(1).Resize(15).Copy
                Sheets("xGoDesign").Cells(LR, TargetCol).PasteSpecial xlPasteValues
BBB:
                TargetCol = TargetCol + 1

            Next cll

            Sheets("xGoDesign").Cells(LR, "C").Resize(15).Value = ActiveSheet.Name
            Sheets("xGoDesign").Cells(LR, "B").Resize(15).Value = Cells(3, mns).Value

       Next mns

    Next sh

   'Msgbox "SDA Done"
'>>>>End OF SDA''''''''''''''''''




'    RTLarr = Array("Retail Sales", "B Tech X", "Market Place", "B2B inside Branches Sales", "Online inside Branches Sales", _
            "CC inside Branches Sales", "B2B Sales", "CC Sales", "Online Sales", "Tech Club Sales ")
    
        RTLarr = Array("Branches Sales", "Call Center Inside B Tech X ", "Call Center Inside Branches 1", _
        "Call Center Sales", "Online Inside B Tech X Sales", "Online Inside Branches Sale 1", _
        "Online Sales", "B2B Inside B Tech X Sales", "B2B Inside Branches Sales 1", _
        "B2B  Sales", "B Tech X Sales", "Deel Sales", _
        "Market Place Sales", "Noon Sales")
        
    
    For Each sh In RTLarr
        Sheets(sh).Select
        Call HelperColumns
        Range(Cells(4, 1), Cells.SpecialCells(xlLastCell)).AutoFilter
        Call cashCredit
        
        For mns = 6 To 39 Step 3
        
            Range("C21:D230").SpecialCells(xlCellTypeVisible).Copy
            LR = Sheets("xGoDesign").Cells(Rows.Count, "D").End(xlUp).Row + 1
            Sheets("xGoDesign").Cells(LR, "D").PasteSpecial xlPasteValues
            
            Cells(21, mns).Resize(210, 3).SpecialCells(xlCellTypeVisible).Copy
            Sheets("xGoDesign").Cells(LR, 6).PasteSpecial xlPasteValues
            Cells(730, mns + 1).Resize(226).SpecialCells(xlCellTypeVisible).Copy
            Sheets("xGoDesign").Cells(LR, 9).PasteSpecial xlPasteValues
            
            Sheets("xGoDesign").Cells(LR, "C").Resize(37).Value = ActiveSheet.Name
            Sheets("xGoDesign").Cells(LR, "B").Resize(37).Value = Cells(4, mns).Value
            
       Next mns
        
    Next sh
    
'>>>>End OF Retail ''''''''''''''''''

' Outlet
        Sheets("Outlet Sales").Select
        Range("4:6").Insert
        Range("9:10").Insert
        Range("C10:AD10").FormulaR1C1 = "=(R[-3]C-R[-2]C)/R[-3]C"
        Range("C3:N10").Copy
        
        Range("C20").PasteSpecial Paste:=xlPasteValues, Transpose:=True
        Selection.Columns(2).Value = ActiveSheet.Name
        Selection.Columns(3).Value = "Outlet"
        Selection.Columns(4).Value = "Outlet"
        Selection.Columns(6).ClearContents
        Selection.Copy
        LR = Sheets("xGoDesign").Cells(Rows.Count, "B").End(xlUp).Row + 1
        Sheets("xGoDesign").Cells(LR, "B").PasteSpecial xlPasteValues
        
 
'>>>>End OF Outlet ''''''''''''''''''

' Service
        
        Sheets("Service out ").Select
        Range("C122:N122").Copy
        Range("c150").PasteSpecial xlPasteValues
        
        Range("C154:N154").FormulaR1C1 = "=(R[-135]C+R[-44]C-R[-120]C-R[-70]C+R[-71]C)"
        Range("C153:N153").FormulaR1C1 = "=R[1]C/R[-3]C"
        
        Range("C150:N153").Copy
        Range("E160").PasteSpecial Paste:=xlPasteValues, Transpose:=True
        Selection.Offset(, -2).Resize(, 2).Value = "Service"
        Selection.Offset(, -3).Resize(, 1).Value = ActiveSheet.Name
        Range("C4:N4").Copy:     Range("A160").PasteSpecial xlPasteValues, Transpose:=True
        Range("A160:H171").Copy
        LR = Sheets("xGoDesign").Cells(Rows.Count, "B").End(xlUp).Row + 1
        Sheets("xGoDesign").Cells(LR, "B").PasteSpecial xlPasteValues
        
        Sheets("xGoDesign").Select
        Sheets("xGoDesign").Range("A:B").Insert
        Sheets("xGoDesign").Range("B2").Value = "Sheet Group"
        Sheets("xGoDesign").Range("A2").Value = "Unii DEFG"
        
        Range("B3").Value = "x=VLOOKUP(E3,chnl_Map!A:E,5,0)"
        Range("C3").Value = "x=VLOOKUP  (E3,chnl_Map!A:E,4,0)"
        Range("H:J,U:W").NumberFormat = "#,##0"
        Range("D2:W2").BorderAround Weight:=xlMedium

MsgBox "All Done"

Exit Sub
GoCreatexGoDesign:
    Call CreatexGoDesign
End Sub

Private Sub cashCredit()

' cashCredit Macro



    ActiveSheet.AutoFilterMode = False
    
    Range("G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q,R:R").Select
    Selection.Insert Shift:=xlToRight
    Range("K:K,M:M,O:O,Q:Q,S:S,U:U,W:W,Y:Y,AA:AA,AC:AC,AE:AE").Select
    Range("H:H,J:J,L:L,N:N,P:P,R:R,T:T,V:V,X:X,Z:Z,AB:AB,AD:AD").Select
    Selection.Insert Shift:=xlToRight

    Range("G21,G27,G31,G35,G39,G43,G47,G51,G59,G63,G67,G71,G75,G79,G86,G90,G110,G115,G119,G123,G132,G136,G140,G149,G153,G161,G165,G169,G182,G198,G202,G206,G210,G214,G222,G226,G230,G132").FormulaR1C1 = "=R[-2]C[-1]"
   
    Range("H21,H27,H31,H35,H39,H43,H47,H51,H59,H63,H67,H71,H75,H79,H86,H90,H110,H115,H119,H123,H132,H136,H140,H149,H153,H161,H165,H169,H182,H198,H202,H206,H210,H214,H222,H226,H230,H132").FormulaR1C1 = "=R[-1]C[-2]"
    
    Range("G744").FormulaR1C1 = "=R[-5]C[-1]"
    Range("G750,G754,G758,G762,G766,G770,G774,G782,G786,G790,G794,G798,G802,G809,G813,G833,G838,G842,G846,G855,G859,G863,G872,G876,G884,G888,G892,G905,G921,G925,G929,G933,G937,G945,G949,G953").FormulaR1C1 = "=R[-1]C[-1]"
    
    Range("H:G").Copy
    Range("J1,M1,P1,S1,V1,Y1,AB1,AE1,AH1,AK1,AN1").Select
    ActiveSheet.Paste
    
    Range("G4").Select
    ActiveSheet.Range(Cells(4, 1), Cells.SpecialCells(xlLastCell)).AutoFilter field:=7, Criteria1:="<>"
     
End Sub

Private Sub HelperColumns()

    ActiveSheet.AutoFilterMode = False
    Columns.Hidden = False
    
''''''''''''''''' Temp Code just to test''''''''''''''''''

'     SH = ActiveSheet.Name
'     Set CurrBK = ActiveWorkbook
'     Set HelpBK = Workbooks.Open("D:\Gomich Folder\Help Cost.xlsb")
'
'     Sheets("Budget").Select
'     CurrBK.Activate
'
'''''''''''''End of temp'''''''''''''''

     MDAarr = Array("Cairo MDA", "Alex MDA", "Delta 1 MDA ", "Delta 2 MDA", "Upper Egy MDA", "Chains MDA", "Miele-Arkan")

     SDAarr = Array("CAIRO SDA", "ALEX SDA", "DELTA 1 SDA", "DELTA 2 SDA", "UPPER EGY. SDA", "CHAINS SDA")

     'RTLarr = Array("Retail Sales", "B Tech X", "Market Place", "B2B inside Branches Sales", "Online inside Branches Sales", _
                 "CC inside Branches Sales", "B2B Sales", "CC Sales", "Online Sales", "Tech Club Sales ")

     RTLarr = Array("Branches Sales", "Call Center Inside B Tech X ", "Call Center Inside Branches 1", _
        "Call Center Sales", "Online Inside B Tech X Sales", "Online Inside Branches Sale 1", _
        "Online Sales", "B2B Inside B Tech X Sales", "B2B Inside Branches Sales 1", _
        "B2B  Sales", "B Tech X Sales", "Deel Sales", _
        "Market Place Sales", "Noon Sales")
        
    If IsNumeric(Application.Match(sh, MDAarr, 0)) Then
        HelperColRNG = "G:J"
    ElseIf IsNumeric(Application.Match(sh, SDAarr, 0)) Then
        HelperColRNG = "M:P"
    Else
        HelperColRNG = "A:D"
        Range("1:1").Insert
    End If

    
    Range("A:A").Delete
    HelpBK.Activate
    Range(HelperColRNG).Copy
    CurrBK.Activate
    Range("A:A").Insert
    

End Sub
Private Sub CreatexGoDesign()
    
    Sheets.Add Before:=Sheets(1)
        ActiveSheet.Name = "xGoDesign"
        Range("A2").Value = "Big Channel"
        Range("B2").Value = "Ref"
        Range("C2").Value = "Channel (Sheet Name)"
        Range("D2").Value = "Brand"
        Range("E2").Value = "Cat"
        Range("F2").Value = "Sales Value :"
        Range("G2").Value = "Cash"
        Range("H2").Value = "Credit"
        Range("I2").Value = "G.P % :"
        Range("J2").Value = "Sales Allow. % :"
        Range("K2").Value = "Display % :"
        Range("L2").Value = "Special discount for installment % :"
        Range("M2").Value = "Special discount for top dealers % :"
        Range("N2").Value = " Salesmen Incentives  % :"
        Range("O2").Value = "Rent %  :"
        Range("P2").Value = "Inv Dis % :"
        Range("Q2").Value = "T. Sales Allow %  :"
        Range("R2").Value = "Net GP %"
        Range("R3").Value = "x=K3-IFERROR(SUM(L3:Q3),0)"
        Range("S2").Value = "GP B4 Allow"
        Range("S3").Value = "x=H3*K3"
        Range("T2").Value = "GP After Allow"
        Range("T3").Value = "x=H3*T3"
        Range("U2").Value = "Allow val"
        Range("U3").Value = "x=U3-V3"
        Range("R:R,T:T").Interior.Color = 16247773
        
    Call MainOne
    
End Sub
Private Sub MoveSheetToBegnin()

    For Each hh In Selection
        Sheets(hh.Value).Move Before:=ActiveWorkbook.Sheets(1)
    Next hh
        
End Sub
Private Sub ConvetToNewVersion()

    CommArr = Array("Cairo MDA", "Alex MDA", "Delta 1 MDA ", "Delta 2 MDA", "Upper Egy MDA", "Chains MDA", "Miele-Arkan", "Total MDA", "CAIRO SDA", "ALEX SDA", "DELTA 1 SDA", "DELTA 2 SDA", "UPPER EGY. SDA", "CHAINS SDA", "Total SDA")
    
    RTLOutletArr = Array("Retail Sales", "B Tech X", "Market Place", "B2B inside Branches Sales", "Online inside Branches Sales", _
            "CC inside Branches Sales", "B2B Sales", "CC Sales", "Online Sales", "Tech Club Sales ", "Total Retail Sales", "Outlet Sales")

    For Each sh In CommArr
        Sheets(sh).Select
        Columns.Hidden = False
        Range("B1,D1,F1,H1,J1,L1,N1,P1,R1,T1,V1,X1,Z1,AB1,AD1,AF1,AH1,AJ1,AL1,AN1,AP1,AR1,AT1,AV1,AX1,AZ1,BB1,BD1").EntireColumn.Delete
        Range("A:A").Insert
    Next sh
    
    For Each sh In RTLOutletArr
        Sheets(sh).Select
        Columns.Hidden = False
        Range("B3:E3,G3,I3,K3,M3,O3,Q3,S3,U3,W3,Y3,AA3,AC3,AE3,AG3,AI3,AK3,AM3,AO3,AQ3,AS3,AU3,AW3,AY3,BA3,BC3,BE3,BG3,BI3").EntireColumn.Delete
    Next sh

    
    Sheets("Service out ").Select
        Columns.Hidden = False
        Range("B4,D4,F4,H4,J4,L4,N4,P4,R4,T4,V4,X4,Z4,AB4,AD4,AF4,AH4,AJ4,AL4,AN4,AP4,AR4,AT4,AV4,AX4,AZ4,BB4,BD4").EntireColumn.Delete
        Range("A:A").Insert
        MsgBox "Now it's new Version"
End Sub

