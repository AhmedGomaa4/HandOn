
Sub offers_Calc()

    Dim Offers_Rows_Count, SrcRows As Integer
    Dim Formu, SRC As Range
    Set SRC = Range("A3").CurrentRegion
    SrcRows = SRC.Rows.Count
    
'sorting Data in Source sheet By Item then Start Date
    SRC.Sort key1:=Range("B3"), order1:=xlAscending, _
                key2:=Range("I3"), order1:=xlAscending, _
                    key3:=Range("D3"), order1:=xlAscending, _
                        Header:=xlYes
'Formulas for 'To Updated' and 'Duplicate same promo'
    Range("F3").Resize(SrcRows - 1).Formula = "=IF(I4=I5,MIN(E4,D4+30,D5-1),E4)"
    Range("F3").Resize(SrcRows - 1).Formula = "=IF(AND(B4=""Offer"",COUNTIFS(I:I,I4,C:C,C4)>1),""Yes"",""No"")"
    
'check for Duplicates
        If Evaluate("=countif(G:G,""Yes"")") > 0 Then
            MsgBox "Check Duplicates in column G then Run Macro Again"
            Exit Sub
        End If
'calculate Offers_Rows_Count then jump to Data sheet
    Offers_Rows_Count = Evaluate("=countif(B:B,""Offer"")")
    Sheets("Data").Select
    
'Formula Backup
    Set Formu = Range([a3], [a3].End(xlToRight))
    
'Clear and resize table then fill with formula
    Worksheets("Data").ListObjects("CalcT").DataBodyRange.Delete
    Worksheets("Data").ListObjects("CalcT").Resize Range("CalcT[#All]").Resize(Offers_Rows_Count + 1)
    Formu.Replace What:="x=", Replacement:="="
    Formu.Copy Range("CalcT")
    Formu.Replace What:="=", Replacement:="x="
    
'Kill table formulas except the 1st line
    Range("CalcT").Resize(Range("CalcT").Rows.Count - 1).Offset(1).Select
    Selection.Copy: Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    [A6].Select
    
'Refresh pivot tables
    ActiveWorkbook.RefreshAll
    
'Clearing ram
    Offers_Rows_Count = Empty
    SrcRows = Empty
    Set Formu = Nothing
    Set SRC = Nothing
    
    MsgBox "Done"
    
End Sub

