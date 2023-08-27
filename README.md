#This Code to be use in VBA Excel
Sub trial_SOG()
Dim file_order, file_PDP As Workbook
'Dim my_FileName As Variant

'Setting workbook Tools
Set file_order = ActiveWorkbook

If file_order.Worksheets("Input Order Raw").Range("A2").Value = "" Then
    Exit Sub
ElseIf file_order.Worksheets("Input PDP").Range("A3").Value = "" Then
    Exit Sub
Else
End If

'move order data from input order raw sheet to input order sheet
file_order.Worksheets("Input Order Raw").Select
row_inputorderraw = file_order.Worksheets("Input Order Raw").UsedRange.Rows.Count
file_order.Worksheets("Input Order Raw").Range("A2:D" & row_inputorderraw).Copy
file_order.Worksheets("Input Order").Select
file_order.Worksheets("Input Order").Range("A2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Input Order Raw").Range("I2:I" & row_inputorderraw).Copy
file_order.Worksheets("Input Order").Range("E2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Input Order Raw").Range("G2:G" & row_inputorderraw).Copy
file_order.Worksheets("Input Order").Range("F2").PasteSpecial Paste:=xlPasteValues

'move order data to wolf
file_order.Worksheets("Input Order Raw").Range("A1:A" & row_inputorderraw).Copy
file_order.Worksheets("WOLF Access").Range("A1").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Input Order Raw").Range("E1:E" & row_inputorderraw).Copy
file_order.Worksheets("WOLF Access").Range("B1").PasteSpecial Paste:=xlPasteValues
row_wolf = file_order.Worksheets("WOLF Access").UsedRange.Rows.Count
file_order.Worksheets("WOLF Access").Select
file_order.Worksheets("WOLF Access").Range("A1:B" & row_wolf).Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$B$" & row_wolf).RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlYes

'start staggering
file_order.Worksheets("Input Order").Select
row_input = file_order.Worksheets("Input Order").UsedRange.Rows.Count
With file_order.Worksheets("Input Order").Range("A2:A" & row_input)
    .NumberFormat = "General"
    .Value = .Value
End With
file_order.Worksheets("Input Order").Range("A2:B" & row_input).Copy
file_order.Worksheets("Database").Range("A2").PasteSpecial Paste:=xlPasteValues
row_database = file_order.Worksheets("Database").UsedRange.Rows.Count
file_order.Worksheets("Database").Range("A1:B" & row_database).RemoveDuplicates Columns:=1, Header:=xlYes
row_database = file_order.Worksheets("Database").UsedRange.Rows.Count

'lookup for tag 71

'Sort for tag 71
With file_order.Worksheets("Input Order").Sort
    .SortFields.Add Key:=Range("A1"), Order:=xlAscending
    .SortFields.Add Key:=Range("F1"), Order:=xlAscending
    .SortFields.Add Key:=Range("E1"), Order:=xlDescending
    .SetRange Range("A1:G" & row_input)
    .Header = xlYes
    .Apply
End With

' Copy Data PDP
row_pdpinput = file_order.Worksheets("Input PDP").UsedRange.Rows.Count
On Error Resume Next
file_order.Worksheets("Input PDP").ShowAllData
On Error GoTo 0
file_order.Worksheets("Input PDP").Range("A2:B" & row_pdpinput & ",M2:N" & row_pdpinput & ",Q2:W" & row_pdpinput).Copy

file_order.Worksheets("DATA PDP").Range("A1").PasteSpecial Paste:=xlPasteValues

row_pdp = file_order.Worksheets("DATA PDP").UsedRange.Rows.Count
file_order.Worksheets("DATA PDP").Range("L1").Value = "TOTAL PDP"
file_order.Worksheets("DATA PDP").Range("L2:L" & row_pdp).Value = "=Sum(E2:K2)"
file_order.Worksheets("Database").Range("C2:C" & row_database).Value = "=IFNA(XLOOKUP(A2,'DATA PDP'!A:A,'DATA PDP'!C:C),0)"
file_order.Worksheets("Database").Range("D2:D" & row_database).Value = "=IFNA(XLOOKUP(A2,'DATA PDP'!A:A,'DATA PDP'!D:D),0)"
file_order.Worksheets("Database").Range("E2:E" & row_database).Value = "=SUMIFS('Input Order'!E:E,'Input Order'!A:A,Database!A2)"
file_order.Worksheets("Database").Range("F2:F" & row_database).Value = "=IFERROR(ROUNDUP(E2/D2,0),0)"
file_order.Worksheets("Database").Range("G2:G" & row_database).Value = "=IF(F2>1,1,0)"
file_order.Worksheets("Database").Range("H2:H" & row_database).Value = "=SUMIFS('Input Order'!E:E,'Input Order'!A:A,Database!A2,'Input Order'!F:F,""<>71"")"

'mark between data that need stagger and not
file_order.Worksheets("Input Order").Range("G2:G" & row_input).Value = "=XLOOKUP(A2,'Database'!A:A,'Database'!G:G)"

'
file_order.Worksheets("Input Order").Range("K1").Value = "=IFERROR(MATCH(0,G:G,0),0)"
file_order.Worksheets("Input Order").Range("K1").Copy
file_order.Worksheets("Input Order").Range("K1").PasteSpecial Paste:=xlPasteValues

If file_order.Worksheets("Input Order").Range("K1").Value <> 0 Then

'If the remarks 0 which means doesn't need Stagger then all of it will be copied into summary and inputed in truck 1
file_order.Worksheets("Input Order").Range("A1:G1").AutoFilter field:=7, Criteria1:=0
file_order.Worksheets("Input Order").Range("A2:D" & row_input & ",F2:F" & row_input).SpecialCells(xlCellTypeVisible).Copy
file_order.Worksheets("Summary").Range("A2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Input Order").Range("E2:E" & row_input).SpecialCells(xlCellTypeVisible).Copy
file_order.Worksheets("Summary").Range("F2").PasteSpecial Paste:=xlPasteValues
row_sum = file_order.Worksheets("Summary").UsedRange.Rows.Count

'If the remarks 1 start staggering order and will be copied in calculation sheet
file_order.Worksheets("Input Order").Range("A1:G1").AutoFilter field:=7, Criteria1:=1
file_order.Worksheets("Input Order").Range("A2:D" & row_input & ",F2:F" & row_input).SpecialCells(xlCellTypeVisible).Copy
file_order.Worksheets("Calculation").Range("D2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Input Order").Range("E2:E" & row_input).SpecialCells(xlCellTypeVisible).Copy
file_order.Worksheets("Calculation").Range("K2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Calculation").Range("L2").PasteSpecial Paste:=xlPasteValues
row_calc = file_order.Worksheets("Calculation").UsedRange.Rows.Count

Else 'if all order data need stagger
'If the remarks 1 start staggering order and will be copied in calculation sheet
file_order.Worksheets("Input Order").Range("A1:G1").AutoFilter field:=7, Criteria1:=1
file_order.Worksheets("Input Order").Range("A2:D" & row_input & ",F2:F" & row_input).SpecialCells(xlCellTypeVisible).Copy
file_order.Worksheets("Calculation").Range("D2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Input Order").Range("E2:E" & row_input).SpecialCells(xlCellTypeVisible).Copy
file_order.Worksheets("Calculation").Range("K2").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Calculation").Range("L2").PasteSpecial Paste:=xlPasteValues
row_sum = file_order.Worksheets("Summary").UsedRange.Rows.Count
row_calc = file_order.Worksheets("Calculation").UsedRange.Rows.Count

End If

'starting calculation for SOG
max_iteration = WorksheetFunction.Max(file_order.Worksheets("Database").Range("F2:F" & row_database))
file_order.Worksheets("Calculation").Range("A2:A" & row_calc + 1).Value = "=IF(D2=D1,0,1)" 'to help find the last row of the data
file_order.Worksheets("Calculation").Range("B2:B" & row_calc).Value = "=XLOOKUP(D2,Database!A:A,Database!F:F,0)"
file_order.Worksheets("Calculation").Range("I2:I" & row_calc).Value = "=XLOOKUP(D2,Database!A:A,Database!C:C,0)"
file_order.Worksheets("Calculation").Range("J2:J" & row_calc).Value = "=XLOOKUP(D2,Database!A:A,Database!E:E,0)"
file_order.Worksheets("Calculation").Range("M2:M" & row_calc).Value = "=IF(C2<B2,IF(AND(I2>=2000,XLOOKUP(D2,Database!A:A,Database!D:D)<1000),MIN(L2,N2),IF(L2>0.2*XLOOKUP(D2,Database!A:A,Database!D:D),IF(H2=71,MIN(ROUND(K2/(B2/2),0),L2,N2),MIN(ROUND(K2/J2*XLOOKUP(D2,Database!A:A,Database!D:D),0),N2,L2)),MIN(L2,N2))),MIN(L2,N2))"
file_order.Worksheets("Calculation").Range("N2:N" & row_calc).Value = "=IF(A2>0,XLOOKUP(D2,Database!A:A,Database!D:D),N1-M1)"
file_order.Worksheets("Calculation").Range("O2:O" & row_calc).Value = "=L2-M2"
file_order.Worksheets("Calculation").Range("P2:P" & row_calc).Value = "=IF(AND(A2=0,A3=1),IF(N2>0,N2-M2,0),P3)"
file_order.Worksheets("Calculation").Range("Q2:Q" & row_calc).Value = "=IF(AND(P2>0,O2>0),MIN(P2,O2,R2),0)"
file_order.Worksheets("Calculation").Range("R2:R" & row_calc).Value = "=IF(A2>0,P2,IF(Q1>0,R1-Q1,R1))"
file_order.Worksheets("Calculation").Range("S2:S" & row_calc).Value = "=M2+Q2"
file_order.Worksheets("Calculation").Range("T2:T" & row_calc).Value = "=L2-S2"
file_order.Worksheets("Calculation").Range("D2:H" & row_calc).Copy
file_order.Worksheets("Summary").Range("A" & row_sum + 1).PasteSpecial Paste:=xlPasteValues
For i = 1 To max_iteration
    file_order.Worksheets("Calculation").Range("C2:C" & row_calc).Value = i
    file_order.Worksheets("Summary").Cells(1, 5 + i).Value = "Truck " & i
    file_order.Worksheets("Calculation").Range("S2:S" & row_calc).Copy
    file_order.Worksheets("Summary").Cells(row_sum + 1, 5 + i).PasteSpecial Paste:=xlPasteValues
    file_order.Worksheets("Calculation").Range("T2:T" & row_calc).Copy
    file_order.Worksheets("Calculation").Range("L2:L" & row_calc).PasteSpecial Paste:=xlPasteValues
Next

row_datapdp = file_order.Worksheets("DATA PDP").UsedRange.Rows.Count
file_order.Worksheets("DATA PDP").Range("A2:A" & row_datapdp).Copy
file_order.Worksheets("PDP Test").Range("A3").PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("DATA PDP").Range("E2:K" & row_datapdp).Copy
file_order.Worksheets("PDP Test").Range("B3").PasteSpecial Paste:=xlPasteValues

Call DivideByGI
Call RemoveDuplicateDTTruck
Call Tag_GI

file_order.Worksheets("Input Order").Select
ActiveSheet.ShowAllData

Worksheets("AccessData").Select

End Sub
Sub clear_tools()
Dim file_order As Workbook
Set file_order = ActiveWorkbook
    
file_order.Worksheets("Input Order").Cells.Delete
file_order.Worksheets("Database").Cells.Delete
file_order.Worksheets("DATA PDP").Cells.Delete
file_order.Worksheets("Summary").Cells.Delete
file_order.Worksheets("Calculation").Cells.Delete
file_order.Worksheets("AccessData").Cells.Delete
file_order.Worksheets("WOLF Access").Cells.Delete

row_splitbygidelete = file_order.Worksheets("Split By GI").UsedRange.Rows.Count
file_order.Worksheets("Split By GI").Range("A1:J" & row_splitbygidelete).Delete

file_order.Worksheets("Calculation").Range("A1").Value = "Remarks"
file_order.Worksheets("Calculation").Range("B1").Value = "Total Truck"
file_order.Worksheets("Calculation").Range("C1").Value = "Count"
file_order.Worksheets("Calculation").Range("D1").Value = "Customer"
file_order.Worksheets("Calculation").Range("E1").Value = "Customer Name"
file_order.Worksheets("Calculation").Range("F1").Value = "SKU"
file_order.Worksheets("Calculation").Range("G1").Value = "Description"
file_order.Worksheets("Calculation").Range("H1").Value = "Order Complexity"
file_order.Worksheets("Calculation").Range("I1").Value = "RPP"
file_order.Worksheets("Calculation").Range("J1").Value = "Total Order"
file_order.Worksheets("Calculation").Range("K1").Value = "Order Quantity"
file_order.Worksheets("Calculation").Range("L1").Value = "Remaining Qty"
file_order.Worksheets("Calculation").Range("M1").Value = "Qty Truck"
file_order.Worksheets("Calculation").Range("N1").Value = "Remaining MOT"
file_order.Worksheets("Calculation").Range("O1").Value = "Remaining Qty"
file_order.Worksheets("Calculation").Range("P1").Value = "Tagging"
file_order.Worksheets("Calculation").Range("Q1").Value = "Maximize"
file_order.Worksheets("Calculation").Range("R1").Value = "Final Remaining MOT"
file_order.Worksheets("Calculation").Range("S1").Value = "Final Truck QTY"
file_order.Worksheets("Calculation").Range("T1").Value = "Final Remaining QTY"
file_order.Worksheets("Summary").Range("A1").Value = "Cust"
file_order.Worksheets("Summary").Range("B1").Value = "Name"
file_order.Worksheets("Summary").Range("C1").Value = "SKU NDM"
file_order.Worksheets("Summary").Range("D1").Value = "Description"
file_order.Worksheets("Summary").Range("E1").Value = "Ord Comp"
file_order.Worksheets("Database").Range("A1").Value = "Sold-To Pt"
file_order.Worksheets("Database").Range("B1").Value = "Name (Sold to)"
file_order.Worksheets("Database").Range("C1").Value = "RPP"
file_order.Worksheets("Database").Range("D1").Value = "MOT"
file_order.Worksheets("Database").Range("E1").Value = "Total Order"
file_order.Worksheets("Database").Range("F1").Value = "Total Truck"
file_order.Worksheets("Database").Range("G1").Value = "Remarks"
file_order.Worksheets("Database").Range("H1").Value = "REG"
file_order.Worksheets("Database").Range("I1").Value = "Total PDP"

file_order.Worksheets("Input Order").Cells.Delete
file_order.Worksheets("DATA PDP").Cells.Delete

file_order.Worksheets("Input Order").Range("A1").Value = "Sold-To Pt"
file_order.Worksheets("Input Order").Range("B1").Value = "Name (Sold to)"
file_order.Worksheets("Input Order").Range("C1").Value = "Material"
file_order.Worksheets("Input Order").Range("D1").Value = "Material Description"
file_order.Worksheets("Input Order").Range("E1").Value = "Order Qty"
file_order.Worksheets("Input Order").Range("F1").Value = "Order Comp"
file_order.Worksheets("Input Order").Range("G1").Value = "Remarks"

row_gicalculation = file_order.Worksheets("PDP Test").UsedRange.Rows.Count
file_order.Worksheets("PDP Test").Select
file_order.Worksheets("PDP Test").Range("A3:H" & row_gicalculation).Value = ""

row_sogtest = file_order.Worksheets("SOG Test").UsedRange.Rows.Count
file_order.Worksheets("SOG Test").Select
file_order.Worksheets("SOG Test").Range("A2:D" & row_sogtest).Value = ""

file_order.Worksheets("Panel").Select

End Sub

Public Sub DivideByGI()

Dim file_order As Workbook
Set file_order = ActiveWorkbook

file_order.Worksheets("Split By GI").Move After:=Worksheets("Summary")


col_summary = file_order.Worksheets("Summary").UsedRange.Columns.Count
row_summary = file_order.Worksheets("Summary").UsedRange.Rows.Count
count_truck = col_summary - 5

Worksheets("Summary").Select
Worksheets("Summary").Range(Cells(1, 1), Cells(row_summary, col_summary)).Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

Dim i As Integer

For i = 1 To count_truck

If i = 1 Then
 paste_start = row_pastestart + 1
 paste_end = (row_summary - 1) * i
 
Else
 
 If file_order.Worksheets("Split By GI").UsedRange.Rows.Count = 1 Then
 Worksheets("Split By GI").Range("A1").Select
 Selection.Offset(1, 0).Select
 row_pastestart = ActiveCell.Row
 
 Else
 Worksheets("Split By GI").Range("A1").Select
 Selection.End(xlDown).Select
 row_pastestart = ActiveCell.Row
 
 End If
 
paste_start = row_pastestart + 1
paste_end = row_pastestart + row_summary - 1

End If

file_order.Worksheets("Summary").Range("A2:E" & row_summary).Copy
file_order.Worksheets("Split By GI").Range("A" & paste_start).PasteSpecial Paste:=xlPasteValues

file_order.Worksheets("Summary").Select
file_order.Worksheets("Summary").Cells(2, i + 5).Select
file_order.Worksheets("Summary").Range(Selection, Cells(row_summary, 5 + i)).Select
Selection.Copy
file_order.Worksheets("Split By GI").Range("F" & paste_start & ":F" & paste_end).PasteSpecial Paste:=xlPasteValues

file_order.Worksheets("Summary").Select
file_order.Worksheets("Summary").Cells(1, i + 5).Select
Selection.Copy
file_order.Worksheets("Split By GI").Range("G" & paste_start & ":G" & paste_end).PasteSpecial Paste:=xlPasteValues
file_order.Worksheets("Split By GI").Range("H" & paste_start & ":H" & paste_end).Value = i

file_order.Worksheets("Split By GI").Select
row_splitbygi = file_order.Worksheets("Split By GI").UsedRange.Rows.Count
row_splitbygiblank = row_splitbygi + 1
file_order.Worksheets("Split By GI").Range("E" & row_splitbygiblank).Value = "Test"
file_order.Worksheets("Split By GI").Range("F" & row_splitbygiblank).Value = ""
file_order.Worksheets("Split By GI").Range("F1:F" & row_splitbygiblank).Select
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete

Next i

file_order.Worksheets("Split By GI").Select
    file_order.Worksheets("Split By GI").Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select

file_order.Worksheets("Split By GI").Range("A1").Value = "Cust"
file_order.Worksheets("Split By GI").Range("B1").Value = "Name"
file_order.Worksheets("Split By GI").Range("C1").Value = "SKU NDM"
file_order.Worksheets("Split By GI").Range("D1").Value = "Description"
file_order.Worksheets("Split By GI").Range("E1").Value = "Ord Comp"
file_order.Worksheets("Split By GI").Range("F1").Value = "Qty"
file_order.Worksheets("Split By GI").Range("G1").Value = "Truck"
file_order.Worksheets("Split By GI").Range("H1").Value = "Truck Number"


End Sub



Public Sub RemoveDuplicateDTTruck()
    
Worksheets("Split By GI").Select
row_sbi = Worksheets("Split By GI").UsedRange.Rows.Count
Worksheets("Split By GI").Range("A1:H" & row_sbi).Select
Selection.Copy
Worksheets("AccessData").Select
        Worksheets("AccessData").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Worksheets("AccessData").Columns("C:F").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    row_ad = Worksheets("AccessData").UsedRange.Rows.Count
    Range("A1:D" & row_ad).Select
    ActiveSheet.Range("$A$1:$D$" & row_ad).RemoveDuplicates Columns:=Array(1, 2, 3, 4), _
        Header:=xlYes
    Range("A1").Select

Worksheets("AccessData").Range("E1").Value = "GI"
Worksheets("AccessData").Range("F1").Value = "GI Date"

End Sub

Public Sub Tag_GI()

Worksheets("AccessData").Select
row_accessdata = Worksheets("AccessData").UsedRange.Rows.Count
Worksheets("AccessData").Range("A2:D" & row_accessdata).Copy

Worksheets("SOG Test").Select
Worksheets("SOG Test").Range("A2").PasteSpecial Paste:=xlPasteValues

Worksheets("SOG Test").Range("A1").Select
 Selection.End(xlDown).Select
 row_copyend = ActiveCell.Row
Worksheets("SOG Test").Range("A2:F" & row_copyend).Copy

Worksheets("AccessData").Select
Worksheets("AccessData").Range("A2").PasteSpecial Paste:=xlPasteValues
Worksheets("AccessData").Range("A1").Select

row_sogtest = Worksheets("SOG Test").UsedRange.Rows.Count
row_sogtest1 = row_sogtest + 1
Worksheets("SOG Test").Select
Worksheets("SOG Test").Range("B" & row_sogtest1).Value = "Test"
Worksheets("SOG Test").Range("A:A").Select
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete

End Sub

Public Sub Clear_InputData()

Dim file_order As Workbook
Set file_order = ActiveWorkbook

row_pdpinput = file_order.Worksheets("Input PDP").UsedRange.Rows.Count
file_order.Worksheets("Input PDP").Range("A3:W" & row_pdpinput).Delete

If file_order.Worksheets("Input Order Raw").UsedRange.Rows.Count = 1 Then
    Exit Sub
Else
row_inputorderraw = file_order.Worksheets("Input Order Raw").UsedRange.Rows.Count
file_order.Worksheets("Input Order Raw").Range("A2:I" & row_inputorderraw).Delete
End If

file_order.Worksheets("Panel").Select

End Sub
