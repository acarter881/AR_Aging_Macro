Private Declare Function WNetGetUser Lib "mpr.dll" _
  Alias "WNetGetUserA" (ByVal lpName As String, _
  ByVal lpUserName As String, lpnLength As Long) As Long

Const NoError = 0       'The funct call was successful
Public Function GetUsername()

  Const lpnLength As Integer = 255  ' return string buffer size
  Dim status As Integer   ' get the return buffer space
  Dim lpName, lpUserName As String  ' user info

  ' Assign the buffer size constant to lpUserName.
  lpUserName = Space$(lpnLength + 1)
  status = WNetGetUser(lpName, lpUserName, lpnLength) ' get log-on name

  ' See whether error occurred.
  If status = NoError Then
      lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
      GetUsername = lpUserName
  Else
      ' An error occurred.
      GetUsername = "Unable to get name."
      End
  End If
End Function

Private Sub credit_limit_report()

' This is the second worksheet

'Boilerplate
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' Clean up the report
Range("A:A").Delete
Rows("1:6").Delete
Rows("2:2").Delete
Range("G:G").Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Cells.EntireColumn.AutoFit
Range("A1").Select

'Boilerplate
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub customer_master_data()

'Boilerplate
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' This is the third worksheet
Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
    :=Array(1, 1), TrailingMinusNumbers:=True
Range("A1").Select

'Boilerplate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
End Sub

Sub aging_report()
' This is the first worksheet. Currently, the above subprocedures need to be run (on their respective worksheets) before this one.

'BOILERPLATE
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'SELECT THE FIRST WORKSHEET (I.E., THE AGING WORKSHEET)
Sheets("aging").Select

'VARIABLE FOR DISPLAYING A MESSAGE TO THE USER
Dim Final As Variant

'VARIABLES FOR THE TOTAL AMOUNTS FROM THE AGING REPORT
Dim Original_Total_Amount As Long
Dim Edited_Total_Amount As Long
Dim Total_Difference As Long

'MACROS FOR THE CREDIT LIMITS AND MASTER DATA WORKSHEETS
Sheets("CL").Select
Call credit_limit_report
Sheets("MD").Select
Call customer_master_data

'SELECT THE FIRST WORKSHEET (I.E., THE AGING WORKSHEET)
Sheets("aging").Select

'DELETE COLUMN A
Range("A:A").Delete

'SELECT CELL A1
Range("A1").Select

'SELECT ALL OF THE DATA IN THE WORKSHEET AND SORT BY COLUMN A
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo

'RETRIEVE THE TOTAL AMOUNT (I.E., THE AMOUNT BEFORE DELETING ANY DATA)
Range("F1").Select
Selection.End(xlDown).Select
Original_Total_Amount = Selection.End(xlDown).Value

'DELETE THE DATA AT THE BOTTOM OF THE WORKSHEET (I.E., THE DATA THAT DO NOT INCLUDE A CUSTOMER ID)
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
Selection.EntireRow.Delete

'RETRIEVE THE EDITED TOTAL AMOUNT (I.E., THE AMOUNT FROM SUMMING THE COLUMN)
ActiveCell.Offset(0, 1).Select
Dim iCol As Long
iCol = ActiveCell.Column
Edited_Total_Amount = WorksheetFunction.Sum(Columns(iCol))

'STORE THE DIFFERENCE IN A VARIABLE
Total_Difference = Original_Total_Amount - Edited_Total_Amount

'INSERT A ROW FOR THE HEADERS AND INPUT THE HEADER NAMES
Rows(1).Insert
Range("A1").Formula2R1C1 = "Acct#"
Range("B1").Formula2R1C1 = "CL"
Columns("F:F").Select
Selection.Delete Shift:=xlToLeft
Columns("G:G").Select
Selection.Delete Shift:=xlToLeft
Columns("D:D").Select
Selection.Insert Shift:=xlToRight
Selection.Insert Shift:=xlToRight
Selection.Insert Shift:=xlToRight
Selection.Insert Shift:=xlToRight
Range("C1").Formula2R1C1 = "BU Name"
Range("D1").Formula2R1C1 = "Address"
Range("E1").Formula2R1C1 = "Address 2"
Range("F1").Formula2R1C1 = "City"
Range("G1").Formula2R1C1 = "State"
Range("H1").Formula2R1C1 = "Zip"
Range("I1").Formula2R1C1 = "Country Code"
Range("J1").Formula2R1C1 = "Total"
Range("K1").Formula2R1C1 = "Current"
Range("L1").Formula2R1C1 = "1-30 days"
Range("M1").Formula2R1C1 = "31-60 days"
Range("N1").Formula2R1C1 = "61-90 days"
Range("O1").Formula2R1C1 = "91-180 days"
Range("P1").Formula2R1C1 = "181+ days"
Range("Q1").Formula2R1C1 = "BU"

'GET LAST ROW IN COLUMN AND FILL COLUMN Q (EXCLUDING THE HEADER) WITH "BD1"
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Range("Q2:Q" & last_row).Formula = "BD1"

'SELECT DATA AND FILTER
Range("A1").Select
ActiveCell.CurrentRegion.Select
Selection.AutoFilter

'FILTER OUT ALL ITEMS THAT ARE NOT US AND NOT PR
Range("I:I").AutoFilter Field:=9, Criteria1:="<>US", Criteria2:="<>PR", Operator:=xlAnd

'DELETE THE FILTERED ROWS THEN CLEAR THE FILTER
ActiveSheet.Range("I1:I" & last_row).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.ShowAllData

'REPLACE ALL INSTANCES OF "PR" WITH "US"
Columns("I").Replace "PR", "US", xlPart, xlByRows, True

'XLOOKUP FOR CL COLUMN
Dim last_row_2 As Long
last_row_2 = Cells(Rows.Count, 1).End(xlUp).Row
Range("B2").FormulaR1C1 = "=XLOOKUP(RC[-1],CL!C[-1],CL!C)"
Range("B2").AutoFill Destination:=Range("B2:B" & last_row_2)

'COPY AND PASTE AS VALUES
Range("B:B").Copy
Range("B:B").PasteSpecial xlPasteValues
Application.CutCopyMode = False

'FILTER OUT 0, 1, 2, 3, 5, AND N/A
Range("B:B").AutoFilter Field:=2, Criteria1:=Array("#N/A", "0", "1", "2", "3", "5"), Operator:=xlFilterValues

'DELETE THE FILTERED ROWS THEN CLEAR THE FILTER
ActiveSheet.Range("B1:B" & last_row_2).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.ShowAllData

'XLOOKUP FOR BU NAME COLUMN
Dim last_row_3 As Long
last_row_3 = Cells(Rows.Count, 1).End(xlUp).Row
Range("C2").FormulaR1C1 = "=XLOOKUP(RC[-2],MD!C[-2],MD!C[-1])"
Range("C2").AutoFill Destination:=Range("C2:C" & last_row_3)

'XLOOKUP FOR ADDRESS COLUMN
Range("D2").FormulaR1C1 = "=XLOOKUP(RC1,MD!C1,MD!C)"
Range("D2").AutoFill Destination:=Range("D2:D" & last_row_3)

'XLOOKUP FOR ADDRESS 2 COLUMN
Range("E2").FormulaR1C1 = "=XLOOKUP(RC[-4],MD!C[-4],MD!C)"
Range("E2").AutoFill Destination:=Range("E2:E" & last_row_3)

'XLOOKUP FOR CITY COLUMN
Range("F2").FormulaR1C1 = "=XLOOKUP(RC[-5],MD!C[-5],MD!C)"
Range("F2").AutoFill Destination:=Range("F2:F" & last_row_3)

'XLOOKUP FOR STATE COLUMN
Range("G2").FormulaR1C1 = "=XLOOKUP(RC[-6],MD!C[-6],MD!C)"
Range("G2").AutoFill Destination:=Range("G2:G" & last_row_3)

'XLOOKUP FOR ZIP COLUMN
Range("H2").FormulaR1C1 = "=XLOOKUP(RC[-7],MD!C[-7],MD!C)"
Range("H2").AutoFill Destination:=Range("H2:H" & last_row_3)

'COPY AND PASTE AS VALUES
Range("C:H").Copy
Range("C:H").PasteSpecial xlPasteValues
Application.CutCopyMode = False

'FILTER OUT N/A AND NON-US STATES
Range("G:G").AutoFilter Field:=7, Criteria1:=Array("#N/A", "GU", "MP", "VI"), Operator:=xlFilterValues

'DELETE THE FILTERED ROWS THEN CLEAR THE FILTER
ActiveSheet.Range("G1:G" & last_row_3).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.ShowAllData

'INACTIVE FILTER
Dim str As String
Dim objMatches As Object

'FIND THE LAST ROW (AGAIN)
last_row_4 = Cells(Rows.Count, "C").End(xlUp).Row

'STRING REGEX LOOP
For i = 2 To last_row_4
    str = Cells(i, "C")
    
    'INACTIVE REGEX
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "\bINACTIVE\b"
    objRegExp.Global = True
    
    Set objMatches = objRegExp.Execute(str)
    
    If objMatches.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
        'Rows(i).EntireRow.Delete 'CAN'T DELETE THE ROW BECAUSE IT WILL MESS UP THE INTEGRITY OF THE LOOP
    End If
    
    'DIP REGEX
    Set objRegExp_2 = CreateObject("VBScript.RegExp")
    objRegExp_2.Pattern = "\bDIP\b"
    objRegExp_2.Global = True
    
    Set objMatches_2 = objRegExp_2.Execute(str)
    
    If objMatches_2.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
    End If
    
    'BD REGEX
    Set objRegExp_3 = CreateObject("VBScript.RegExp")
    objRegExp_3.Pattern = "\bBD\b"
    objRegExp_3.Global = True
    
    Set objMatches_3 = objRegExp_3.Execute(str)
    
    If objMatches_3.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
    End If
    
    'CAREFUSION REGEX
    Set objRegExp_4 = CreateObject("VBScript.RegExp")
    objRegExp_4.Pattern = "\bCAREFUSION\b"
    objRegExp_4.Global = True
    objRegExp_4.IgnoreCase = True
    
    Set objMatches_4 = objRegExp_4.Execute(str)
    
    If objMatches_4.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
    End If
    
    'CASH REGEX
    Set objRegExp_5 = CreateObject("VBScript.RegExp")
    objRegExp_5.Pattern = "\bCASH\b"
    objRegExp_5.Global = True
    objRegExp_5.IgnoreCase = True
    
    Set objMatches_5 = objRegExp_5.Execute(str)
    
    If objMatches_5.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
    End If
Next i

'NUMBER REGEX LOOP
For j = 2 To last_row_4
    str_2 = Cells(j, "A")
    
    '1002857542 REGEX
    Set objRegExp_6 = CreateObject("VBScript.RegExp")
    objRegExp_6.Pattern = "1002857542"
    objRegExp_6.Global = True
    
    Set objMatches_6 = objRegExp_6.Execute(str_2)
    
    If objMatches_6.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
    End If
    
    '1002658753 REGEX
    Set objRegExp_7 = CreateObject("VBScript.RegExp")
    objRegExp_7.Pattern = "1002658753"
    objRegExp_7.Global = True
    
    Set objMatches_7 = objRegExp_7.Execute(str_2)
    
    If objMatches_7.Count > 0 Then
        Cells(i, "R").Value = "TO BE DELETED"
    End If
Next j

'SET HEADER FOR THE DELETED COLUMN
Range("R1").Value = "Header for Deletion"

'SELECT DATA AND FILTER
Range("A1").Select
ActiveCell.CurrentRegion.Select
Selection.AutoFilter
Selection.AutoFilter

'FILTER FOR TO BE DELETED
Range("R:R").AutoFilter Field:=18, Criteria1:="TO BE DELETED", Operator:=xlFilterValues

'DELETE THE FILTERED ROWS THEN CLEAR THE FILTER
ActiveSheet.Range("R1:R" & last_row_4).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
ActiveSheet.ShowAllData

'DELETE COLUMN R
Columns("R:R").Delete

'REPLACE HYPHENS WITH BLANKS ON THE ADDRESS AND ADDRESS 2 COLUMNS
Columns("D:E").Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows

'FIND THE LAST ROW (AGAIN)
last_row_5 = Cells(Rows.Count, "A").End(xlUp).Row

'REPLACE BLANKS WITH ZEROES ON THE CURRENT - 181+ DAYS COLUMNS
Range("K1:P" & last_row_5).Replace What:="", Replacement:="0", LookAt:=xlWhole, SearchOrder:=xlByRows

'CHANGE STYLE TO COMMA THEN REMOVE DECIMAL PLACES
Range("B:B,H:H,J:P").Select
Selection.NumberFormat = "#,##0"

'RETRIEVE FULL PATH TO SAVE (USER'S DESKTOP IN THIS EXAMPLE)
User = GetUsername()
Pre = "C:\Users\"
Post = "\Desktop"
full_path = Pre & User & Post
ChDir full_path

Dim my_str

my_str = Format(Date, "mm.dd.yyyy")

full_csv_name = full_path & "\BD1 US Aging " & my_str & ".csv"

ActiveWorkbook.SaveAs Filename:=full_csv_name, FileFormat:=xlCSVUTF8, CreateBackup:=False
       
'Boilerplate
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Script is complete"
        
End Sub
