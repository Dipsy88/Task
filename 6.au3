; Task 6. Create two columns with desired number of records. Compare values in adjacent cell. Produce the report pass/fail

#include <Excel.au3>
#include <MsgBoxConstants.au3>

; location of the excel file
Local $path = @ScriptDir & "\Excel1.xlsx"

; Create application object and open a workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_SheetList Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, $path)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_SheetList Example", "Error opening workbook '" & @ScriptDir & "\Extras\_Excel2.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

; Write a part of a 2D array (1 for each column) to the active sheet in the active workbook
Local $aArray2D[8][2] = [[11, 11], [20, 22], [5,5], [1, 1], [200, 202], ["A","A"], ["apple","apple"],["apple","app"]]
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aArray2D, "B1")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example 3", "Error writing to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

$oRange = $oWorkbook.ActiveSheet.UsedRange.SpecialCells($xlCellTypeLastCell)

; The two columns can have uneven length, so add some more data to a particular column
; Write a 1D array to the active sheet in the active workbook
$lastRow = $oRange.row
Local $aArray1D[3] = ["AA", "BB", "CC"]

_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aArray1D, "B"&$lastRow+1)

$oRange = $oWorkbook.ActiveSheet.UsedRange.SpecialCells($xlCellTypeLastCell)
$lastRow = $oRange.row

; finally the comparison part
For $i = 1 To $lastRow
   $col1 = _Excel_RangeRead($oWorkbook, Default, "B"&$i)
   $col2 = _Excel_RangeRead($oWorkbook, Default, "C"&$i)
   If $col1 == $col2 Then
	  _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "True", "D"&$i)
   Else
	  _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "False", "D"&$i)
   EndIf
Next

MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
; finally close the excel file
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()