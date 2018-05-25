; Task 7. Type any text with in the cell, read it and print the text, it could be either first few characters or middle characters

#include <Array.au3>
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

; Write text
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "Test String", "B1")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example 1", "Error writing to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; MsgBox($MB_SYSTEMMODAL, "The written first few characters", "String successfully written.")

; Read data from a single specified cell on the active sheet of the specified workbook
Local $sResult = _Excel_RangeRead($oWorkbook, Default, "B1")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
ConsoleWrite("The first few characters of the written text is: " & StringLeft($sResult, 6))

; the final output
MsgBox($MB_SYSTEMMODAL, "Output", "The first four characters of the written text is: "& StringLeft($sResult, 4))
MsgBox($MB_SYSTEMMODAL, "Output", "Total characters in the written text is: "& $sResult)

MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
; finally close the excel file
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()