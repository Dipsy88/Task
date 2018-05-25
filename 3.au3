; Task 3. User must be prompted with input box (accepts text of any tab names), and then focus must be automatically switched on to the specific tab

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

; Display a list of all worksheets for a specific Workbook
Local $aWorkSheets = _Excel_SheetList($oWorkbook)
Local $var = InputBox("Select Tab", "Enter the name of the tab you want to focus on")

; Display a list of all worksheets for a specific Workbook
Local $aWorkSheets = _Excel_SheetList($oWorkbook)
Local $isPresent = False

; loop all the tabs
For $i = 0 To UBound($aWorkSheets) - 1
   If $var == $aWorkSheets[$i][0] Then
	  $isPresent = True
	  $oWorkbook.Sheets($aWorkSheets[$i][0]).Activate
	  MsgBox($MB_SYSTEMMODAL, "Error", "Success! Selected tab is focused now")
   EndIf
Next

; If unsuccessful to find the selected tab
 If $isPresent == False Then
   MsgBox($MB_SYSTEMMODAL, "Error", "Failure! Selected tab is not present")
   ConsoleWrite("Selected tab is not present")
EndIf


; finally close the excel file
MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()
