; Task 2. Switch between different tabs, and print count of tabs

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

; loop all the sheets
For $i = 0 To UBound($aWorkSheets) - 1
   $oWorkbook.Sheets($aWorkSheets[$i][0]).Activate
   # Sleep 0.4 second for each tab to show the tab switching to the user
   Sleep(400)
Next

; print the number of total tabs
MsgBox($MB_SYSTEMMODAL, "Tabs", "Number of tabs is " &Ubound($aWorkSheets))
ConsoleWrite("Number of tabs is " &Ubound($aWorkSheets))

; finally close the excel file
MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()
