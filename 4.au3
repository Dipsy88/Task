; Task 4. Print different control types (for example ribbon, list) within the Excel

#include <Excel.au3>
#include <MsgBoxConstants.au3>

; location of the excel file
Local $path = @ScriptDir & "\Excel1.xlsx"

; close excel if it was open
Local $oExcel = _Excel_Open()
$oExcel.Quit()

; Create application object and open a workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_SheetList Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, $path)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_SheetList Example", "Error opening workbook '" & @ScriptDir & "\Extras\_Excel2.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
 EndIf

 ; wait some time
Sleep(200)
; Display a list of all worksheets for a specific Workbook
Local $aWorkSheets = _Excel_SheetList($oWorkbook)
Sleep(500)
Local $aWorkSheets = _Excel_SheetList($oWorkbook)

; output that will be shown at the end
global $output = ""


; send keys to open the border menu, better than clicking
Send("!F")
Sleep(300)
Send("T")
Sleep(300)
Send("C")
Sleep(300)
Send("{TAB 5}")
Send("{LEFT}")

; first go to the top of dropdown
Send("{TOP 10}")

While(1)
   clipPut("")
   Send("!M")
   Sleep(100)
   Send("^c")
   $text = ClipGet()
   $newText = ClipGet()
   $output = $output & " " & $newText
   Send("{ESC}")
   Send("{LEFT 3}")

   Send("{DOWN}")
   Sleep(100)
   If $newText="" Then
	  ExitLoop
   EndIf

WEnd

; the output
MsgBox($MB_SYSTEMMODAL, "Menus", "Menu names are " &$output)
ConsoleWrite("Menu items are:  " &$output)


; finally close the excel file
MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()