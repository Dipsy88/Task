; Task 5. Expand the borders list, count the number of options in this list, and click on any of the border. Verify that selected border is applied.

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
Sleep(1000)
; Display a list of all worksheets for a specific Workbook
Local $aWorkSheets = _Excel_SheetList($oWorkbook)
Sleep(1000)
Local $aWorkSheets = _Excel_SheetList($oWorkbook)

; send keys to open the border menu, better than clicking
Send("!H")
Send("B")

Sleep(200)
Send("{DOWN}")
Sleep(200)
Send("{DOWN}")
Sleep(100)

Send("{DOWN}")
Sleep(200)
Send("{DOWN}")
Sleep(200)
Send("{DOWN}")
Sleep(100)
Send("{DOWN}")


Send("{ENTER}")
HotKeySet ("^{F1}", "gettext")

Local $clipboard
ConsoleWrite("The text is: " &$clipboard)
Sleep(400)

; finally close the excel file
MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()

Func gettext ()
    $clipboard = ClipGet ()
    Send ("^c")
    ToolTip (ClipGet ())
    Sleep (5000)
    ToolTip ("")
    ClipPut ($clipboard)
EndFunc


; Equivalent to "All Borders" for range given
;$oExcel.ActiveSheet.range("F2:G3").Borders.LineStyle = 2