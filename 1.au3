; Task 1. Launch the excel and print the version number

#include<excel.au3>
#include <Constants.au3>
#include <File.au3>


; location of the excel file
Local $path = @ScriptDir & "\Excel1.xlsx"

; open excel
Local $oExcel = _Excel_Open()
Local $oWorkbook = _Excel_BookOpen($oExcel, $path)


; print the excel version
Local $OfficeCaption, $OfficeVersion, $ospp_path
getOfficeVersion($OfficeCaption, $OfficeVersion)

; finally close the excel file
MsgBox($MB_SYSTEMMODAL, "Message", "Closing excel now")
_Excel_BookClose($oWorkbook, False)
$oExcel.Quit()


; function to get the excel version
Func getOfficeVersion(ByRef $OfficeCaption, ByRef $OfficeVersion)
    SplashTextOn("Checking", 'Determining which version of' & @CRLF & 'Microsoft Excel is installed ...', 400, 100, -1, -1, 16)
    ;ported code from http://support.moonpoint.com/os/windows/office/office_versions.php

    Local $strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\cimv2")

    $colApps = $objWMIService.ExecQuery _
        ("Select * from Win32_Product Where Caption Like '%Microsoft Office Professional%'")
    For $objApp in $colApps
        $OfficeCaption = $objApp.Caption
        $OfficeVersion = $objApp.Version
        SplashOff()
        MsgBox ( 0, 'Office Version',  $OfficeCaption & ", " & $OfficeVersion)
		 ConsoleWrite("The excel version is : " &$OfficeVersion)
        Return
    Next
    SplashOff()
EndFunc