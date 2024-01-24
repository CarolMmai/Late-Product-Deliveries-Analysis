Dim objShell, strPath1, strAttr1, strAttr2, strAttr3, VariantList
Set objShell = CreateObject ("WScript.Shell")

'Variables that pass Username ------------
Set wshShell = CreateObject( "WScript.Shell" )
strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

'This is to check and open SAP
'Name the variables to be passed to open SAP instance
strPath1 = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""
strAttr1 = " -T ws start "

'Execute or Run Variables as objects
objShell.Run strPath1 & strAttr1
WScript.Sleep 3000
'----objShell.Run strAttr2 & strAttr3

'Look for a live session of the Landing SAP session and Select the main connection
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
   'Set connection = application.OpenConnection("210 - Template Light Harmonie Q2K : Quality P2K Link", True)
   Set connection = application.OpenConnection("110 - Template Light Harmonie P2K : Production MondayFridayCET Link", True)
   Set Session = connection.children(0)
End If


'---#####Check open connection of SAPGUI#####---
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If


'----Create a list of variants----------------------------------

' List of values using a normal dimension. Its traditional VB list.
Dim values(21)
values(0) = "/Transporter_1_variant"
values(1) = "/Transporter_2_variant"
values(2) = "/Transporter_3_variant"
values(3) = "/Transporter_4_variant"
values(4) = "/Transporter_5_variant"


' Constant for the duration of the Year and is adjusted in the new year.
sEndYr = "31.12.2023"

' Start the loop
For Each i In values
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFF3X"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").text = i
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = sEndYr
    session.findById("wnd[0]/usr/ctxtPA_STIDA").setFocus
    session.findById("wnd[0]/usr/ctxtPA_STIDA").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey 4
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "E:\Transporter_late_product_deliveries_data\Files_to_append"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = i & ".xlsx"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 15
    session.findById("wnd[2]/tbar[0]/btn[11]").press
    session.findById("wnd[1]/tbar[0]/btn[11]").press

'----------------------------------Close Excel-----------------------------------------------------
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate, " _
    & "(Debug)}!\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Excel.exe'")
On Error Resume Next
For Each objProcess In colProcessList
    objProcess.Terminate()
Next

Wscript.sleep(2000)


' Close Open Office
'----------------------------------Ends Open Office-----------------------------------------------------
' ... (Your existing code)
' Construct PowerShell Command (PS syntax)
strPSCommand = "Get-Process -Name 'soffice.bin' | Stop-Process -Force"


' Consruct DOS command to pass PowerShell command (DOS syntax)
strDOSCommand = "powershell -command " & strPSCommand & ""

' Create shell object
Set objShell = CreateObject("Wscript.Shell")

' Execute the combined command
Set objExec = objShell.Exec(strDOSCommand)

' Read output into VBS variable
strPSResults = objExec.StdOut.ReadAll


'---------------------------------End of Report GENERATION------------------------
Wscript.sleep(4000)

Next

' End of the loop


If Err.Number <> 0 Then
            MsgBox "An error occurred. Error description: " & Err.Description
        End If

        On Error GoTo 0

        ' Delay for 1 second before proceeding to the next iteration
        WScript.Sleep 1000
	'Next

' Delay for 30 seconds
WScript.Sleep 15000 ' Delay for 15 seconds
	
	
'---------------------------------Close SAP GUI-----------------------------------

session.findbyid("wnd[0]").close
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set SapGuiAuto = Nothing
Wscript.sleep(2000)

'---------------------------------Close SAPLOGON-------------------------------------

Set WshShell = CreateObject("WScript.Shell")

' Define the path to the batch script
batchScriptPath = "\\E\Generic_all_purpose_scripts\SAPScript\close_saplogon.bat"

' Run the batch script
WshShell.Run batchScriptPath, 0, True

Set WshShell = Nothing

'--------------------------------End of Script-------------------------------------------
