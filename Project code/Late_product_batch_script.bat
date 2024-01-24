@echo off

:: Start the VBScript in the background
start /b cscript "\\E\Transporter_late_product_deliveries\Late_product_data_extraction_VBS_script.vbs"

:: Wait for 50 seconds or until the VBScript finishes
timeout /t 50 /nobreak

:: Terminate wscript.exe and cscript.exe processes
taskkill /F /IM wscript.exe
taskkill /F /IM cscript.exe

echo The VBScript execution is completed or has timed out.

echo Running the Python script...
E\Transporter_late_product_deliveries\Late_product_data transformation_and_loading.py
ping 127.0.0.1 -n 6 > nul

echo Python script completed.

:: Terminate wscript.exe and cscript.exe processes (in case they were started again)
taskkill /F /IM wscript.exe
taskkill /F /IM cscript.exe

echo The Windows Host session is ended.
exit