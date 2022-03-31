@echo off

REM COMMAND SWITCHES FOR POWERPOINT
REM /n "<FileName>"	Create a new file with the specified file name and path. Example: start powerpnt.exe /n "C:\PATH\TO\POWEROINT\my_presentation.pptx"
REM /s "<FileName>"	Start the slide show with the specified presentation file. Example: start powerpnt.exe /s "C:\PATH\TO\POWEROINT\my_presentation.pptx"
REM /p "<FileName>"	Print the presentation to the default printer. Example:start powerpnt.exe /p "C:\PATH\TO\POWEROINT\my_presentation.pptx"
REM /pt <PrinterName> "" "" "<FileName>"	Print the presentation to the specified printer.  Example: start powerpnt.exe /pt "HP LaserJet" "" "" "C:\PATH\TO\POWEROINT\my_presentation.pptx"
REM /m "<FileName>" <MacroName>	Open the presentation file specified in FileName and run the macro specified in MacroName in it. Example: start powerpnt.exe /m "C:\PATH\TO\POWEROINT\my_presentation.pptx" HelloWorld
REM /c	Start PowerPoint with no presentation. Example: start powerpnt.exe /c
REM /b	Start PowerPoint with a blank presentation. Example: start powerpnt.exe /b
REM /o "<FileName1>" "<FileName2>"	Start PowerPoint with a list of presentations to open. Example: start powerpnt.exe /o "C:\PATH\TO\POWEROINT\my_presentation.pptx" "C:\PATH\TO\POWEROINT\my_presentation2.pptx"

REM NAVIGATE TO THE MS OFFICE FOLDER CONTAINING THE POWERPOINT EXECUTABLE
cd C:\Program Files\Microsoft Office\root\Office16

REM START POWERPNT.EXE AND POINT TO THE POWERPOINT FILE USING THE /S SWITCH TO MAKE IT START PLAYING
start powerpnt.exe /s "C:\PATH\TO\POWEROINT\my_presentation.pptx"

