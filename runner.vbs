Dim strBasePath : strBasePath = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))

'Kill the processes
ConsoleOutputBlankLine(1)
Call KillProcess("UFT.exe")
Call KillProcess("QtpAutomationAgent.exe")
Call KillProcess("iexplore.exe")

'Create QTP object
ConsoleOutputBlankLine(1)
Set QTP = CreateObject("QuickTest.Application")
ConsoleOutput("Launching QTP Application")
QTP.Launch
QTP.Visible = TRUE

msgbox "Hola Mundo"
