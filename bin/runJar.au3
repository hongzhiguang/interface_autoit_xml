Run("cmd.exe")
WinWaitActive("[CLASS:ConsoleWindowClass]")

Send("cd ..\\tool{ENTER}")
Send("java -jar OmcaXmlNBI.jar{ENTER}")