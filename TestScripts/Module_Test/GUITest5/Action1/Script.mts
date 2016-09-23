Window("Start menu").WinObject("Items View").WinList("Items View").Select "Notepad" @@ hightlight id_;_2014435208_;_script infofile_;_ZIP::ssf1.xml_;_
Window("Notepad").WinEditor("Edit").Type "Test" @@ hightlight id_;_24516488_;_script infofile_;_ZIP::ssf2.xml_;_
Window("Notepad").WinMenu("SystemMenu").Select "File;Save	Ctrl+S"
Window("Notepad").Dialog("Save As").WinEdit("File name:").Set "*.txt" @@ hightlight id_;_3539848_;_script infofile_;_ZIP::ssf4.xml_;_
Window("Notepad").Dialog("Save As").WinButton("Cancel").Click @@ hightlight id_;_6689640_;_script infofile_;_ZIP::ssf5.xml_;_
