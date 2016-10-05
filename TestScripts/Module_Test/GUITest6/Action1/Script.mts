'Launch the Applciation

invokeapplication ("C:\Program Files (x86)\Solumina\Solumina Browser\sf32.exe /A http://20.17.50.36:8080/solumina-G7/gateway /XC") @@ hightlight id_;_65678_;_script infofile_;_ZIP::ssf10.xml_;_


while not loaded
wait(1)
loaded =Delphiwindow("DelphiWindow_1").DelphiEdit("UserName").Exist
Wend
'Login in to the application
If Loaded=true Then
	Delphiwindow("DelphiWindow_1").DelphiEdit("UserName").Set"RSAHEB2"
	Delphiwindow("DelphiWindow_1").DelphiEdit("Password").Set"RSAHEB2"
End If
Delphiwindow("DelphiWindow_1").DelphiObject("Logon").Click 31,10 @@ hightlight id_;_1050340_;_script infofile_;_ZIP::ssf12.xml_;_

'click the window
'Delphiwindow("DelphiWindow").DelphiObject("ok").Click 601,407

wait(10)

OptionalStep.DelphiWindow("DelphiWindow").DelphiObject("ok").Click 34,15 @@ hightlight id_;_1642710_;_script infofile_;_ZIP::ssf19.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject").Click 14,13 @@ hightlight id_;_6163218_;_script infofile_;_ZIP::ssf24.xml_;_
DelphiWindow("DelphiWindow_3").DelphiObject("DelphiObject").Click 67,21 @@ hightlight id_;_8327042_;_script infofile_;_ZIP::ssf25.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_2").Click 58,31 @@ hightlight id_;_2162734_;_script infofile_;_ZIP::ssf26.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_2").DblClick 57,39 @@ hightlight id_;_2162734_;_script infofile_;_ZIP::ssf27.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_3").Drag 128,15 @@ hightlight id_;_1380636_;_script infofile_;_ZIP::ssf28.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_3").Drop 128,375 @@ hightlight id_;_1380636_;_script infofile_;_ZIP::ssf29.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_4").Click 9,7 @@ hightlight id_;_4067028_;_script infofile_;_ZIP::ssf30.xml_;_
DelphiWindow("DelphiWindow_4").DelphiEdit("DelphiEdit").Set "Newtitle" @@ hightlight id_;_5506086_;_script infofile_;_ZIP::ssf31.xml_;_
DelphiWindow("DelphiWindow_4").DelphiObject("DelphiObject").Click 32,18 @@ hightlight id_;_1313746_;_script infofile_;_ZIP::ssf32.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_5").Click 790,323 @@ hightlight id_;_3214086_;_script infofile_;_ZIP::ssf33.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_6").Click 11,10 @@ hightlight id_;_1376312_;_script infofile_;_ZIP::ssf34.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_7").Click 22,15 @@ hightlight id_;_2427442_;_script infofile_;_ZIP::ssf35.xml_;_
DelphiWindow("DelphiWindow_5").DelphiObject("DelphiObject").Click 30,5 @@ hightlight id_;_8130252_;_script infofile_;_ZIP::ssf36.xml_;_
DelphiWindow("DelphiWindow_2").DelphiObject("DelphiObject_8").Click 3,17 @@ hightlight id_;_3279426_;_script infofile_;_ZIP::ssf37.xml_;_
 @@ hightlight id_;_2297364_;_script infofile_;_ZIP::ssf22.xml_;_

