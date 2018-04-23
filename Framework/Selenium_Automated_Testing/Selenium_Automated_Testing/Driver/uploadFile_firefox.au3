;Firefox browser
Local $file = "C:\Users\longvu\Pictures\Game\lords-of-the-fallen.jpg"
WinWaitActive("File Upload")
if WinExists("File Upload") Then
	; Both code are working fine	
	WinActivate("File Upload")   
	ControlFocus("File Upload","","[CLASS:Edit; INSTANCE:1]")   
	ControlSetText("File Upload", "", "[CLASS:Edit; INSTANCE:1]", $file)
    ControlClick("File Upload", "", "[CLASS:Button; TEXT:&Open]")        
    
	; Both code are working fine	
	;ControlFocus("File Upload","","[CLASS:Edit; INSTANCE:1]")   
	;ControlSetText("File Upload", "", "Edit1", $file)
	;Sleep(1500)
	;ControlClick("File Upload", "", "Button1")
EndIf