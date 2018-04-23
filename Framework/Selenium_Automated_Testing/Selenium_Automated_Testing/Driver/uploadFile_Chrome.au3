;MsgBox(0,"test",@WorkingDir)
Local $file = @WorkingDir & "\UploadingFile.png"
;MsgBox(0,"test",$file)
WinWaitActive("Open")
if WinExists("Open") Then
	; Both code are working fine	
	WinActivate("Open")   
	ControlFocus("Open","","[CLASS:Edit; INSTANCE:1]")   
	ControlSetText("Open", "", "[CLASS:Edit; INSTANCE:1]", $file)
    ControlClick("Open", "", "[CLASS:Button; TEXT:&Open]")        
    
	; Both code are working fine	
	;ControlFocus("Open","","[CLASS:Edit; INSTANCE:1]")   
	;ControlSetText("Open", "", "Edit1", $file)
	;Sleep(1500)
	;ControlClick("Open", "", "Button1")
EndIf

;---------------------------------------------------------------------------

;Firefox browser
Local $file = "D:\Automation Framework\Framework\Selenium2FW - InventiveIT\Selenium2FW\Selenium2FW\autoIt\UploadingFile.png"
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