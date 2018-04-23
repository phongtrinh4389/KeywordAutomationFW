#include <IE.au3>
; Internet Explorer is partly integrated in shell.application
$oShell = ObjCreate("shell.application") ; Get the Windows Shell Object
$oShellWindows=$oShell.windows   ; Get the collection of open shell Windows
$MyIExplorer=""
for $Window in $oShellWindows    ; Count all existing shell windows
  ; Note: Internet Explorer appends a slash to the URL in it's window name
  if StringInStr($Window.LocationURL,"http://localhost:8080/dvwa/vulnerabilities/upload/") then
      $MyIExplorer=$Window
      exitloop
  endif
next
;WinActivate("PDF Converter")
$oForm = _IEGetObjByName($MyIExplorer, "uploaded")
_IEAction($oForm, "click")

