#include <Constants.au3>

;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Jonathan Bennett (jon at autoitscript dot com)
;
; Script Function:
;   Demo of using multiple lines in a message box
;
 WinWaitActive("[TITLE:Product Location Maintenance; CLASS:SunAwtFrame]");

Opt("SendKeyDelay",50)
HotKeySet("^!x", "MyExit")

Const $sFilePath = "C:\Users\awhite\B17Viega.csv"
$aArray = FileReadToArray($sFilePath)
If @error Then
   MsgBox($MB_SYSTEMMODAL, "", "There was an error reading the file. @error: " & @error) ; An error occurred reading the current script file.
Else
   For $i = 0 To UBound($aArray) - 1 ; Loop through the array.
	  $sCurLocation = StringSplit($aArray[$i],",")
	  If @error <> 1 Then
		 ;excel uses double quotes at the beginning and end of the string if a double quote was in the String
		 ;excel also adds a double quote in front of the double quote in the String
		 $sCurLocation[3] = StringRegExpReplace($sCurLocation[3],'^\"', "") ; remove double quote at beginning
		 $sCurLocation[3] = StringRegExpReplace($sCurLocation[3],'\"$', "") ; remove double quote at end
		 $sCurLocation[3] = StringRegExpReplace($sCurLocation[3],'""', '"') ; channge double double quote to single double quote
		 $sCurLocation[3] = StringMid($sCurLocation[3],1,StringLen($sCurLocation[3]))
		 enterProductId("." & $sCurLocation[1], $sCurLocation[3])
		 sleep(2000) ; GIVES USER TIME TO INTERUPT PROGRAM
		 enterLocation($sCurLocation[1],$sCurLocation[2], $sCurLocation[3])
	  Else
		 MsgBox($MB_SYSTEMMODAL, "Line Split", "Failed to split a line")
	  EndIf
   Next
EndIf



;$sProduct = ".br549"
;enterProductId($sProduct)
;waitForLag()
;If productLoaded($sProduct) == True Then
;   MsgBox($MB_SYSTEMMODAL, "Title", "Success")
;Else
;   MsgBox($MB_SYSTEMMODAL, "Title", "Failure")
;EndIf

Func enterProductId($sId, $sShortDesc)
   While MouseGetCursor() == 15
	  sleep(60)
   WEnd

MouseMove(596,118)
   MouseClick("primary")
   Sleep(100)
   While Hex(PixelGetColor(596,118),6) <> "F5F5C8"
	  MouseClick("primary")
	  Sleep(100)
   WEnd
   Send("{ctrldown}ac{ctrlup}")
   Sleep(100)
   Send($sId)
   Send("{enter}")
  ; SLEEP(5000)
   While MouseGetCursor() == 15
	  sleep(60)
   WEnd

WinWaitActive("[TITLE:Product Location Maintenance - " & $sShortDesc & "; CLASS:SunAwtFrame]");

   While MouseGetCursor() == 15
	  sleep(60)
   WEnd
EndFunc

; pause until the wait cursor goes away
Func waitForLag()
   While MouseGetCursor() == 15
   WEnd
EndFunc

; Check to see if the text is still just the part id
Func productLoaded($sId)
   ;send("^a")
   send("^c")
   if ClipGet == $sId then
	  return False
   Else
	  return True
   EndIf
EndFunc

Func enterLocation($sPartId, $sLocation, $sShortDesc)
   MouseMove(46,253)
   MouseClick("primary",46,263,1)
   While Hex(PixelGetColor(46,263),6) <> "F5F5C8"
	  MouseClick("primary",46,263,1)
	  Sleep(2000)
   WEnd
   $bLooking = True
   While $bLooking ; find the location line we need
	  ClipPut("") ; clear the clipboard, the following code willnot copy anyting to the clipboard if nothing is there
	  send("^c") ; copy the current contents of type field
	  Sleep(100)
	  $sType = ClipGet()
	  if $sType == "S" Then ; we have found the stock line
		 $bLooking = False
	  Else
		 If $sType == "" Then ; this is a blank line, setup the new line
			Send("S")
			sleep(60)
			Send("{tab}")
			sleep(60)
			Send("{tab}")
			sleep(60)
			Send("{tab}")
			sleep(60)
			Send("P")
			sleep(60)
			Send("{shiftdown}")
			sleep(60)
			Send("{tab}")
			sleep(60)
			Send("{tab}")
			sleep(60)
			Send("{tab}")
			sleep(60)
			Send("{shiftup}") ; back up to the location
			sleep(50)
			$bLooking = False
		 Else
			Send("{down}")
		 EndIf
	  EndIf
   WEnd
   enterStockLoc($sLocation, $sShortDesc)
   Sleep(1000)
   saveChanges($sShortDesc)
EndFunc

Func enterStockLoc($sLocation, $sShortDesc)
   sleep(100)
   Send("{tab}")
   Sleep(100)
   Send($sLocation, $SEND_RAW) ; some strange location out there
   Send("{tab}")
EndFunc

Func saveChanges($sShortDesc)
   While MouseGetCursor() == 15
	  sleep(60)
   WEnd
   Send("{altdown}5{altup}")
   WinWaitActive("[TITLE:Product Location Maintenance - " & $sShortDesc & "; CLASS:SunAwtDialog]");
   Send("BIN UPDATE{tab}{enter}")
EndFunc

Func MyExit()
    Exit
EndFunc