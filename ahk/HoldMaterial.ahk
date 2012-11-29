#NoEnv
#SingleInstance force
keyDelay = 25
i = 0
OutputVar = ""
ItemRecordIterations = 10
controlDelay = 40
directReaderOn = ""
checkDirectReader = ""
StartTime = ""
EndTime = ""
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%

	Gui, Library:Add, Text, x12 y10 w380 h30 , Pick a Library
	Gui, Library:Add, Text, x12 y50 w380 h20 , Library:
	Gui, Library:Add, DropDownList, x12 y80 w380 vLibrary, Loch Raven||Parkville|Cockeysville
	Gui, Library:Add, Button, x12 y110 w90 h30 Default gLibraryButtonOK, &OK
	Gui, Library:Add, Button, x157 y110 w90 h30 Disabled gLibraryButtonDone , &Done
	Gui, Library:Add, Button, x302 y110 w90 h30 gLibraryButtonCancel, &Cancel
	Gui, Library:Show, x127 y87 h160 w404, Pick a Library
	Return
	
	LibraryGuiClose:
	LibraryGuiEscape:
	LibraryButtonCancel:
	LibraryButtonDone:
		Gui, Library:Destroy
		ExitApp
	
	LibraryButtonOK:
		Gui, Library:Submit
		Gui, Library:Destroy
		StringReplace, Library, Library, %A_SPACE%, , All
		_Library = %Library%


;*** Files
;* Settings
settings = Z:\settings.ini

;********************************
;
; # Windows Key
; ! Alt
; ^ Control
; + Shift
;
;********************************

;********************************
;
; Fuctions
;
;********************************

controlDirectReader(a=0)
{
	if (a = 0)
	{
		b=OFF
	}
	else if (a = 1)
	{
		b=ON
	}

	IfWinExist, CircControl-DirectReader
	{
		if (a = 2)
		{
			WinActivate, CircControl-DirectReader
			Sleep, 250
			Send !x
			;return
		}
		else
		{
			; Try to control DirectReader up to 5 times with a 100ms delay between tries.
			; Sometimes it doesn't change on the first try.
			Loop
			{
				if (a_index > 5)
				{
					StringLower, b_lower, b
					MsgBox, Couldn't turn %b_lower% Directreader.  Please turn it off manually.
					break  ; Terminate the loop
				}
				ControlGet, checkDirectReader, Enabled, , %b%, CircControl-DirectReader
				if (checkDirectReader = 1)
				{
					SetControlDelay -1
					ControlClick, %b%, CircControl-DirectReader,,,, NA
					Sleep, 100
				}
				else
				{
					break
				}
			}
			SetControlDelay, 40
		}
	}
}

breakScript(type = 0)
{
	if (type = npu)
		title = Item Records - Barcode Find Tool
	else if (type = missing)
		title = Request Manager - Hold Requests
	else if (type if integer)
		title = ahk_id %active_id%

	IfWinExist, %title%
	{
		WinClose
	}
	
	controlDirectReader(1)
	Progress, Off
	BlockInput, Off
	Gui, Holds:Destroy
	Gui, HldShlf:Destroy
	Gui, PrblmMtrl:Destroy
	Exit
}

log(text)
{
	FileAppend, %A_Now% - %text%`n, Z:\log.txt
}

;*******
;*
;* Usage
;* startTime :=db_startTime()
;* db_endTime(startTime)
;*
;*******
db_startTime()
{
	StartTime := A_TickCount
	
	return %StartTime%
}
db_endTime(StartTime)
{
	EndTime := A_TickCount
	ElapsedTime := (EndTime - StartTime)/1000
	MsgBox, Start Time: %StartTime%`nEnd Time: %EndTime%`n`nTotal Time: %ElapsedTime% seconds
}

;********************************
;
; Check Item out to the Hold Shelf
; Control + Alt + h
;
;********************************
;FileRead, HSCN,
IniRead, HSCN, %settings%, %_Library%, HoldShelfCardNumber, %A_Space%
if (HSCN="")
{
	;********************************
	; Hold Shelf GUI
	;********************************
	Gui, HldCrd:Add, Text, x12 y10 w380 h30 , Hold Shelf Card Number
	Gui, HldCrd:Add, Text, x12 y50 w380 h20 , Card Number:
	Gui, HldCrd:Add, Edit, x12 y80 w380 h20 Limit14 vCardNumber Number
	Gui, HldCrd:Add, Button, x12 y110 w90 h30 Default gHldCrdButtonOK, &OK
	Gui, HldCrd:Add, Button, x157 y110 w90 h30 Disabled gHldCrdButtonDone , &Done
	Gui, HldCrd:Add, Button, x302 y110 w90 h30 gHldCrdButtonCancel, &Cancel
	Gui, HldCrd:Show, x127 y87 h160 w404, Hold Shelf Card Number
	Return
	
	HldCrdButtonOK:
		Gui, HldCrd:Submit
		Gui, HldCrd:Destroy
		;IniWrite, Value, Filename, Section, Key 
		IniWrite, %CardNumber%, %settings%, %_Library%, HoldShelfCardNumber
		HSCN = %CardNumber%
	HldCrdGuiClose:
	HldCrdGuiEscape:
	HldCrdButtonCancel:
	HldCrdButtonDone:
		Gui, HldCrd:Destroy
		breakScript()
}

^!h::
{
	ControlDirectReader(0)
	
	MsgBox, 1 , Check Item Out to the Hold Shelf, Automatically check item out to the hold shelf.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	
	HoldShelfLabel:
	;********************************
	; Hold Shelf GUI
	;********************************
	Gui, HldShlf:Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically check the item out to the hold shelf.
	Gui, HldShlf:Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, HldShlf:Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, HldShlf:Add, Button, x12 y110 w90 h30 Default gButtonOK, &OK
	Gui, HldShlf:Add, Button, x157 y110 w90 h30 gButtonDone , &Done
	Gui, HldShlf:Add, Button, x302 y110 w90 h30 gButtonCancel, &Cancel
	Gui, HldShlf:Show, x127 y87 h160 w404, Problem Shelf
	Return
	

	ButtonOK:
		; Check out to problem material and return for more.
		Gui, HldShlf:Submit  ; Save the input from the user to each control's associated variable.
		Gui, HldShlf:Destroy
		
		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183) {
			BlockInput, On

			WinWait, Check Out, ,1
			if ErrorLevel
			{
				ErrorLevel = 0
				WinActivate, Polaris ILS
				Sleep 50
				Send {F3}
				WinGet, active_id, ID, Check Out - 
				;Exit
				;MsgBox, Check Out window isn't open.
			}
			
			Sleep, 150
			
			SetTitleMatchMode, Slow
			IfWinNotExist, Check Out, %HSCN%
			{
				SetTitleMatchMode, Fast

				WinActivate
				ControlFocus, Edit1, Check Out
				SetKeyDelay, 0
				Send %HSCN%{Enter}
				SetKeyDelay, %keydelay%
				Sleep, 300
				IfWinActive, Patron Blocks
				{
					Send y
				}
				SetKeyDelay, 0
				Send %ItemBarcode%{Enter}
				SetKeyDelay, %keydelay%
			}
			else
			{
				SetTitleMatchMode, Fast
				
				WinActivate
				ControlFocus, Edit3, Check Out
				SetKeyDelay, 0
				Send %ItemBarcode%{Enter}
				SetKeyDelay, %keydelay%
			}
		}
		else
		{
			textToLog = Input (%ItemBarcode%) doesn't look correct.  Please try again. ItemBarcodeLength: %ItemBarcodeLength% (Should be 14) ItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)
			log(textToLog)
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
		}
		
		Sleep, 100
		SetTitleMatchMode, Slow
		IfWinExist, Polaris, Do you want to renew?
		{
			Send n
			MsgBox, Item already checked out to the hold shelf.
		}
		SetTitleMatchMode, Fast
		
		
		BlockInput, Off
		Goto, HoldShelfLabel

	GuiClose:
	GuiEscape:
	ButtonDone:
	ButtonCancel:
		; Check if any items have been checked out to
		; the problem material card.  If any has, finish
		; the transaction, if not, just quit.
		Gui, HldShlf:Destroy
		;MsgBox Close/Done/Cancel.
		
		SetTitleMatchMode, Slow
		IfWinExist, Check Out, %HSCN%
		{
			ControlGet, OutputVar, List, , SysListView323, Check Out
			;MsgBox, %OutputVar%
			if (OutputVar != "")
			{
				WinActivate, CheckOut
				ControlFocus, Edit3, Check Out
				Send {Enter}
				;Polaris
				;ahk_class #32770
				WinWaitActive, Polaris, Do you want to print a check-out receipt?, 3
				Send n
				Send !{F4}
			}
		}
		SetTitleMatchMode, Fast
		
		breakScript()

}
