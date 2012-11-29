#NoEnv
#SingleInstance force
keyDelay = 25
i = 0
OutputVar = ""
ItemRecordIterations = 10
controlDelay = 40
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%

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
;********
;
; ControlDirectReader
; 0 for OFF, 1 for ON
;
;******

controlDirectReader(a=0)
{
	if(a=0)
		b=OFF
	else
		b=ON

	IfWinExist, CircControl-DirectReader
	{
		; Try to control DirectReader up to 5 times with a 100ms delay between tries.
		; Sometimes it doesn't turn change on the first try.
		Loop
		{
			if a_index > 5
			{
				StringLower, b_lower, b
				MsgBox, Couldn't turn %b_lower% Directreader.  Please turn it off manually.
				break  ; Terminate the loop
			}
			ControlGet, directReaderOn, Enabled, , %b%, CircControl-DirectReader
			if (directReaderOn = 1)
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
	Exit
}

;********************************
;
; Request Manager - Missing
; Control + Alt + m
;
;********************************

^!m::
{
	MsgBox, 1 , Missing Request Manager Items, Mark Request Manager items as missing.`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	IfMsgBox Cancel
	{
		Exit
	}

	WinWait, Request Manager - Hold Requests, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		MsgBox, Request Manager window isn't open.
		WinActivate, Polaris ILS
		Sleep 50
		Send {Alt}rr
		Sleep, 100
		;Exit
	}

	BlockInput, On
	
	ControlGet, OutputVar, List, Col9 , SysListView321, Request Manager - Hold Requests
	
	if (OutputVar == "")
	{
		MsgBox, There are no items to be marked as missing.
		breakScript()
	}
	
	MsgBox, 260, Mark items as missing?, The following barcodes will be marked as missing.  Do you wish to continue?`n`n%OutputVar%
	IfMsgBox No
	{
		Exit
	}
	
	MissingLabel:
	Loop, Parse, OutputVar, `n  ; Rows are delimited by linefeeds (`n).
	{
	
		ItemBarcode = %A_LoopField%
		; If Item Records window isn't open, open it
		WinWait, Item Records - Barcode Find Tool, ,0
		if ErrorLevel
		{
			ErrorLevel = 0
			WinActivate, Polaris ILS
			Sleep 50
			Send ^{F9}
		}

		; Activate Item Records window and focus on the barcode input box
		; @todo: make sure Item Records is expecting item barcode
		WinActivate, Item Records - Barcode Find Tool
		ControlFocus, Edit3, Item Records - Barcode Find Tool

		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183)
		{
			BlockInput, On
			ControlFocus, Edit3, Item Records - Barcode Find Tool
			SetkeyDelay, 0
			Send %ItemBarcode%{Enter}
			SetkeyDelay, %keyDelay%
			; Waits x times every 250ms (one quarter second)
			; ItemRecordIterations = 10 (at top of script)
			; until the item gets returned by Polaris
			retry:
			ControlGet, OutputVar, List, Selected Col9, SysListView321, Item Records
				if (ErrorLevel or OutputVar = "")
				{
					if (i < ItemRecordIterations)
					{
						i++
						Sleep, 250
						Goto, retry
					}
					else
					{
						MsgBox Error Level: %ErrorLevel%`r`nOutput Var: %OutputVar%`r`ni: %i%
						breakScript()
					}
				}
				if (%ItemBarcode% != %OutputVar%)
				{
					MsgBox, Barcode doesn't match or item doesn't exist.
					Goto, MissingLabel
				}
			Sleep, 100
			
			Send {Enter}
			
			WinWaitActive, Item Record, Barcode:
			Sleep, 100
			ControlFocus, ComboBox5, Item Record, Barcode:
			
			Send m
			;Send ^s ;Save changed information
			;Send !{F4} ;Close window
		}
		;MsgBox, %ItemBarcode%
		
		WinWaitClose, Item Record, Barcode:
	}
	;MsgBox, %OutputVar%
}