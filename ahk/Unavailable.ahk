#NoEnv
#SingleInstance force
keydelay = 25
SetKeyDelay, %keydelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, 40

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
; Mark Item as Unavailable
; Control + Alt + u
;
;********************************



^!u::
{
	MsgBox, 1 , Mark Item as Unavailable, Automatically mark item as unavailable.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	IfMsgBox Cancel
	{
		Exit
	}

	Unavailable:
	
	WinWait, Item Records - Barcode Find Tool, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		;MsgBox, Item Records - Barcode Find Tool window isn't open.
		WinActivate, Polaris ILS
		Sleep 50
		Send ^{F9}
		;Exit
	}
	
	WinActivate, Item Records - Barcode Find Tool
	ControlFocus, Edit3, Item Records - Barcode Find Tool
	
	InputBox, ItemBarcode, Item Barcode, Scan the item barcode here.  It will automatically mark the item as 

	if ErrorLevel
	{
		ErrorLevel = 0
		Exit
	}
	else
	{
		;StartTime := A_TickCount
		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183) {
			BlockInput, On
			ControlFocus, Edit3, Item Records - Barcode Find Tool
			SetKeyDelay, 0
			Send %ItemBarcode%{Enter}
			SetTitleMatchMode, Slow
			SetKeyDelay, %keydelay%
			Sleep, 150
			WinwaitActive, , %ItemBarcode%, 2
			;MsgBox, True
			;Exit
			Send {Enter}
			Sleep, 150
			SetTitleMatchMode, Fast
			WinWaitActive, Item Record, ,1 
			ControlFocus, ComboBox5, Item Record
			Send u^s!{F4}
			BlockInput, Off
			;EndTime := A_TickCount
			;ElapsedTime := EndTime - StartTime
			;MsgBox, Start Time: %StartTime%`nEnd Time: %EndTime%`n`nTotal Time: %ElapsedTime%
			;Goto, NPULabel
			;Exit
		}
		else
		{
			MsgBox, Input (%ItemBarcode%) doesn't look correct.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
			Exit
		}
	}
	Exit
}