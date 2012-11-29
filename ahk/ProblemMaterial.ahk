#NoEnv
#SingleInstance force
keyDelay = 25
controlDelay = 40
OutputVar = ""
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
ControlDirectReader(a=0)
{
	if(a=0)
		b=OFF
	else
		b=ON
	
	IfWinExist, CircControl-DirectReader
	{
		SetControlDelay -1
		ControlClick, %b%, CircControl-DirectReader
		SetControlDelay, %controlDelay%
	}
}

;********************************
;
; Check Item out to the Problem Shelf
; Control + Alt + p
;
;********************************

MsgBox, 1 , Chck Item Out to Problem Shelf, Automatically check item out to the problem shelf.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7

;FileRead, OutputVar, Filename
;FileRead, PMCN, problemmaterialcardnumber.txt

;LO
PMCN = 21183016338757
;PA
;PMCN = 21183001187771

^!p::
{
	ControlDirectReader(0)
	
	ProblemMaterialLabel:
	;********************************
	; Problem Material GUI
	;********************************
	Gui, Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically check the item out to the problem shelf.
	Gui, Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, Add, Button, x12 y110 w90 h30 Default gButtonOK, &OK
	Gui, Add, Button, x157 y110 w90 h30 gButtonDone , &Done
	Gui, Add, Button, x302 y110 w90 h30 gButtonCancel, &Cancel
	Gui, Show, x127 y87 h160 w404, Problem Shelf
	Return
	

	ButtonOK:
		; Check out to problem material and return for more.
		Gui, Submit  ; Save the input from the user to each control's associated variable.
		Gui, Destroy
		;MsgBox You entered "%ItemBarcode%".
		
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
				;Exit
				;MsgBox, Check Out window isn't open.
			}
			
			Sleep, 150
			
			SetTitleMatchMode, Slow
			IfWinNotExist, Check Out, %PMCN%
			{
				SetTitleMatchMode, Fast

				WinActivate
				ControlFocus, Edit1, Check Out
				SetKeyDelay, 0
				Send %PMCN%{Enter}
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
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
		}
		
		Sleep, 100
		SetTitleMatchMode, Slow
		IfWinExist, Polaris, Do you want to renew?
		{
			Send n
		}
		SetTitleMatchMode, Fast
		
		
		BlockInput, Off
		Goto, ProblemMaterialLabel

	GuiClose:
	GuiEscape:
	ButtonDone:
	ButtonCancel:
		; Check if any items have been checked out to
		; the problem material card.  If any has, finish
		; the transaction, if not, just quit.
		Gui, Destroy
		;MsgBox Close/Done/Cancel.
		
		SetTitleMatchMode, Slow
		IfWinExist, Check Out, %PMCN%
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
			}
		}
		SetTitleMatchMode, Fast
		
		Goto, BreakScript

}
BreakScript:
	ControlDirectReader(1)
	Exit