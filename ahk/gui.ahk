Gui, Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically check the item out to the problem shelf.
Gui, Add, Text, x12 y50 w380 h20 , Item Barcode:
Gui, Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number , 31183123456789
Gui, Add, Button, x12 y110 w90 h30 Default gButtonOK, &OK
Gui, Add, Button, x157 y110 w90 h30 gButtonDone , &Done
Gui, Add, Button, x302 y110 w90 h30 gButtonCancel, &Cancel

Gui, Show, x127 y87 h160 w404, Problem Shelf
Return

GuiClose:
ExitApp
ButtonOK:
Gui, Submit  ; Save the input from the user to each control's associated variable.
MsgBox You entered "%ItemBarcode%".
ExitApp
ButtonDone:
MsgBox Done.
ExitApp
ButtonCancel:
MsgBox Cancel.
ExitApp