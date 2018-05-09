#Persistent
#IfWinActive, ahk_class XLMAIN
#SingleInstance Force

global xl
while True {
	try {
		xl := ComObjActive("Excel.Application")   ; ID active Excel Application
		ComObjConnect(xl, "xl_")	; connect Excel events to corresponding script functions with the prefix "xl_".
		Break
	} catch {
		Sleep, 10000
	}
}


global cellIndexOffset := 0
global cellHistory := Object()

global tabIndexOffset := 1
global tabHistory := Object()

; Register a function to be called on exit:
OnExit("ExitFunction")

ExitFunction(){ ;release the excel object we bound on startup
    objRelease(xl)
}


xl_SheetActivate() {	; event that fires when new sheet is activated

	global xl
	
	global tabIndexOffset
	global tabHistory		;Array which stores the tab indexes accessed in chronological order **NOTE: AHK uses 1-indexed arrays**
	global thisSheetIndex	;Current Sheet Index
	
	global cellIndexOffset
	global cellHistory
	

	try {

		If (tabIndexOffset = 1) { ;This routine is only fired if the new tab is NOT activated by the tab cycling routine
            
            thisSheetIndex := xl.ActiveSheet.Index

			;Prevent any repeat values in the tabHistory array
            loopCount := tabHistory._MaxIndex()
			Loop, %loopCount% {
                ;A_Index begins at 1 and loops though to maxIndex
				If (tabHistory[A_Index] = thisSheetIndex) { ;remove any history of the current tab
					tabHistory.RemoveAt(A_Index)
				}
			}
			
			;Add the current sheet index to the tabHistory array
			tabHistory.InsertAt(1, thisSheetIndex)
    
			
			;Reset Cell History	in between sheet index changes
			cellHistory := []
			cellIndexOffset := 0
			
		}
		
	}
}
return


;Ctrl + ` cycles through the previously activated sheets in chronological order
^`::

	global xl
	global tabIndexOffset
	global tabHistory
    
	global cellIndexOffset
	global cellHistory

    while (GetKeyState("LCtrl", "P")) {

        ;try {
            tabIndexOffset := tabIndexOffset + 1
            if (tabIndexOffset > tabHistory.MaxIndex()) {
                tabIndexOffset := 1
            } 
            selectTabIndex := tabHistory[tabIndexOffset]
            xl.ActiveWorkbook.Worksheets(selectTabIndex).Activate
            KeyWait, ``  ;Wait for an escaped back-tick to be released
        ;}
        while (GetKeyState("``", "P") = 0 and GetKeyState("LCtrl", "P")) {  ;This causes an infinite loop until the back-tick is pressed again as long as Left Control is constantly held down
            sleep, 50
        }
    }

    try {
       
        loopCount := tabHistory._MaxIndex() ;This is necesary because of the weird way AHK evaluates loops
        Loop, %loopCount% {
            ;A_Index begins at 1 and loops though to maxIndex
            If (tabHistory[A_Index] = thisSheetIndex) { ;remove any history of the current tab
                tabHistory.RemoveAt(A_Index)
            }
        }

        ;Add the current tab index to the tabHistory array 
        thisSheetIndex := xl.ActiveSheet.Index
        tabHistory.InsertAt(1, thisSheetIndex)
        
        ;Reset Cell History	
        cellHistory := []
        cellIndexOffset := 0
            
        }


return


xl_SheetSelectionChange() { ;event that fires when a new sheet is selected
	
	global xl
	global cellHistory
	global thisCellAddress
	global cellIndexOffset
	
	If (cellIndexOffset = 0) {
	
		cellHistory.Insert(thisCellAddress)

		thisCellAddress := xl.Selection.Address
		
		loopCount := cellHistory._MaxIndex()
		Loop, %loopCount% {
		
			If (cellHistory[A_Index] = thisCellAddress) {
				
				cellHistory.RemoveAt(A_Index)
			}
		}
		
	}
	
}
return


;Left Alt + ` cycles through the history of selected cells in chronological order
<!`::

	global xl
	global cellIndexOffset
	global cellHistory

	try {
		cellIndexOffset := cellIndexOffset + 1

		cellIndex := cellHistory._MaxIndex() - cellIndexOffset + 1

		selectCellAddress := cellHistory[cellIndex]

		xl.ActiveSheet.Range(selectCellAddress).Select
	}
return


^F1:: ;Just a little debugging
	global tabHistory
	loopCount := tabHistory._MaxIndex()
	msgText :=""
	Loop %loopCount% {
		msgText := msgText . `A_Index . ":" . tabHistory[A_Index] . "  "
	}
msgBox, %msgText%
msgText2 := tabHistory[0]
msgBox, %msgText2%
return

Exit, objRelease(xl)