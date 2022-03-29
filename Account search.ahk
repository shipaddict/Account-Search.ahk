#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Event
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 30
SetTitleMatchMode, 2 
#Persistent 
#SingleInstance force 
AutoTrim On

iniFile = C:\Users\%A_UserName%\AHK\In Use\Accounts.ini
IniRead, icon, %iniFile%, settings, icon
IniRead, affirms, %iniFile%, settings, affirms
IniRead, lvw, %iniFile%, settings, lvw

Menu, Tray, Icon, C:\Users\%A_UserName%\Documents\Scripting\AutoHotKey\icons\contacts blue.png
Menu, Tray, Tip, Account Search
Menu, Tray, Add, Show Account Search, ShowList 
TrayTip, Account Search

Row := 0 
datenow := FormatTime, datenow, %A_Now%, ddMMMyyyy

!l:: ;refocus in the ListView area   
	GuiControl, Focus, MyListView
Return 

^+\::
	Gui, Font, Calibri s11
	Gui, +Resize +Border 
	Gui, Add, ListView, section x15 w1540 r15 LV0x4500 Sort BackgroundAqua AltSubmit gMyListView vMyListView, Account Name | Acct# | Cust ID | Contact | Email | Invoice Email | Phone | Terms | Reg Code | bhive | Notes 

	Gui, Add, GroupBox, x15 w1540 h350 section, Add/Edit Entry
		Gui, Add, Text, xs+30  yp+75, Customer &Name
		Gui, Add, Edit, xp+120 yp+m w400 vacctName
		Gui, Add, Text, xp+445 yp+m , Account N&umber
		Gui, Add, Edit, x+m yp+m w130 vacctNum
		Gui, Add, Text, x+m yp+m, Customer &ID
		Gui, Add, Edit, x+m yp+m w130 vcustID
		
		Gui, Add, Text, xs+30 yp+35, C&ontact Person(s)
		Gui, Add, Edit, xp+120 yp+m w400 vcontact
		Gui, Add, Text, xp+445 yp+m , E&mail(s)
		Gui, Add, Edit, x+m yp+m w410 vemail
		
		Gui, Add, Text, xs+30 yp+35, &Phone 
		Gui, Add, Edit, xp+120 yp+m w400 vphone
		Gui, Add, Text, xp+445 yp+m , In&voice Email(s)
		Gui, Add, Edit, x+m yp+m w365 vinvEmail

		Gui, Add, Text, xs+30 yp+35 w125, Payment &terms
		Gui, Add, Edit, xp+120 yp+0 w400 vterms
		Gui, Add, Text, xp+445 yp+m , &bhive
		Gui, Add, Edit, x+m yp+m w130 vbhive 
		Gui, Add, Text, x+m yp+m , &Reg Code   
		Gui, Add, Edit, x+m yp+m w130 vregCode 
		
		Gui, Add, GroupBox, xs+1130 y355 w400 h210 section, Account Notes&. 
		Gui, Add, Edit, xp+7 yp+18 r11 w385 vnotes
		
		Gui, Add, GroupBox, xp-5 yp+200 w400 h100 section, Old Stuff&, 
		Gui, Add, Edit, xp+7 yp+18 r4 w385 vold

		Gui, Add, Button, x190 yp+35 w100 Default, &Assign 
		Gui, Add, Button, x+m yp w100, &Clear
		Gui, Add, Button, x+m yp w100, &Delete
		Gui, Add, Button, x+m yp w100, &Edit
		Gui, Add, Button, x+m yp w100, Ne&w
		Gui, Add, Button, x+m yp w100, &Save
		Gui, Add, Button, x+m yp w100, &Quit.

			Gui, Add, StatusBar,, 
			GoSub, statbar 

	Loop, Read, %iniFile%
		{
			searchterm := StrSplit(A_LoopReadLine , "=")
			If searchterm[1] = "[Accounts]"
				Continue
			If searchterm[1] = "[settings]"
				Break

			key := searchterm[1]
			customer := StrSplit(searchterm[2] , "|")
			acctName := trim(customer[1])
			acctNum := trim(customer[2])
			custID := trim(customer[3])
			contact := trim(customer[4])
			email := trim(customer[5])
			invEmail := trim(customer[6])
			phone := trim(customer[7])
			terms := trim(customer[8])
			regCode := trim(customer[9])
			bhive := trim(customer[10])
			notes := trim(customer[11])
			oldStuff := trim(customer[12])
			LV_Add("", acctName,acctNum,custID,contact,email,invEmail,phone,terms,regCode,bhive,notes,oldStuff)  
		}
			LV_ModifyCol(1,"450 Text" "Left" "Sort") ;acctName
			LV_ModifyCol(2,"Auto Integer" "Center")  ;acctNum
			LV_ModifyCol(3,"80 Integer" "Center")  ;custID
			LV_ModifyCol(4,"150 Text" "Left")  ;contact 
			LV_ModifyCol(5,"200 Text" "Left")  ;email
			LV_ModifyCol(6,"150 Text" "Left") ;invEmail
			LV_ModifyCol(7,"100 Text" "Left")  ;phone
			LV_ModifyCol(8,"75 Text" "Left") ;terms
			LV_ModifyCol(9,"80 Text" "Center")  ;regCode
			LV_ModifyCol(10,"80 Text" "Left")  ;bhive
			LV_ModifyCol(11,"250 Text" "Left")  ;Notes 
			
	Gui, Show, AutoSize Center, Account Search -- Premier Accounts 

	Menu, MyContextMenu, Add, &Assign, Assign ;assign all values in row to hotkeys 
	Menu, MyContextMenu, Add, &Clear Edit Area, Clear
	Menu, MyContextMenu, Add, &Delete Record, Delete
	Menu, MyContextMenu, Add, &Edit, Edit
	Menu, MyContextMenu, Default, &Edit  ; Make "Edit" a bold font to indicate that double-click does the same thing.
	Menu, MyContextMenu, Add, &Primary Email (Send New), MailToPrimary
	Menu, Tray, Add, Show Account Search, ShowList 

	Menu, FileMenu, Add, &Open Customer Folder, GoToFolder
	Menu, FileMenu, Add, &Save in/Create Customer Folder, SaveToFolder
	Menu, FileMenu, Add, Open &App Folder, OpenAppFolder
	Menu, FileMenu, Add, Open &Backups Folder, OpenBackupsFolder
	Menu, FileMenu, Add, &Close Window, Quit
	Menu, FileMenu, Add, E&xit Ctrl+X, Exit

	Menu, CustMenu, Add, &Assign values to Hotkeys`tCtrl+Alt+A, Assign
	Menu, CustMenu, Add, &Clear Add/Edit Area, Clear
	Menu, CustMenu, Add, &Delete current entry, Delete
	Menu, CustMenu, Add, &Edit Customer Data, Edit
	Menu, CustMenu, Add, &Save Customer Data, Save 

	Menu, SearchMenu, Add, Look up by bhive # in bhive, lookupbhive
	Menu, SearchMenu, Add, Look up by custID in rev.io, srchRevio 
	Menu, SearchMenu, Add, Look up by acctNum in Outlook, srchOutlook
	Menu, SearchMenu, Add, Look up by acctName in Salesforce, srchSalesforce	
	Menu, SearchMenu, Add, Look up by acctNum in Aging, srchAging
	
	;separate bhive menu?  
	Menu, SearchMenu, Add, Download all invoices on page in bhive, dlInvsbhive
	; Menu, SearchMenu, Add, Look up payments in bhive, 
	; Menu, SearchMenu, Add, View 2022 invoices in bhive 
	; Menu, SearchMenu, Add, View 2021 invoices in bhive 
	; Menu, SearchMenu, Add, View 2020 invoices in bhive 
	; Menu, SearchMenu, Add, View 2019 invoices in bhive 
	; Menu, SearchMenu, Add, View 2018 invoices in bhive 
	; Menu, SearchMenu, Add, View 2017 invoices in bhive 
	
	
	; Menu, PasteMenu, Add, &Download Folder Path (on PC), PasteDL
	; Menu, PasteMenu, Add, W-&9 Document Path, PasteW9
	; Menu, PasteMenu, Add, &Remit Instructions Path, PasteRemit \
	
	;search menu 
	
	 
	Menu, SendMenu, Add, Email &Primary Contact `tWin+Alt+NumPad3, MailToPrimary 
	Menu, SendMenu, Add, Select & Email Primary Contact `tNumPadSub+NumPad3, MailToSelPrimary
	Menu, SendMenu, Add, Send to Invoice Email address `tWin+Alt+NumPad6, MailToInvoice
	Menu, SendMenu, Add, Select & Email Invoice Email Address `tNumPadSub+NumPad6, MailToSelInv
	
	Menu, HelpMenu, Add, &Preferences, Assign

	Menu, MyMenuBar, Add, &File, :FileMenu 
	Menu, MyMenuBar, Add, &Customer, :CustMenu
	Menu, MyMenuBar, Add, &Search, :SearchMenu
	Menu, MyMenuBar, Add, &Send, :SendMenu 
	Menu, MyMenuBar, Add, &Help, :HelpMenu
	Gui, Menu, MyMenuBar
	GoSub, Backups
Return 

MyListView:
	If (A_GuiEvent = "DoubleClick") {
			Row := LV_GetNext(0, "F")  ; Find the focused row.
			If not Row  ; No row is focused.
				Return
			LV_GetText(acctName, Row, 1) 
			LV_GetText(acctNum, Row, 2) 
		If (A_GuiEvent = "Normal")
        || (A_GuiEvent = "K") {
            LV_GetText(fileBack, A_EventInfo, 1)
        }
    	GoSub, codes
	}
	If A_GuiEvent = ColClick
		{
		Row := LV_GetNext()
		LV_Modify(Row, "Vis")
		}
Return 

GuiContextMenu:
	{
		If (A_GuiControl != "MyListView")  ;  displays the menu only for clicks inside the ListView.
			Return
		Menu, MyContextMenu, Show, %A_GuiX%, %A_GuiY%
	}
Return

ShowList: 
	Gui, Show, , Account Search -- Premier Accounts 
Return 

AddItem: 
ButtonNew:
Edit:
ButtonEdit:
	Gui, Submit, NoHide
	If (Row = 0) 
	{
		LV_GetText(lastPayAmt, Row, 15) 
		LV_Add("Focus Select",trim(acctName),trim(acctNum),trim(custID),trim(contact),trim(email),trim(invEmail),trim(phone),trim(terms),trim(regCode),trim(bhive),trim(notes),trim(oldStuff))  
	} else {
		LV_GetText(oldacctName, Row, 1) 
		LV_Modify(Row,,trim(acctName),trim(acctNum),trim(custID),trim(contact),trim(email),trim(invEmail),trim(phone),trim(terms),trim(regCode),trim(bhive),trim(notes),trim(oldStuff))  	
		
		If (oldacctName != acctName) 
		{
			newPath := % StrReplace(fullPath, oldacctName, acctName)
			FileMoveDir, %fullPath%, %newPath%, R
		}
	}
	GoSub, Save
	Gui, Show
	Sleep 100
	GuiControl, Focus, MyListView
Return 

Assign:
ButtonAssign:
	Gui, Submit 
	Row := LV_GetNext(0, "F")  ; Find the focused row.
	If not Row { ; No row is focused.
		MsgBox Please choose a customer 
		} else {
	LV_GetText(acctName, Row, 1) 
	LV_GetText(acctNum, Row, 2)  
	GoSub, codes
	Gui, Show
		}
Return 

ButtonClear: 
Clear: 
	Row := LV_GetNext(0, "F")  ; Find the focused row.
	LV_Modify(Row, "-Focus") 
	GuiControl, , acctName
	GuiControl, , acctNum
	GuiControl, , custID
	GuiControl, , contact
	GuiControl, , email
	GuiControl, , invEmail 
	GuiControl, , phone
	GuiControl, , terms
	GuiControl, , regCode
	GuiControl, , bhive
	GuiControl, , notes
	GuiControl, , oldStuff
Return 

ButtonDelete:
Delete: 
	Row := LV_GetNext(0, "F")  ; Find the focused row.
	If not Row { ; No row is focused.
			MsgBox Please choose a customer to delete 
		} else {
	LV_GetText(acctName, Row, 1) 
	LV_GetText(acctNum, Row, 2)  
		Msgbox, 308, Hey!!, Are you sure you want to `nPERMANENTLY delete %acctName% %acctNum%?   
		IfMsgBox No
			Return  
		IfMsgBox Yes 
			LV_Delete(Row)
			key = %acctName%%acctNum%
			IniDelete, %iniFile%, Accounts, %key%
		}
Return 

ButtonSave:
Save: 
	Gui, Submit, NoHide
	Gui, Default
	key = %acctName%%acctNum%
	FormatTime, timeOfEdit,, MM/d/yy h:mm tt
	newVal := acctName . "|" . acctNum . "|" . custID . "|" . contact . "|" . email . "|" . invEmail . "|" . phone . "|" . terms . "|" . regCode . "|" . bhive . "|" . notes . "|" . oldStuff . "*** updated " . timeOfEdit
	IniWrite, %newVal%, %iniFile%, Accounts, %key%
	FileAppend, %newVal%`n`n, *C:\Users\%A_UserName%\Documents\Backup\ini File Edits log.txt
	GoSub, codes 
	GuiControl, Focus, MyListView
Return 

#!NumPad3:: 
MailToPrimary:
{
	Row:= LV_GetNext(0, "F")  ; Find the focused row.
	If not Row { ; No row is focused.
			MsgBox Please select a customer
		} else { 
	GoSub, codes 
	If not email {
		MsgBox There is no email address for the primary contact
	} else {
		emails := % email 
		facctName := facctName(acctName) 
		subj1 := % "Broadvoice Communications ~~ " . acctNum . " ~ " . facctName
		}
	SetCapsLockState, Off
	newOutlookEmail(emails, subj1) 
	}
}
Return 

#^NumPad3::
SelPrimaryEMail:
{
	Row:= LV_GetNext(0, "F")  ; Find the focused row.
	If not Row { ; No row is focused.
			MsgBox Please select a customer
		} else { 
	GoSub, codes 
	If not email {
			MsgBox There is no email address for the primary contact
		} else {
	email := RegExReplace(email, ",\s*|;\s*|\s{2,}", ";")
	
	Loop, Parse, email, ";"
		Gui, 2: Add, Button, gLabel, &%A_Index% %A_LoopField% 
		Gui, 2: Show 
	Return 
		}
	}
}
Return 

NumPadSub & NumPad3::
MailToSelPrimary:
{
	GoSub, SelPrimaryEMail 
	subj1 := % "Broadvoice Communications ~~ " . acctNum . " ~ " . facctName
	mail = yes 
	; Sleep 1000 
	; SetCapsLockState, Off
	GoSub, NumPadSub 
}
Return 

NumPadSub: 
{		
	NumPadSub::Send -
} 
Return 

; #!3:: ;get most recent invoice number from customer's folder 
getMostRecentInvoice:
{
	BlockInput On 
	SetCapsLockState Off 
	tmpCB := ClipBoard
	ClipBoard := ""
	StringLower, fullPath, fullPath
	invFold = %fullPath%invoices 		
	Loop, Files, %invFold%\*.pdf, F
	{
		FileList .= SubStr(A_LoopFileName,1
			, ( StrLen(A_LoopFileName) - StrLen(A_LoopFileExt) -1 ) ) "`n"
	}
	Sort, FileList, R
	SelectedFile := SubStr( FileList, 1, InStr( FileList, Chr(10))-1 )
	SelectedFile := StrReplace(SelectedFile, ".pdf", "")
	MsgBox % SelectedFile
	ClipBoard := tmpCB
	BlockInput Off
}
Return 

#!NumPad6::
MailToInvoice:  ;;;;;;;; new email to invoice email address. attach any new invoices created today 
{
	fileList := 
	Row:= LV_GetNext(0, "F")  ; Find the focused row.
	If not Row { ; No row is focused.
		MsgBox Please select a customer 
	} else {
		GoSub, codes 
		If not invEmail {
			MsgBox There is no Invoice-only email address 
		} else {
			emails = % invEmail 
			StringLower, fullPath, fullPath
			invFold = %fullPath%Invoices\
		}
		fileList := getTodaysInvs(invFold)
		}
		subj1 := % entity . " Invoice Submission " . acctNum . " ~ " . facctName
		newOutlookEmail(emails, subj1) 
		Sleep 1000
		WinWaitActive, - Message,,5
		If not ErrorLevel
		{
			If (fileList != "") 
			{
				Send !2b
				WinWaitActive, Insert File
				Send !n
				Send %invFold%{Enter}
				Sleep 100
				Send %fileList%{Enter}
			}
		}
}
Return 

getTodaysInvs(fold)
{
	Loop, Files, %fold%\*.pdf, F
	{
		Horas := A_Now 
		EnvSub, Horas, %A_LoopFileTimeCreated%, days 
		If (Horas < 1) 
		{
			fileList .= """" A_LoopFileName """"
		}
	}
	Return % fileList
}


#^NumPad6::
SelInvEMail:
{
	Row:= LV_GetNext(0, "F")  ; Find the focused row.
	If not Row ; No row is focused.
	{ 
		MsgBox Please select a customer
	} else { 
		GoSub, codes 
		If not invEmail 
		{
			MsgBox There is no email address for the Invoice-only contact
		} else {
			invEmail := RegExReplace(invEmail, ",\s*|;\s*|\s{2,}", ";")
			Loop, Parse, invEmail, ";"
			{
				Gui, 2: Add, Button, gLabel, &%A_Index% %A_LoopField% 
				Gui, 2: Show 
				Return 
			}
		}
	Send !{Tab}
	}
}
Return 

NumPadSub & NumPad6::
MailToSelInv:
{
	GoSub, SelInvEMail
	StringLower, fullPath, fullPath
	invFold = %fullPath%Invoices\
	fileList := getTodaysInvs(invFold)
	subj1 := % entity . " Invoice Submission " . acctNum . " ~ " . facctName
	mail = yes 
	GoSub, NumPadSub 
	WinWaitActive, - Message,,5
	If not ErrorLevel
	{
		If (fileList != "") 
		{
			Send !2b
			WinWaitActive, Insert File
			Send !n
			Send %invFold%{Enter}
			Sleep 100
			Send %fileList%{Enter}
		}
	}
}
Return 

Label: 
{
	SetCapsLockState, Off
	Gui, 2: Submit, NoHide 
	email = %A_GuiControl% 
	email := StrSplit(email, A_Space)
	ClipBoard = % email[2]
	email = % email[2]
	If (mail = "yes")
	{
		newOutlookEmail(email, subj1) 
		; GoSub, Mail 
	}
	Gui, 2: Destroy 
	; Send !{Tab}
}
Return 

OpenAppFolder:
	IniRead, appFolder, %iniFile%, settings, appFolder
	Run, % appFolder
Return 

OpenBackupsFolder:
	IniRead, bkups, %iniFile%, settings, bkups
	Run, % bkups
Return 

#NumPad0:: 
::acctNumx::
PasteAcctNum:
	Send % acctNum
Return

#NumPad1:: 
::acctNamex::
PasteAcctName:
	acctName := StrReplace(acctName, "#", "{#}")
	Send % acctName
Return

#NumPad2:: 
::custIDd::
PastecustID: 
	Send %custID%
Return 

#NumPad3:: 
::priEmail::
PastePrimaryEmail:
	email := RegExReplace(email, ",\s*|;\s*|\s{2,}", "; ")
	Send %email%
Return 

#NumPad4::
::contc:: 
PasteContact:
	Send %contact%
Return 

#NumPad5::  
SaveToFolder:
{
	Send, %fullPath%
	Gosub, IsFolderIs 
}
Return 

#!NumPad5:: ;send 'file-friendly' customer file path to customer invoice folder 
GoToFolder: 
	Gosub, IsFolderIs 
	fullPath := % fullPath . "Invoices\" 
	Sleep 250 
	Send % fullPath 
Return 

#NumPad6:: 
::phNum::
PastePhone:
	Send %phone%
Return 

#NumPad7::
::aterms::
PasteTerms:
	Send %terms%
Return 

#NumPad8::
::regCodex::
PasteregCode:
	Send %regCode%
Return 

#a:: ;search for customer in open Aging report from Account Search 
srchAging:
{
	Sleep 100 
	WinActivate, Aging 
	WinWaitActive, Aging,, 2 
	xlFindAll(custIDD)
}
Return 

#b::  ;search for customer in bhive from Account Search 
srchBhive:
{
	bhivelookup(bhive) 
}
Return 

#f::  ;search for customer in rev.io 
srchRevio: 
{
	revioLookUpCustiD(custID)
}
Return 

#o:: ;;search for customer number in Outlook from Account Search  
srchOutlook:
{
	OutlookSearch(acctNum)
}
Return 

#u:: ;;search for customer name in Salesforce from Account Search 
srchSalesforce: 
{
	Sleep 1000
	yesno := chkSalesforceWindow()
	If (yesno = "yes") 
	{
		; MsgBox % acctName
		Sleep 500 
	srchInSalesforce(acctName) 
	}
}
Return 

pasteDL:
	Send c:\users\%A_UserName%\downloads\
Return 

#NumPad9:: ;;paste bhive account number 
::bhv::
Pastebhive:
	Send %bhive%
Return

#!NumPad9:: ;;look up bhive account 
::gobhv::
lookupbhive:
	Run https://my.broadvoice.com/locations/%bhive%/
	Send %bhive%
Return

#+NumPad9:: ;;look up payments in bhive 
::paybhv::
paybhive: 
	Run https://my.broadvoice.com/locations/%bhive%/transactions
Return

#^NumPad9:: ;;look up 2020 invoices in bhive 
::invbhv::
invoicesbhive:
	Run https://my.broadvoice.com/locations/%bhive%/statements?year=2020
Return

::remt::
PasteRemit: 
	Send, %remitInfo%
Return 

+!m::
	emailFile := % "C:\Users\" . A_UserName . "\Documents\tempFolder\emailAddressTemp.txt" 
	If FileExist(emailFile) 
	{	
		FileRead, Contents, %emailFile%
		Contents := StrReplace(Contents, "`r`n", "`n") 
		Contents := trim(Contents)
		Loop, Parse, Contents, "`n"
			Gui, 3: Add, Button, gemailLabel, &%A_Index% %A_LoopField%
			Gui, 3: Show 
	}
Return 

emailLabel: 
	Gui, 3: Submit, NoHide 
	email = %A_GuiControl% 
	email := StrSplit(email, A_Space)
	ClipBoard = % email[2]
	Gui, 3: Destroy 
	Send !{Tab}
Return

sum(params*)
{
	for index, param in params
		if param is number
			result += param
	return (result) ? result : 0
}

#IfWinactive, Account Search -- Premier Accounts 
F6::
    GoSub, PopulateLV
return

#IfWinActive 
Backups: 
{
	fileDate := fileDate()
	FileCopy, %iniFile%, C:\Users\%A_UserName%\Documents\Backup\Accounts %fileDate%.ini  ;newest backup will NOT overwrite any previous ones from that same calendar day 
	FileList := 
	Loop Files, C:\Users\%A_UserName%\Documents\Backup\Accounts *.ini
		FileList .= A_LoopFileTimeModified "`t" A_LoopFileName "`n"
		Sort, FileList, R
	Loop, Parse, FileList, `n
	{
		If (A_LoopField = "")  ; Omit the last linefeed (blank item) at the end of the list.
			continue
		If A_Index > 10
		{
			fileName := SubStr(A_LoopField, 16)
			FileDelete, C:\Users\%A_UserName%\Documents\Backup\%fileName%
		}
	}
}
Return 

codes:
{
		RowNum := A_EventInfo 
	LV_GetText(acctName,Row, 1) 
	LV_GetText(acctNum,Row,2)
	LV_GetText(custID,Row,3)
	LV_GetText(contact,Row,4)
	LV_GetText(email,Row,5)
	LV_GetText(invEmail,Row,6)
	LV_GetText(phone,Row,7)
	LV_GetText(terms,Row,8)
	LV_GetText(regCode,Row,9)
	LV_GetText(bhive,Row,10)
	LV_GetText(notes,Row,11)
	LV_GetText(oldStuff,Row,11)
		GuiControl, Text, acctName, %acctName%
		GuiControl, Text, acctNum, %acctNum%
		GuiControl, Text, custID, %custID%
		GuiControl, Text, contact, %contact%
		GuiControl, Text, email, %email%
		GuiControl, Text, invEmail, %invEmail%
		GuiControl, Text, phone, %phone%
		GuiControl, Text, terms, %terms%
		GuiControl, Text, regCode, %regCode%
		GuiControl, Text, bhive, %bhive%
		GuiControl, Text, notes, %notes%
		GuiControl, Text, oldStuff, %oldStuff%
	
		facctName := facctName(acctName)
		IniRead, acctPath, %iniFile%, settings, acctPath
		IniRead, W9, %iniFile%, settings, gasW9
		IniRead, remitInfo, %iniFile%, settings, remitInfo
		StringLower, acctPath, acctPath 
		fullPath := % acctPath . facctName . " " . acctNum . "\"
	Gui, Show
	GuiControl, Focus, MyListView
}
Return 

statbar: 
{
	affirms = C:\Users\%A_UserName%\OneDrive - BROADVOICE\AHK\In Use\Affirmations.txt
	IniRead, x, %iniFile%, settings, aff
	FileReadLine, line, %affirms%, %x%
	SB_SetParts(50)
	SB_SetText(line,2) 
	x += 1
		If (x >= 61){
			x = 1
		}
	IniWrite, %x%, %iniFile%, settings, aff
}
Return 

^F3:: ;look up by account number in rev.io 
	InputBox, acctNum, Enter rev.io account number 
	GoSub, Search2
Return

!F3:: ;look up by bhive number in bhive 
	InputBox, bhive, Enter bhive number
	acctName := findacctNameFrombhive(bhive) 
	bhivelookup(bhive, acctName) 
Return 

+!F3:: ;look up by acctNum in bhive 
	InputBox, acctNum, Enter bhive number, custID number, or rev.io account number
	bhive := findbhiveFromacctNum(acctNum) 
	; MsgBox % bhive 
	acctName := findacctNameFrombhive(bhive) 
	bhivelookup(bhive, acctName) 
Return 

PopulateLV:
{
    LV_Delete()

	Loop, Read, %iniFile%
	{
		searchterm := StrSplit(A_LoopReadLine , "=")
		If searchterm[1] = "[Accounts]"
			Continue
		If searchterm[1] = "[settings]"
			Break
		key := searchterm[1]
		customer := StrSplit(searchterm[2] , "|")
		acctName := trim(customer[1])
		acctNum := trim(customer[2])
		custID := trim(customer[3])
		contact := trim(customer[4])
		email := trim(customer[5])
		invEmail := trim(customer[6])
		phone := trim(customer[7])
		terms := trim(customer[8])
		regCode := trim(customer[9])
		bhive := trim(customer[10])
		notes := trim(customer[11])
		oldStuff := trim(customer[11])
		LV_Add("", acctName,acctNum,custID,contact,email,invEmail,phone,terms,regCode,bhive,notes,oldStuff)  ;,terms,tax,regCode,lastPayDate,lastPayAmt,lastSaleDate,lastSaleAmt,pastDueAmt,lastUpdate
		Gui, Show 
		GuiControl, Focus, MyListView
    }
}
Return

#IfWinActive, But Her Emails
!v:: 
	acctName := RegExReplace(acctName, "`#", "{#}")
	Send % acctNum
	Send {Tab}
	Send % acctName  
	Send {Tab}%email%{Tab}%terms%{Tab}%regCode%{Tab}%custID%
	; MsgBox % newEmailInfo 
	; Send % newEmailInfo 
Return 


;GuiEscape: 
GuiClose:
Quit: 
ButtonQuit.:
	Gui, Destroy 
	Exit
Return 

Exit:
	ExitApp
Return

#IfWinActive Account search.ahk - Notepad++
^s::
Send ^s
	Reload() 
Return 

#Include C:\Users\%A_UserName%\AHK\In Use\Lib\MyFunctions.ahk