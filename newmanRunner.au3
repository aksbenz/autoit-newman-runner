#include <File.au3>
#include <Array.au3>
#include <Process.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiTreeView.au3>
#include <GuiScrollBars.au3>
#include <GuiRichEdit.au3>
#include <String.au3>
#include <GuiEdit.au3>

Opt("SendKeyDelay",0)
Opt("SendKeyDownDelay",0)
Global Const $CB_CLICKED = -24
Global Const $TEST_MODE = False
Global Const $LOG_LEVEL = 1

#Region ### START Koda GUI section ###
$runner = GUICreate("Newman Runner", 653, 798, 190, 109)
$collection = GUICtrlCreateCombo("", 112, 56, 433, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$allFolders = GUICtrlCreateTreeView(112, 88, 433, 129, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
$env = GUICtrlCreateCombo("", 112, 224, 433, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$global = GUICtrlCreateCombo("", 112, 264, 433, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$reporting = GUICtrlCreateInput("cli,html", 112, 304, 433, 24)
$reportPath = GUICtrlCreateInput("", 112, 328, 369, 24)
$btnRepSelect = GUICtrlCreateButton("Select Folder", 448, 352, 97, 17)
$btnDefaultRepPath = GUICtrlCreateButton("Default", 392, 352, 49, 17)
$template = GUICtrlCreateCombo("", 112, 376, 433, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$btnCertFile = GUICtrlCreateButton("F", 560, 408, 33, 25)
$btnCertClear = GUICtrlCreateButton("Clear", 600, 408, 41, 25)
$btnKeyFile = GUICtrlCreateButton("F", 560, 440, 33, 25)
$btnKeyClear = GUICtrlCreateButton("Clear", 598, 440, 41, 25)
$preCmd = GUICtrlCreateEdit("", 32, 512, 569, 65, BitOR($ES_AUTOVSCROLL,$ES_WANTRETURN,$WS_VSCROLL))
GUICtrlSetData(-1, "preCmd")
$tagIncTest = GUICtrlCreateInput("", 112, 584, 361, 24)
$cmd = GUICtrlCreateEdit("", 32, 640, 577, 121, BitOR($ES_AUTOVSCROLL,$ES_WANTRETURN,$WS_VSCROLL))
GUICtrlSetData(-1, "cmd")
$btnRun = GUICtrlCreateButton("Run", 288, 768, 113, 25)
$btnCopy = GUICtrlCreateButton("Copy", 504, 760, 105, 17)
$btnRefresh = GUICtrlCreateButton("Refresh", 552, 0, 97, 25)
$lblCollection = GUICtrlCreateLabel("Collection", 8, 56, 63, 20)
$lblFolder = GUICtrlCreateLabel("Folder", 8, 96, 43, 20)
$lblEnv = GUICtrlCreateLabel("Env", 8, 224, 27, 20)
$lblGlobal = GUICtrlCreateLabel("Global", 8, 264, 44, 20)
$lblReport = GUICtrlCreateLabel("Report Type", 8, 304, 80, 20)
$lblTemplate = GUICtrlCreateLabel("HTML Template", 8, 376, 102, 20)
$lblPreCmd = GUICtrlCreateLabel("Pre Commands to Run", 32, 496, 138, 20, $WS_CLIPSIBLINGS)
$lblNewmanCmd = GUICtrlCreateLabel("Newman Cmd", 8, 616, 88, 20, $WS_CLIPSIBLINGS)
$sslKey = GUICtrlCreateInput("", 112, 440, 433, 24)
$lblCert = GUICtrlCreateLabel("SSL Cert File", 8, 408, 81, 20)
$lblKey = GUICtrlCreateLabel("SSL Key File", 8, 440, 80, 20)
$settings = GUICtrlCreateCombo("", 112, 0, 433, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$lblBasePathValue = GUICtrlCreateLabel("Base Path", 112, 32, 66, 20)
$lblSettings = GUICtrlCreateLabel("Settings", 8, 0, 52, 20)
$lblBasePath = GUICtrlCreateLabel("Base Path", 8, 32, 66, 20)
$sslCert = GUICtrlCreateInput("", 112, 408, 433, 24)
$btnUnselect = GUICtrlCreateButton("Unselect All", 552, 88, 97, 17)
$btnUndo = GUICtrlCreateButton("Undo", 552, 200, 41, 17)
$ckbFolders = GUICtrlCreateCheckbox("Separate Command for each Folder", 136, 616, 241, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$btnCollapse = GUICtrlCreateButton("Collapse All", 552, 176, 97, 17)
$btnExpand = GUICtrlCreateButton("Expand All", 552, 152, 97, 17)
$lblReportPath = GUICtrlCreateLabel("Report Folder", 8, 336, 99, 20)
$repPathExt = GUICtrlCreateInput("\newman", 480, 328, 65, 24, BitOR($GUI_SS_DEFAULT_INPUT,$ES_READONLY))
$lblTagIncTest = GUICtrlCreateLabel("Tag Include Test", 8, 584, 105, 20)
$btnHistory = GUICtrlCreateButton("History", 504, 776, 105, 17)
$ckbIgnore = GUICtrlCreateCheckbox("Ignore Redirects", 112, 472, 137, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$ckbInsecure = GUICtrlCreateCheckbox("Insecure", 256, 472, 169, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$threads = GUICtrlCreateInput("1", 576, 608, 33, 24)
GUICtrlSetState(-1, $GUI_DISABLE)
$lblThreads = GUICtrlCreateLabel("Threads", 512, 616, 55, 20)
$btnSelect = GUICtrlCreateButton("Select All", 552, 112, 97, 17)
#EndRegion ### END Koda GUI section ###

GUICtrlSetData($preCmd, "")
$cmdPos = ControlGetPos(GUICtrlGetHandle($cmd),"",0)
GUICtrlDelete($cmd)
$cmd = _GUICtrlRichEdit_Create($runner, "", $cmdPos[0], $cmdPos[1], $cmdPos[2], $cmdPos[3], BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL))

;~ TODO: Error if settings file not exists
;~ TODO: Error if no sections in settings file

Dim $path, $collectionPath, $envPath, $templatePath, $collectionList, $envList, $templateList, $defaultSetting, $selectedSetting
Dim $folderClicked, $folderEvent
Dim $folderHistory[1] = ["ID,OPERATION"]
Dim $cmdHistory[1] = ["1"]
$folderClicked = 0
$folderEvent = True

_ArrayDelete($folderHistory,0)
_ArrayDelete($cmdHistory,0)

if FileExists("settings.ini") = 0 Then
	MsgBox($MB_ICONERROR,"Settings File","Settings file not found: settings.ini");
	Exit
EndIf
$sections = IniReadSectionNames('settings.ini')
GUICtrlSetData($settings, _ArrayToString($sections,"|",1),$sections[1])

$hImage = _GUIImageList_Create(16, 16, 5, 3)
_GUIImageList_AddIcon($hImage, "shell32.dll", 110)
_GUIImageList_AddIcon($hImage, "shell32.dll", 131)
_GUICtrlTreeView_SetNormalImageList($allFolders, $hImage)

GUISetState(@SW_SHOW)
$newmanRunnerHandle = WinGetHandle("Newman Runner")
Refresh()
ConsoleWrite("Newman Handle: " & $newmanRunnerHandle & " - " & ControlGetFocus($newmanRunnerHandle))

GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

$prevTagLength = 0
$curTagLength = 0

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $settings
			Refresh()
		Case $btnRun
			ConsoleWrite("BTNRUN CLICKCED" & @CRLF)
			executeNewman()
		Case $collection
			ConsoleWrite("COLLECTION CLICKCED" & @CRLF)
			_GUICtrlTreeView_DeleteAll($allFolders)
			getFolders()
			setFolderTree()
			updateNewmanCmd()
		Case $env, $global, $template, $reporting
			ConsoleWrite("ENV GLOBAL TEMP REP CLICKCED" & @CRLF)
			updateNewmanCmd()
		Case $ckbFolders
			If (GUICtrlRead($ckbFolders) = $GUI_UNCHECKED) Then
				GUICtrlSetState($threads, $GUI_ENABLE)
			Else
				GUICtrlSetState($threads, $GUI_DISABLE)
			EndIf
			updateNewmanCmd()
		case $threads, $ckbIgnore, $ckbInsecure
			updateNewmanCmd()
		Case $btnUnselect
			_ArrayAdd($folderHistory, "ALL" & "," & "UNSELECT")
			unselectAll()
			updateNewmanCmd()
		Case $btnSelect
			_ArrayAdd($folderHistory, "ALL" & "," & "SELECT")
			selectAll()
			updateNewmanCmd()
		Case $btnRefresh
			Refresh()
		Case $btnCollapse
			_GUICtrlTreeView_Expand($allFolders,0,False)
			scrollToTop()
		Case $btnExpand
			_GUICtrlTreeView_Expand($allFolders)
			scrollToTop()
		Case $btnUndo
			Undo()
		Case $btnCertFile
			ConsoleWrite("CERT FILE CLICKCED" & @CRLF)
			Local $sFileOpenDialog = FileOpenDialog("Select CERT File", $path, "All (*.*)", $FD_FILEMUSTEXIST)
			GUICtrlSetData($sslCert,$sFileOpenDialog)
			updateNewmanCmd()
		Case $btnCertClear
			GUICtrlSetData($sslCert,"")
			updateNewmanCmd()
		Case $btnKeyFile
			ConsoleWrite("KEY FILE CLICKCED" & @CRLF)
			Local $sFileOpenDialog = FileOpenDialog("Select KEY File", $path, "All (*.*)", $FD_FILEMUSTEXIST)
			GUICtrlSetData($sslKey,$sFileOpenDialog)
			updateNewmanCmd()
		Case $btnKeyClear
			GUICtrlSetData($sslKey,"")
			updateNewmanCmd()
		Case $btnRepSelect
			ConsoleWrite("REPORT SELECT FOLDER" & @CRLF)
			Local $currFolder = GUICtrlRead($reportPath)
			Local $sFolderSelectDialog = FileSelectFolder("Select Report Folder", $currFolder)
			If $sFolderSelectDialog = "" Then
				GUICtrlSetData($reportPath,$currFolder)
			Else
				GUICtrlSetData($reportPath,$sFolderSelectDialog)
			EndIf
		Case $reportPath
			If GUICtrlRead($reportPath) = "" Then
				GUICtrlSetData($reportPath,$path)
			EndIf
		Case $btnDefaultRepPath
			GUICtrlSetData($reportPath,$path)
		Case $btnCopy
			ClipPut(_GUICtrlRichEdit_GetText($cmd))
		Case $tagIncTest
			updateNewmanCmd()
		Case $btnHistory
;~ 			$hist = _ArrayToString($cmdHistory,@CRLF)
			_ArrayDisplay($cmdHistory,"Command History","",80)
	EndSwitch
	If $folderClicked = 1 Then
		$folderClicked = 0
		$selectedTreeItem = _GUICtrlTreeView_GetSelection($allFolders)
		_ArrayAdd($folderHistory, $selectedTreeItem & "," & _GUICtrlTreeView_GetChecked($allFolders,$selectedTreeItem) & "," & _GUICtrlTreeView_GetText($allFolders,$selectedTreeItem))
;~ 		 MsgBox($MB_SYSTEMMODAL, "Information", StringFormat("Item handle for index %d: %s\r\nIsPtr = %d IsHWnd = %d", 0, _GUICtrlTreeView_GetItemHandle($folderHistory, $selectedTreeItem), _
;~             IsPtr(_GUICtrlTreeView_GetItemHandle($folderHistory, $selectedTreeItem)), IsHWnd(_GUICtrlTreeView_GetItemHandle($folderHistory, $selectedTreeItem))))
		updateNewmanCmd()
	EndIf
	; Update Newman Command on any input to Tag Include Test field
	If StringLen(GUICtrlRead($tagIncTest)) <> $prevTagLength Then
		$prevTagLength = StringLen(GUICtrlRead($tagIncTest))
		updateNewmanCmd()
		_GUICtrlEdit_SetSel($tagIncTest,-1,0)
	EndIf
WEnd

Func Refresh()
	clearUI()
	$selectedSetting = GUICtrlRead($settings)
	$path = _PathFull(@ScriptDir & IniRead("settings.ini",$selectedSetting,"basePath","\..")) & "\"
	$collectionPath = $path & IniRead("settings.ini",$selectedSetting,"collectionfolder","TestCollections")
	$envPath = $path & IniRead("settings.ini",$selectedSetting,"environmentfolder","TestEnvironments")
	$templatePath = $path & IniRead("settings.ini",$selectedSetting,"templatefolder","TestResultTemplate")

	; Error checking for folders
	if FileExists($collectionPath) = 0 Then
		MsgBox($MB_ICONERROR,"Folder Not Found","Collection Folder not found: " & $collectionPath);
		Exit
	EndIf
	If FileExists($envPath) = 0 Then
		MsgBox($MB_ICONERROR,"Folder Not Found","Env Folder not found: " & $envPath);
		Exit
	EndIf
	If FileExists($templatePath) = 0 Then
		MsgBox($MB_ICONERROR,"Folder Not Found","Template Folder not found: " & $templatePath);
		Exit
	EndIf

	Local $collectionList = _FileListToArray($collectionPath, "*.json",1)
	Local $envList = _FileListToArray($envPath, "*.json",1)
	Local $templateList = _FileListToArray($templatePath, "*.hbs",1)

	; Error checking for files
	if $collectionList[0] = 0 Then
		MsgBox($MB_ICONERROR,"Files Not Found","No JSON Files in Collection Folder: " & $collectionPath);
		Exit
	EndIf
	if $envList[0] = 0 Then
		MsgBox($MB_ICONERROR,"Files Not Found","No JSON Files in Environment Folder: " & $collectionPath);
		Exit
	EndIf
	if $templateList[0] = 0 Then
		MsgBox($MB_ICONERROR,"Files Not Found","No MBS Files in Template Folder: " & $collectionPath);
		Exit
	EndIf

	$defaultCollection = getDefault($collectionList,IniRead("settings.ini",$selectedSetting,"defaultcollection",$collectionList[1]))
	$defaultEnv = getDefault($envList,IniRead("settings.ini",$selectedSetting,"defaultenvironment",$envList[1]))
	$defaultGlobal = getDefault($envList,IniRead("settings.ini",$selectedSetting,"defaultglobal","globals.json"))
	$defaultTemplate = getDefault($templateList,IniRead("settings.ini",$selectedSetting,"defaulttemplate",$templateList[1]))

	GUICtrlSetData($lblBasePathValue,$path)
	GUICtrlSetData($collection, _ArrayToString($collectionList,"|",1) , $defaultCollection)
	GUICtrlSetData($env, _ArrayToString($envList,"|",1) , $defaultEnv)
	GUICtrlSetData($global, _ArrayToString($envList,"|",1) , $defaultGlobal)
	GUICtrlSetData($template, _ArrayToString($templateList,"|",1) , $defaultTemplate)
	GUICtrlSetData($reportPath,$path)
	GUICtrlSetState($ckbFolders, $GUI_CHECKED)
	GUICtrlSetData($threads, "1")
	GUICtrlSetState($threads, $GUI_DISABLE)
	GUICtrlSetState($ckbIgnore, $GUI_CHECKED)
	GUICtrlSetState($ckbInsecure, $GUI_CHECKED)
	GUICtrlSetData($tagIncTest,"")
	GUICtrlSetData($preCmd,"")

	getFolders()
	setFolderTree()
	updateNewmanCmd()
EndFunc

Func clearUI()
	GUICtrlSetData($lblBasePathValue,"")
	GUICtrlSetData($collection, "")
	GUICtrlSetData($env, "")
	GUICtrlSetData($global, "")
	GUICtrlSetData($template, "")
	GUICtrlSetData($sslCert,"")
	GUICtrlSetData($sslKey,"")
	_GUICtrlTreeView_DeleteAll($allFolders)
EndFunc

Func getFilesList($folderPath,$fileType)
	Local $filter = ""
	If $fileType = "" Then
		$filter = "*"
	Else
		$filter = "*." & $fileType
	EndIf
    Local $aFileList = _FileListToArray($folderPath, $filter,1)
EndFunc

func getDefault($arr, $val)
	Local $idx = _ArraySearch($arr,$val)
	Local $name = ""
	if $idx <> -1 Then
		$name = $arr[$idx]
	Else
		$name= $arr[1]
	EndIf
	Return $name
endfunc

Func getFolders()
	$selectedCollection = GUICtrlRead($collection)
	Local $collFilePath = $collectionPath & "\" & $selectedCollection
	FileDelete("list2.txt")
	ShellExecuteWait("jsonreader.exe","""" & $collFilePath & """" & " list2.txt","","",@SW_HIDE)
;~ 	Local $fileContent = FileRead("list2.txt")
;~ 	if ($fileContent = "") Then
;~ 		MsgBox($MB_ICONERROR,"Error","Could not read collection file")
;~ 	EndIf
EndFunc

Func setFolderTree()
	; Enable checkboxes for tree view
	GUICtrlSetStyle($allFolders,BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))

	Local $list[1] = [""]
	Local $tree
	$tree = ObjCreate("Scripting.Dictionary")
	_ArrayDelete($list,0)
	Local $opn = _FileReadToArray('list2.txt',$list)
	Local $itemid

	; If no folders are present then disable checkboxes and show no folder string
	If $opn = 0 OR UBound($list) = 0 Then
		GUICtrlSetStyle($allFolders,$GUI_SS_DEFAULT_TREEVIEW)
		_GUICtrlTreeView_Add($allFolders, 0, "Error Finding Folders", 1, 1)
		Return
	EndIf

	; Create Tree View
	For $i=1 to $list[0]
		Local $fdr = $list[$i]
		logger($fdr,8)
		if StringInStr($fdr,'/') = 0 Then
			$itemid = _GUICtrlTreeView_Add($allFolders, 0, correctFldrName($fdr), 0, 0)
			logger("  Itemid: " & $itemid,8)
			$tree.ADD ($fdr, $itemid)
		Else
			$parent = StringMid($fdr,1,StringInStr($fdr,'/',0,-1)-1)
			$child = StringMid($fdr,StringInStr($fdr,'/',0,-1)+1)
			$parentTree = $tree.Item ($parent)
			$itemid = _GUICtrlTreeView_AddChild($allFolders, $parentTree, correctFldrName($child), 0, 0)
			logger("  Parent: " & $parent & " ITEMID: " & $parentTree,8)
			logger("  Child: " & correctFldrName($child) & " ITEMID: " & $itemid,8)

			;TODO: Find parent WILL FAIL if this case
			;Handle Duplicate folder names
			;TODO: Replace RANDOM with function to incrementally check tree for available number
			if $tree.Exists ($fdr) = True Then
				$fdr = $fdr & "(" & Random(1,99) & ")"
			EndIf

			$tree.ADD ($fdr, $itemid)
		EndIf
	Next

;~ 	_GUICtrlTreeView_Expand($allFolders)
	scrollToTop()
EndFunc

Func correctFldrName($name)
	$name = StringReplace($name,"0x2f","/")
	$name = StringRegExpReplace($name,'O\d+_','',1)
	return $name
EndFunc

Func scrollToTop()
	Local $hTreeView = GUICtrlGetHandle($allFolders)
	_GUIScrollBars_Init($hTreeView)
	_GUIScrollBars_SetScrollInfoPos($hTreeView,0,0)
	_GUIScrollBars_SetScrollInfoPos($hTreeView,1,0)
;~ 	_GUIScrollBars_ScrollWindow($hTreeView,1,1)
	_GUICtrlTreeView_EndUpdate($hTreeView)
EndFunc

Func updateNewmanCmd()
	; _GUICtrlRichEdit changes focus, hence capturing the original focused field to revert the focus at end
	$currFocus = ControlGetFocus($newmanRunnerHandle)

	_GUICtrlRichEdit_SetText($cmd, "")

	$selectedCollection = GUICtrlRead($collection)
	Local $selectedFolders = getCheckedFolders()
	$selectedFolder = ""
	$selectedEnv = GUICtrlRead($env)
	$selectedGlobal = GUICtrlRead($global)
	$selectedReport = GUICtrlRead($reporting)
	$selectedTemplate = GUICtrlRead($template)
	$selectedCert = GUICtrlRead($sslCert)
	$selectedKey = GUICtrlRead($sslKey)
	$tagIncludeTest = GUICtrlRead($tagIncTest)
	$certCommand = ""
	$multiFolderCmds = BitAND(GUICtrlRead($ckbFolders), $GUI_CHECKED)
	$threadsCmd = ""
	$tagCmd = ""
	$ignRdCmd = ""
	$ignInsec = ""

	Local $cmdName = "newman-ext "
	If ($multiFolderCmds = True) Then
		$cmdName = "newman "
	EndIf
	If UBound($selectedFolders) > 0 Then
		$selectedFolder =  " --folder " & """" & $selectedFolders[0] & """"
	EndIf
	If $selectedReport <> "" Then
		$selectedReport = " -r " & """" & $selectedReport & """"
	EndIf
	If $selectedCert <> "" AND $selectedKey <> "" Then
		$certCommand = " --ssl-client-cert " & """" & $selectedCert & """" & " --ssl-client-key " & """" & $selectedKey & """"
	EndIf
	If (GUICtrlRead($ckbFolders) = $GUI_UNCHECKED) Then
		If (Number(GUICtrlRead($threads)) > 0 AND GUICtrlGetState($threads) = ($GUI_ENABLE+$GUI_SHOW) ) Then
			$threadsCmd = " --threads " & Number(GUICtrlRead($threads)) & " "
		EndIf
	EndIf
	If $tagIncludeTest <> "" Then
		$tagCmd =  " --tags " & """" & $tagIncludeTest & """"
	EndIf
	If (GUICtrlRead($ckbIgnore) = $GUI_CHECKED) Then
		$ignRdCmd = " --ignore-redirects"
	EndIf
	If (GUICtrlRead($ckbInsecure) = $GUI_CHECKED) Then
		$ignInsec = " -k"
	EndIf

	$newmanCmd = $cmdName & "run " & """" & $collectionPath & "\" & $selectedCollection & """ -e " & """"& $envPath & "\" & $selectedEnv & """ -g " & """"& $envPath & "\" & $selectedGlobal & """" & $selectedReport & " --reporter-html-template " & """"& $templatePath& "\" & $selectedTemplate & """" & $certCommand & $threadsCmd & $tagCmd & $ignInsec & $ignRdCmd & $selectedFolder
	Local $newmanCmds[1] = [$newmanCmd]

	For $i=1 to UBound($selectedFolders)-1
		$selectedFolder =  " --folder " & """" & $selectedFolders[$i] & """"
		if ($multiFolderCmds = True) Then
			$newmanCmd = $cmdName & "run " & """"& $collectionPath & "\" & $selectedCollection & """ -e " & """"& $envPath & "\" & $selectedEnv & """ -g " & """"& $envPath & "\" & $selectedGlobal & """" & $selectedReport & " --reporter-html-template " & """" & $templatePath & "\" & $selectedTemplate & """" & $certCommand & $threadsCmd & $tagCmd & $ignInsec & $ignRdCmd & $selectedFolder
			_ArrayAdd($newmanCmds,$newmanCmd)
		Else
			$newmanCmd = $newmanCmd & $selectedFolder
			$newmanCmds[0] = $newmanCmd
		EndIf
	Next

	_GUICtrlRichEdit_PauseRedraw($cmd)
	_GUICtrlRichEdit_SetText($cmd, _ArrayToString($newmanCmds,@CRLF))
	If ($multiFolderCmds = False) Then
		colorize()
	EndIf
	_GUICtrlRichEdit_ResumeRedraw($cmd)
	_GUICtrlRichEdit_SetScrollPos($cmd, 0, 0)
	ControlFocus($newmanRunnerHandle,"",$currFocus) ; Set focus back to original field
EndFunc

Func colorize()
	Local $clnewman = Dec('8888FF') ; Light Red
	Local $clCmd = Dec('0A5AF3') ;Orange
	Local $clPath = Dec('f7b48e') ;Light Blue
	Local $clCollection = Dec('aeea02');

	Local $value = _GUICtrlRichEdit_GetText($cmd)
	Local $allValues = StringRegExp($value,'(".*?")|(\s+-[a-z]\s+)|(--[a-zA-Z-]+)',3)
	_ArrayAdd($allValues,"newman run");
	_ArrayAdd($allValues,"newman-ext run");

	For $i=0 to UBound($allValues)-1
		$currVal = StringStripWS($allValues[$i],3)
		if StringLen($currVal) > 0 Then
			StringReplace($value,$currVal,"")
			$count = @extended
;~ 			ConsoleWrite("Count: " & $count & @CRLF)
			For $j=1 to $count
				$start = StringInStr($value,$currVal,0,$j)
;~ 				ConsoleWrite($currVal & " _ " & $start-1 & " - " & $start-1+StringLen($currVal) & @CRLF)
				_GUICtrlRichEdit_SetSel($cmd, $start-1, $start-1+StringLen($currVal))
				If StringLeft($currVal,1) = "-" Then ;Check if it is an option
					_GUICtrlRichEdit_SetCharBkColor($cmd, $clCmd)
				ElseIf StringLeft($currVal,6) = "newman" Then ;Check if it is newman command
					_GUICtrlRichEdit_SetCharBkColor($cmd, $clnewman)
				ElseIf StringMid($value,$start-4,3) = "run" Then ;Check if it is a collection
					_GUICtrlRichEdit_SetCharBkColor($cmd, $clCollection)
				Else
					_GUICtrlRichEdit_SetCharBkColor($cmd, $clPath) ;Else its a path/value
				EndIf
				_GUICtrlRichEdit_Deselect($cmd)
			Next
		EndIf
	Next
EndFunc

Func executeNewman()
	Local $newmanCmdsToRun =StringSplit(_GUICtrlRichEdit_GetText($cmd),@CRLF)
	Local $newmanCmdToRun = ""
	Local $temp
	Local $selectedReportPath = GUICtrlRead($reportPath)

	Local $reportPathDrive
	_PathSplit($selectedReportPath,$reportPathDrive,$temp,$temp,$temp)
	ConsoleWrite($reportPathDrive)

	If Not FileExists($selectedReportPath) Then
		If Not DirCreate($selectedReportPath) Then
			$selectedReportPath = $path
			GUICtrlSetData($reportPath,$path)
			MsgBox($MB_ICONWARNING,"Report Folder Error","Could not create Report Folder" & @CRLF & "Will Run from: " & $path)
		EndIf
	EndIf

	Local $preCmdsToRun = StringSplit(GUICtrlRead($preCmd),@CRLF)
	Local $preCmdToRun = ""

	For $i=1 to $preCmdsToRun[0]
		If $preCmdsToRun[$i] <> "" Then
			$preCmdToRun = $preCmdToRun & $preCmdsToRun[$i] & "&&"
		EndIf
	Next
	$preCmdToRun = StringTrimRight($preCmdToRun,2)

	For $i=1 to $newmanCmdsToRun[0]
		$newmanCmdToRun = $newmanCmdsToRun[$i]
		Local $cmdtorun = $reportPathDrive & "&&" & "cd " & """" & $selectedReportPath & """"
		If $newmanCmdToRun <> "" Then
			if $preCmdToRun <> "" Then
				$cmdtorun = $cmdtorun & "&&" & $preCmdToRun
			EndIf
			$cmdtorun = $cmdtorun & "&&" & $newmanCmdToRun
			If NOT $TEST_MODE Then
				_ArrayAdd($cmdHistory,@YEAR&@MON&@MDAY&"_"&@HOUR&":"&@MIN&":"&@SEC&"."&@MSEC & "&&" & $cmdtorun & "&&",0,"&&")
				Run( @COMSPEC & " /k " & $cmdtorun, "", @SW_SHOW)
			EndIf
		EndIf
	Next

EndFunc

Func getCheckedFolders()
;~ 	ConsoleWrite("**** getCheckedFolders ****" & @CRLF)
	Local $item = _GUICtrlTreeView_GetFirstItem ($allFolders)
	Local $folders[1] = [""]
	Local $stack[1] = [$item]
	Local $allItems[1] = [""]
	_ArrayDelete($allItems,0)
	_ArrayDelete($folders,0)

	Local $sibling = _GUICtrlTreeView_GetNextSibling($allFolders,$item)

	While $sibling <> 0
		_ArrayInsert($stack,0,$sibling)
		$sibling = _GUICtrlTreeView_GetNextSibling($allFolders,$sibling)
	WEnd

	While UBound($stack) > 0
		$val = _ArrayPop($stack)
		$text = _GUICtrlTreeView_GetText($allFolders,$val)
		_ArrayAdd($allItems,$text)

		$childCount = _GUICtrlTreeView_GetChildCount($allFolders, $val)

		; If Folder is checked THEN we ignore any children
		if _GUICtrlTreeView_GetChecked ( $allFolders, $val ) Then
			_ArrayAdd($folders,$text)
		Else
			For $i=$childCount-1 to 0 Step -1
				$child = _GUICtrlTreeView_GetItemByIndex($allFolders,$val,$i)
				_ArrayAdd($stack,_GUICtrlTreeView_GetItemByIndex($allFolders,$val,$i))
			Next
		EndIf
	WEnd

	Return $folders
EndFunc

Func unselectAll()
	$folderClicked = 0
	$folderEvent = False
	Local $item = _GUICtrlTreeView_GetFirstItem ($allFolders)
	Local $stack[1] = [$item]
	Local $sibling = _GUICtrlTreeView_GetNextSibling($allFolders,$item)
	While $sibling <> 0
		_ArrayInsert($stack,0,$sibling)
		$sibling = _GUICtrlTreeView_GetNextSibling($allFolders,$sibling)
	WEnd

	_GUICtrlTreeView_BeginUpdate ($allFolders)
	While UBound($stack) > 0
		$val = _ArrayPop($stack)
		_GUICtrlTreeView_SetChecked($allFolders, $val, false)
		$childCount = _GUICtrlTreeView_GetChildCount($allFolders, $val)
		For $i=$childCount-1 to 0 Step -1
			$child = _GUICtrlTreeView_GetItemByIndex($allFolders,$val,$i)
			_ArrayAdd($stack,_GUICtrlTreeView_GetItemByIndex($allFolders,$val,$i))
		Next
	WEnd
	_GUICtrlTreeView_EndUpdate($allFolders)
	$folderEvent = True
	$folderClicked = 1
EndFunc

Func selectAll()
	$folderClicked = 0
	$folderEvent = False
	Local $item = _GUICtrlTreeView_GetFirstItem ($allFolders)
	Local $stack[1] = [$item]
	Local $sibling = _GUICtrlTreeView_GetNextSibling($allFolders,$item)
	While $sibling <> 0
		_ArrayInsert($stack,0,$sibling)
		$sibling = _GUICtrlTreeView_GetNextSibling($allFolders,$sibling)
	WEnd

	_GUICtrlTreeView_BeginUpdate ($allFolders)
	While UBound($stack) > 0
		$val = _ArrayPop($stack)

		; If Expanded then select the children and NOT select the parent
		; Select PARENT only if NOT expanded or no children
		If (_GUICtrlTreeView_GetExpanded($allFolders,$val)) Then
			$childCount = _GUICtrlTreeView_GetChildCount($allFolders, $val)
			For $i=$childCount-1 to 0 Step -1
				$child = _GUICtrlTreeView_GetItemByIndex($allFolders,$val,$i)
				_ArrayAdd($stack,_GUICtrlTreeView_GetItemByIndex($allFolders,$val,$i))
			Next
		Else
			_GUICtrlTreeView_SetChecked($allFolders, $val, True)
		EndIf
	WEnd
	_GUICtrlTreeView_EndUpdate($allFolders)
	$folderEvent = True
	$folderClicked = 1
EndFunc


Func Undo()
	$folderClicked = 0
	Local $lastOp = _ArrayPop($folderHistory)
	Local $items = StringSplit($lastOp,",")
	Local $treeItem = $items[1]
	If $treeItem = "ALL" Then
		undoAll()
	ElseIf $treeItem <> "" AND UBound($items) > 2 Then
		Local $value
		If $items[2] = "true" Then
			$value = False
		Else
			$value = True
		EndIf
		_GUICtrlTreeView_SetChecked($allFolders,Ptr($treeItem),$value)
	EndIf
	$folderClicked = 1
EndFunc

Func undoAll()
	$folderClicked = 0
	_GUICtrlTreeView_BeginUpdate ($allFolders)
	For $i=0 to UBound($folderHistory)-1
		Local $items = StringSplit($folderHistory[$i],",")
		If UBound($items) > 2 Then
			If $items[1] = "ALL" AND $items[2] = "UNSELECT" Then
				unselectAll()
			ElseIf $items[1] = "ALL" AND $items[2] = "SELECT" Then
				selectAll()
			ElseIf $items[1] <> "ALL" Then
				Local $value
				If $items[2] = "true" Then
					$value = True
				Else
					$value = False
				EndIf
				_GUICtrlTreeView_SetChecked($allFolders,Ptr($items[1]),$value)
			EndIf
		EndIf
	Next
	_GUICtrlTreeView_EndUpdate($allFolders)
	$folderClicked = 1
EndFunc

;~ ConsoleWrite($items[1] & " - " & _GUICtrlTreeView_GetText($allFolders,Ptr($items[1])) & @CRLF)
;~ ConsoleWrite("Unselected" & @CRLF)
;~ MsgBox(0,"Test","Unselected All")

Func WM_NOTIFY($hWndGUI, $iMsgID, $wParam, $lParam)
    #forceref $hWndGUI, $iMsgID
    Local $tagNMHDR = DllStructCreate("int;int;int;int", $lParam)
    If @error Then Return
    Local $iEvent = DllStructGetData($tagNMHDR, 3)
    Select
		Case $wParam = $allFolders
            Switch $iEvent
				Case $CB_CLICKED
					if ($folderEvent = True) Then
						$folderClicked = 1
					EndIF
			EndSwitch
		Case $wParam = $tagIncTest
			ConsoleWrite("Event on tagIncTest" & $iEvent)
    EndSelect
    $tagNMHDR = 0
    Return $GUI_RUNDEFMSG

EndFunc   ;==>WM_Notify_Events

func _Process2Win($pid)
    if isstring($pid) then $pid = processexists($pid)
    if $pid = 0 then return -1
    $list = WinList()
    for $i = 1 to $list[0][0]
        if $list[$i][0] <> "" AND BitAnd(WinGetState($list[$i][1]),2) then
            $wpid = WinGetProcess($list[$i][0])
            if $wpid = $pid then return $list[$i][1]
        EndIf
    next
    return -1
endfunc

func logger($msg, $lvl)
	If ($lvl < $LOG_LEVEL) Then
		ConsoleWrite($msg & @CRLF)
	EndIf
EndFunc