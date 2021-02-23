#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

TVRegex = i)S?(?<Season>\d{1,2})[E|X](?<Episode>\d{1,3})

SubExtensions := ["srt", "sub", "ssa", "ttml", "sbv", "vtt"]
VideoExtensions := ["mp4", "avi", "mkv"]
RenameActions := []
ActivePath =

IniPath = %A_ScriptDir%\settings.ini
ActionsListView =
LanguageCodeEdit =
CreateCustomTrayMenu()
Return

#IfWinActive ahk_class CabinetWClass
^R::
	global RenameActions
	global ActivePath
	
	ActivePath := GetActiveExplorerPath()
	LoadFolder(ActivePath)
Return
#IfWinActive

GetActiveExplorerPath()
{
	explorerHwnd := WinActive("ahk_class CabinetWClass")
	if (explorerHwnd)
	{
		for window in ComObjCreate("Shell.Application").Windows
		{
			if (window.hwnd==explorerHwnd)
			{
				return window.Document.Folder.Self.Path
			}
		}
	}
}

GetSeasonAndEpisode(filename) {
	global TVRegex
	RegExMatch(filename, TVRegex, m)
	s := Format("{:02}", mSeason)
	e := Format("{:02}", mEpisode)
	if !(mSeason && mEpisode)
		return 0
	return % "S" . s . "E" . e
}

UpdateGui() {
	global RenameActions
	global ActivePath
	global vActionsListView

	Gui, SubRenamer:+LastFoundExist
		If !(WinExist())
			GoSub, CreateGui

	Gui, SubRenamer:Default
	LV_Delete()
	
	For index, action in RenameActions
		LV_Add("", Checked, action.ep, action.video, action.sub)

	LV_Modify(0, "Check")
	LV_ModifyCol(2,"Sort")
	If (RenameActions.MaxIndex() < 1) {
		Loop % LV_GetCount("Col")
			LV_ModifyCol(A_Index,80)
			LV_ModifyCol(1)
	} else {
		LV_ModifyCol()
	}
	Gui, SubRenamer:Show, AutoSize
}

CreateGui:
	Gui, SubRenamer:Font, s15
	Gui, SubRenamer:Add, Text,,Subtitle renamer
	Gui, SubRenamer:Font, s9
	Gui, SubRenamer:Add, ListView, AltSubmit -Multi Checked Hdr r20 w800 vActionsListView, |Episode|Video file|Subtitle file
	LV_ModifyCol(2, "Text")
	
	Gui, SubRenamer:Add, Text,,Language code:
	Gui, SubRenamer:Add, Edit, r1 w30 vLanguageCodeEdit
	Gui, SubRenamer:Add, Button, gRenameFiles, &Execute
	
	Menu, FileMenu, Add, &Open folder`tCtrl+O, OpenFolder
	Menu, MenuBar, Add, &File, :FileMenu
	Gui, SubRenamer:Menu, MenuBar
Return

LoadFolder(FolderPath) {
	global SubExtensions
	global VideoExtensions
	global ActiveFolder
	
	VideoFiles := []
	SubtitleFiles := []
	RenameActions := []
	
	ActiveFolder := FolderPath
	
	If (!FolderPath) {
		UpdateGui()
		Return
	}
	
	Loop, Files, %FolderPath%\*
		If HasVal(SubExtensions, A_LoopFileExt)
			SubtitleFiles.Push(A_LoopFileName)
		Else If HasVal(VideoExtensions, A_LoopFileExt)
			VideoFiles.Push(A_LoopFileName)
	For index, v_filename in VideoFiles
	{
		v_episode := GetSeasonAndEpisode(v_filename)
		if (!v_episode)
			Continue
		For index, s_filename in SubtitleFiles
		{
			s_episode := GetSeasonAndEpisode(s_filename)
			if (!s_episode)
				Continue
			if (v_episode == s_episode) {
				action := {"ep": s_episode, "video": v_filename, "sub": s_filename}
				RenameActions.Push(action)
			}
		}
	}
	UpdateGui()
	
	if (RenameActions.MaxIndex() < 1 && ActivePath != "")
		MsgBox, Found no video files with matching subtitles!
	MsgBox % ActivePath
}

OpenFolder:
	IniRead, LastFolder, % IniPath, Cache, LastFolder, ""
	FileSelectFolder, Folder,*%LastFolder%,2,Select a folder containing video and subtitle files.
	Folder := RegexReplace(Folder, "\\$")
	SaveSetting("Cache", "LastFolder", Folder)
	LoadFolder(Folder)
Return

RenameFiles() {
	GuiControlGet, LanguageCodeEdit
	langCode =
	if (LanguageCodeEdit)
		langCode = .%LanguageCodeEdit%
	Gui, ListView, vActionsListView
	RenamedFiles := 0
	
	Loop % LV_GetCount()
	{
		Gui +LastFound
		SendMessage, 4140, A_Index - 1, 0xF000, SysListView321  ; 4140 is LVM_GETITEMSTATE. 0xF000 is LVIS_STATEIMAGEMASK.
		IsChecked := (ErrorLevel >> 12) - 1  ; This sets IsChecked to true if RowNumber is checked or false otherwise.
		
		if (IsChecked) {
			LV_GetText(videofile, A_Index, 3)
			LV_GetText(subsfile, A_Index, 4)
			SplitPath, % ActivePath . "\" . videofile,,,,new_filename
			SplitPath, % ActivePath . "\" . subsfile,,,sub_ext
			FileMove, % ActivePath . "\" . subsfile, % ActivePath . "\" . new_filename . langCode .  "." . sub_ext
			RenamedFiles++
		}
	}
	Gui, Destroy
	MsgBox, Successfully renamed %RenamedFiles% files.
}

SaveSetting(s, k, v) {
	global IniPath
	IniWrite, % v, % IniPath, % s, % k
}

HasVal(haystack, needle) {
	for index, value in haystack
		if (value = needle)
			return index
	if !IsObject(haystack)
		throw Exception("Bad haystack!", -1, haystack)
	return 0
}

CreateCustomTrayMenu() {
	Menu, Tray, NoStandard
	Menu, Tray, DeleteAll
	Menu, Tray, Add, &Open, LoadFolder
	Menu, Tray, Default, &Open
	Menu, Tray, Add, E&xit, ExitApp
}

LoadFolder:
	LoadFolder("")
Return

ExitApp:
	ExitApp
Return

Close:
GuiClose:
GuiEscape:
	Gui, Destroy
Return
