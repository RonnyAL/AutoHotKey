#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

TVRegex = (?<=[S|s])?(?<Season>\d{1,2})[E|e|X|x](?<Episode>\d{1,3})

SubExtensions := ["srt", "sub", "ssa", "ttml", "sbv", "vtt"]
VideoExtensions := ["mp4", "avi", "mkv"]
RenameActions := []
activePath =

ActionsListView =
LanguageCodeEdit =

#IfWinActive ahk_class CabinetWClass
^R::
	global RenameActions
	global activePath
	
	activePath := GetActiveExplorerPath()
	VideoFiles := []
	SubtitleFiles := []
	RenameActions := []
	
	Loop, Files, %activePath%\*
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
	if (RenameActions.MaxIndex() < 1)
		MsgBox, Found no video files with matching subtitles!
	Else
		ShowGui()

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

ShowGui() {
	global RenameActions
	global activePath
	global vActionsListView
	
	Gui, Destroy
	
	Gui, Font, s15
	Gui, Add, Text,,Subtitle renamer
	Gui, Font, s9
	Gui, Add, ListView, AltSubmit -Multi Checked Hdr r20 w600 vActionsListView, |Episode|Video file|Subtitle file
	LV_Delete()
	
	For index, action in RenameActions
		LV_Add("", Checked, action.ep, action.video, action.sub)

	LV_Modify(0, "Check")
	LV_ModifyCol()
	
	Gui, Add, Text,,Language code:
	Gui, Add, Edit, r1 w30 vLanguageCodeEdit
	Gui, Add, Button, gRenameFiles, &Execute
	
	Gui, Show, AutoSize
	Return
}

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
			SplitPath, % activePath . "\" . videofile,,,,new_filename
			SplitPath, % activePath . "\" . subsfile,,,sub_ext
			FileMove, % activePath . "\" . subsfile, % activePath . "\" . new_filename . langCode .  "." . sub_ext
			RenamedFiles++
		}
	}
	Gui, Destroy
	MsgBox, Successfully renamed %RenamedFiles% files.
}

HasVal(haystack, needle) {
	for index, value in haystack
		if (value = needle)
			return index
	if !IsObject(haystack)
		throw Exception("Bad haystack!", -1, haystack)
	return 0
}

Close:
GuiClose:
GuiEscape:
	Gui, Destroy
Return
