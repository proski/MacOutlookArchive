set myName to "Mac Outlook Categorize Mails"

tell application "Microsoft Outlook"
	set frontWin to front window
	set winName to name of frontWin
	set currMsgs to current messages
	if currMsgs = {} then
		display notification ("Current window: " & winName) with title myName subtitle "No current messages"
		return 0
	end if
	set firstMsg to item 1 of currMsgs
	
	--- http://stackoverflow.com/questions/17887174/assign-outlook-for-mac-categories-with-applescript
	set theCategories to the text returned of (display dialog "Categories" default answer "")
	
	--- Get comma separated string and convert to list
	set old_delimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {", "}
	set category_list to text items of theCategories
	set AppleScript's text item delimiters to old_delimiters
	
	--- The set category needs a special list to work with multiple tags...
	--- Create a list {category "tag1", category "tag2", ...}
	set my_categorylist to {}
	repeat with categoryitem in category_list
		set end of my_categorylist to category categoryitem
	end repeat
	
	try
		repeat with theMessage in currMsgs
			set category of (item 1 of theMessage) to my_categorylist
		end repeat
	on error errorMessage number errorNumber
		display alert myName message ("Categorizing failed. " & errorMessage & ", errorNumber: " & errorNumber) as critical
		return 0
	end try
	
	set msgCount to (count items in currMsgs)
	if msgCount = 1 then
		set msg to item 1 of currMsgs
		set subj to subject of msg
		display notification with title myName subtitle ("Categorized: " & subj)
	else
		display notification with title myName subtitle ("Categorized " & msgCount & " messages")
	end if
end tell
