set myName to "Mac Outlook Archive"
set mailName to "Microsoft Outlook"
set inboxName to "Inbox"
set archiveName to "Archive"

tell application "Microsoft Outlook"
	set frontWin to front window
	set winName to name of frontWin
	set currMsgs to current messages
	if currMsgs = {} then
		display notification ("Current window: " & winName) with title myName subtitle "No current messages"
		return 0
	end if
	set firstMsg to item 1 of currMsgs
	set currAccount to account of firstMsg
	set allFolders to mail folder of currAccount
	set theInbox to null
	repeat with theFolder in allFolders
		if name of theFolder is inboxName then
			set theInbox to theFolder
		end if
	end repeat
	if theInbox is null then
		display alert myName message ("Mailbox \"" & inboxName & "\" not found") as critical
		return 0
	end if
	try
		set destFolder to folder archiveName of theInbox
	on error errorMessage number errorNumber
		display alert ("Archive folder \"" & archiveName & "\" not found. ") message (errorMessage & ", errorNumber: " & errorNumber) as critical
		return 0
	end try
	try
		repeat with theMessage in currMsgs
			set (is read) of theMessage to true
			move theMessage to destFolder
		end repeat
	on error errorMessage number errorNumber
		display alert myName message ("Archiving failed. " & errorMessage & ", errorNumber: " & errorNumber) as critical
		return 0
	end try
	set msgCount to (count items in currMsgs)
	if msgCount = 1 then
		set msg to item 1 of currMsgs
		set subj to subject of msg
		display notification with title myName subtitle ("Archived: " & subj)
	else
		display notification with title myName subtitle ("Archived " & msgCount & " messages")
	end if
end tell
