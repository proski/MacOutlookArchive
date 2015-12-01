# Mac Outlook Archive

## Motivation

There are scripts for mail archival in Microsoft Outlook all over the
Internet, but they are not under version control. You cannot send a patch
or report an issue on a random page. You can do it on GitHub.

The latest versions of Microsoft Outlook for Mac don't include the script
menu. The scripting has to be done in Automator, which requires extra
care. You don't want to archive email accidentally, without seeing what
is being archived.

Microsoft Outlook would often ignore the hotkey key. This script attempts
to inform the user whenever it's run. If everything is fine, a
non-obtrusive notification is used.

## Installation

Create folder "Archive" under "Inbox" if you haven't already done it.

Run Automator (just click on the Launchpad icon and type "Automator").
Select "Service" in the initial dialog. Choose "Utilities" in the
leftmost column. Double click "Run AppleScript" in the second column from
the left.

Copy and paste the contents of MacOutlookArchive.scpt to replace the
suggested script. Above the script, select: Service receives *no input*
in *Microsoft Outlook*. Below the script, check "Ignore this action's
input".

Select File -> Save in the menu, save service as "Mac Outlook Archive" or
any name of your choice. Quit Automator.

Now let's associate the service with a hotkey. Open System Preferences ->
Keyboard -> Shortcuts. Select "Services" in the left column, find the
service under "General", set shortcut to the desired value.

That should be sufficient to make the script work. If it doesn't, check
the Outlook -> Services menu in Microsoft Outlook and run the script from
there.

## Deinstallation

Open Finder. Press Shift-Opt-G and enter ~/Library/Services to open the
Services directory. Drag the service to trash.
