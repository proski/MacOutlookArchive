# Mac Outlook Archive

## Motivation

There are scripts for mail archival in Microsoft Outlook all over the
Internet, but they are not under version control. You cannot send a patch
or report an issue on a random page. You can do it on GitHub.

The latest versions of Microsoft Outlook for Mac don't include the script
menu. The scripting has to be done in Automator, which required extra
steps.

Microsoft Outlook would often ignore the hotkey key. This script attempts
to inform the user of all possible problems. If everything is fine,
non-obtrusive notification is used. The user is always informed about the
script execution status.

## Installation

Run Automator (just click on Launchpad and enter "Automator"). Select
"Service" in the initial dialog, "Utilities" in the leftmost column,
double click "Run AppleScript" in the second column from the left.

Copy the contents of MacOutlookArchive.scpt to replace the suggested
boilerplate code.

Above the script, for "Service receives" select "no input". Save as
"MacOutlookArchive", quit Automator.

This is how to use Automator to set up a service that can be associated
with a key. 

Open System Preferences -> Keyboard -> Shortcuts. Select "Services" in
the left column, find MacOutlookArchive, set shortcut to the desired
value. Unfortunately, the key will be recognized in all applications, so
use a key that you are only going to use in Microsoft Outlook.

MacOutlookArchive should appear in the main menu of all applications
(application name -> Services), but it would only archive mail if Outlook
is active. In other applications, it would be do nothing (but you should
still get a notification).

## TODO

Recognize the status of Outlook and warn the user if some dialog is
stealing the focus.

The script should pass an unconsumed key to the foreground application,
but I cannot implement that functionality without causing a loop under
some circumstances.

If a message is shown but not selected in the message list, it may not be
archived. Fixes are welcome.
