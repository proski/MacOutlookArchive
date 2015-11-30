# How to install

This is how to use Automator to set up a service that can be associated
with a key. Unfortunately, the key will be recognized globally, so use
a key that you are not going to use it in any other application.

Run Automator (just click on Launchpad and enter "Automator"). Select
"Service" in the initial dialog, "Utilities" in the leftmost column,
double click "Run AppleScript" in the second column from the left.

Copy the contents of macoutlookarchive.scpt to replace the suggested
boilerplate code.

Above the script, for "Service receives" select "no input". Save as
"MacOutlookArchive", quit Automator.

Open System Preferences -> Keyboard -> Shortcuts. Select "Services" in
the left column, find MacOutlookArchive, set shortcut to the desired value.

MacOutlookArchive should appear in the main menu of all applications
(application name -> Services), but it would only archive mail if Outlook
is active. In other applications, it would be ignored.

# TODO

The script should pass an unconsumed key to the foreground application,
but I cannot implement that functionality without causing a loop under
some circumstances.

If a message is shown but not selected in the message list, it may not be
archived. Fixes are welcome.
