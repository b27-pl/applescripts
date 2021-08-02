# applescripts

## getMessageId.scpt

Allows to get ID of the mail message from MS Outlook on MAC OS. Copies Message Id of selectedo mail to clipboard as "outlook://<messageId>"

Link "outlook://messageId" can be later used by a custom protocol handler

Script based on http://blog.hakanserce.com/post/outlook_automation_mac/

## openOutlookLocation.scpt

In order to implement a custom protocol handler with AppleScript, we simply implement an open location handler as follows:

Then we need to turn this into an app package and edit its Info.plist to define our new URL scheme. Basically we need to go through the following steps:
Save the script as an app package in Script Editor. (Save... > Format: Application)
Open Finder and locate the saved file. Right click and select Show Package Contents.
Under the Contents folder, open Info.plist with your favorite text editor, and add the following content before the last two lines.

TODO: finish this


Opens email with given message ID

Script based on http://blog.hakanserce.com/post/outlook_automation_mac/