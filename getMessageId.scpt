use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Microsoft Outlook"
	set selectedMessages to selected objects
	set messageId to id of item 1 of selectedMessages
	set uri to "outlook://" & messageId
	set the clipboard to uri
	display notification "URI " & uri & " copied to clipboard"
end tell