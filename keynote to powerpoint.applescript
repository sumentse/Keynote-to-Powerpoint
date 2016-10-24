on run {input, parameters}
	--check if there are files
	if (count of input) is 0 then
		return
	end if

	set pathList to {}

	repeat with itemNum from 1 to count of input
		tell application "System Events"
			copy POSIX path of (container of (item itemNum of input)) to end of pathList
		end tell
	end repeat

	set pathList to (item 1 of pathList as string) & "/"
	set tempFolder to POSIX path of pathList & "PPTX"
	
	--create folder to store powerpoint if folder does not exist
	if not fileExists(tempFolder) then
		do shell script "mkdir " & quoted form of POSIX path of tempFolder
	end if

	--the file exists	
	set the defaultDestinationFolder to POSIX file tempFolder as alias
		
	tell application "Keynote"
		activate
		--repeat for each file
		repeat with keynotefile in input
			--check if file is a keynote file
			if name extension of (info for keynotefile) is not "key" then
				exit repeat
			end if
			open keynotefile
			try
				if playing is true then tell the front document to stop
				
				if not (exists document 1) then error number -128
				
				tell front document
					set documentName to its name
					if documentName ends with ".key" then ¬
						set documentName to text 1 thru -5 of documentName
					set movieCount to the count of every movie of every slide
					set audioClipCount to the count of every audio clip of every slide
				end tell
				
				set MicrosoftPowerPointFileExtension to "pptx"
				
				tell application "Finder"
					set newExportItemName to documentName & "." & MicrosoftPowerPointFileExtension
					set incrementIndex to 1
					repeat until not (exists document file newExportItemName of defaultDestinationFolder)
						set newExportItemName to ¬
							documentName & "-" & (incrementIndex as string) & "." & MicrosoftPowerPointFileExtension
						set incrementIndex to incrementIndex + 1
					end repeat
				end tell
				set the targetFileHFSPath to (defaultDestinationFolder as string) & newExportItemName
				-- EXPORT THE DOCUMENT
				with timeout of 1200 seconds
					export front document to file targetFileHFSPath as Microsoft PowerPoint
				end timeout
				
				close front document without saving
				
			on error errorMessage number errorNumber
				display alert "EXPORT PROBLEM" message errorMessage
				error number -128
			end try
		end repeat
		
	end tell

end run

on fileExists(posixPath)
	return ((do shell script "if test -e " & quoted form of posixPath & "; then
echo 1;
else
echo 0;
fi") as integer) as boolean
end fileExists