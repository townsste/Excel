<#
This is a powershell command that is used to copy a specific file based on the name from one folder 
to a different folder.  This is based on the * wildcard. 

Change [FOLDER LOACTION PATH] to the originating folder path location (ex: C:\folderName)
Change [FILENAME] to the name of the file you want to copy-move	(ex: *quick)
Change [DIFFERENT FOLDER LOACTION PATH] to the new folder path location (ex: D:\newFolderName)

This will search the originating folder for everything with the given filename.  
If there are matches then it will copy that file to the destination folder
#>

Get-Childitem "[FOLDER LOACTION PATH]" -recurse -filter "*[FILENAME]" | Copy-Item -Destination "[DIFFERENT FOLDER LOACTION PATH]"