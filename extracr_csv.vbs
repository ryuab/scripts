'Path Destination Folder of the file ; you can finish the path with \
DestFolder = "C:\Users\User\Desktop\"

'Path Target Folder to Colect stream txt's
TargetFolder = "C:\Users\User\Downloads\Targetfolder\"

'Setting the File name
nameFile = "_extract"

'Extesion File generated
ext = ".csv"

'Sanitizing date for systems using xx/xx/xxx format date
sysdate = Cstr(Date)
if InStr(sysdate,"/") then    
    sysdate = replace(sysdate,"/","_")
End if

'Creating obj to handle system object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Setup target, destination and files
Set csvFile = objFSO.CreateTextFile(DestFolder + sysdate + nameFile + ext, True,True) 'file overwriten = true; false to off
Set objFolder = objFSO.GetFolder(TargetFolder)
Set colFiles = objFolder.Files

'Iteration at each stream file at target folder
For Each objFile in colFiles			
	if InStr(objFile,"txt") then
		Set f = objFSO.GetFile(objFile.Path)
		Set file = f.OpenAsTextStream (1,-2)
        Do While file.AtEndOfStream <> True             
             line = file.Readline
             if line <> "" then
                 if file.AtEndOfStream <> True then
                     csvFile.Write(line) + " ,"   
                 else
                     csvFile.Write(line)
                 end if
             end if
             'Splite when necessary
             'splitLine = Split(line)
             'For Each x in splitLine
              '   csvFile.Write(x) + " , "
             'Next
         Loop

'New line blank 
        csvFile.WriteLine

'Closing fetched File
        file.Close       
	End if
Next

'Closing created file and informing to the user the task is done
csvFile.Close
wscript.Echo "File Generated Click ok to Finish"