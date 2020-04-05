'Destination Folder of CSV file ; PS: finish with \ (windows)
DestFolder = "C:\Users\User\Desktop\"

'Target Folder to Colect txt's
TargetFolder = "C:\Users\User\Downloads\teste"

'Setting File Name
nameFile = "_extract"

'Extesion file generated
ext = ".csv"

'Sanitizing the Data
data = Cstr(Date)
if InStr(data,"/") then    
    data = replace(data,"-","-_-")
End if

'Creating obj's to handle system object
Set objFSO = CreateObject("Scripting.FileSystemObject")	

'Setup target, destination and files
Set csvFile = objFSO.CreateTextFile(DestFolder + data + nameFile + ext, True,True)
Set objFolder = objFSO.GetFolder(TargetFolder)
Set colFiles = objFolder.Files

'Iteration at each file on folder
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
             'splite when necessary
             'splited = Split(line)                        
             'For Each x in splited                            
              '   csvFile.Write(x) + " , "                            
             'Next                                         
         Loop

'New line blank 
        csvFile.WriteLine 
        file.Close              
	End if
Next

'Finishing closing and informing user
csvFile.Close
wscript.Echo "File Generated Click ok to Finish" 