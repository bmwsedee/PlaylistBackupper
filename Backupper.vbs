Set objFSO = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
strBasePath = "C:\Users\Ben\Music\Afspeellijsten"
Set folder = objFSO.GetFolder(strBasePath)

strLogLocation = "C:\projects\logs\CopyMusic.log"
set objLog = objFSO.OpenTextFile(strLogLocation, 8, True)
objLog.WriteLine "--" & Now & "--"

For Each file in folder.Files 
	objLog.WriteLine "Handling " & file.Path
	
	strFileName = file.Name
	strFileName = Left(strFileName, InStrRev(strFileName, "."))
	strNewLoc = strBasePath & "\" & strFileName
	If Not objFSO.FolderExists(strNewLoc) Then
		set newFolder = objFSO.CreateFolder(strNewLoc)
	Else
		objLog.WriteLine "Folder already created"
	End If
	strNewLoc = strNewLoc & "\"
	
	xmlDoc.Load(file.Path)
	strQuery = "/smil/body/seq/media"
	Set colNodes = xmlDoc.selectNodes( strQuery )
	
	For Each objNode in colNodes
		strFileName = objNode.getAttribute("src")
		strFilePath = strBasePath & "\" & strFileName
		
		strSimpleFileName = Left(strFileName, InStrRev(strFileName, "."))
		strSimpleFileName = Right(strSimpleFileName, InStrRev(strSimpleFileName, "\"))
		objLog.WriteLine "   Handling " & strSimpleFileName
	
		If objFSO.FileExists(strFilePath) Then
			If Not objFSO.FileExists(strNewLoc & "\" & strSimpleFileName) Then
				objFSO.CopyFile strFilePath, strNewLoc
			Else
				objLog.WriteLine "   File already exists, skipping"
			End If
		Else
			objLog.WriteLine " Err: File does not exist, skipping"
		End If
	Next
Next

objLog.Close
WScript.Echo "Done"