strBasePath = "C:\Users\Ben\Music\Afspeellijsten"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
strProgramPath = objFSO.GetAbsolutePathName(".")

Set folder = objFSO.GetFolder(strBasePath)

objDate = Now
strDate = Year(objDate)
If Month(objDate) < 10 Then
	strDate = strDate & "0"
End If
strDate = strDate & Month(objDate)
If Day(objDate) < 10 Then
	strDate = strDate & "0"
End If
strDate = strDate & Day(objDate)
strLogLocation = strProgramPath & "\logs\" & strDate  & ".log"
If Not objFSO.FolderExists(strProgramPath & "\logs") Then
	objFSO.CreateFolder(strProgramPath & "\logs")
End If
set objLog = objFSO.OpenTextFile(strLogLocation, 8, True)
objLog.WriteLine "--" & Now & "--"

For Each file in folder.Files 
	objLog.WriteLine "Handling " & file.Path
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")  
	Set xmlList = CreateObject("Microsoft.XMLDOM")
	
	strFileName = file.Name
	strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1)
	strNewLoc = strBasePath & "\" & strFileName
	If Not objFSO.FolderExists(strNewLoc) Then
		set newFolder = objFSO.CreateFolder(strNewLoc)
	Else
		objLog.WriteLine "Folder already created"
	End If
	strNewLoc = strNewLoc & "\"
	
	strNewPlaylist = strNewLoc & strFileName & ".wpl"
	
	strQuery = "/smil/body/seq/media"
	Set newFiles = Nothing
	Set objSeq = Nothing
	If Not objFSO.FileExists(strNewPlaylist) Then
		WScript.Echo strNewPlaylist
		Set objRoot = xmlList.CreateElement("smil")
		xmlList.AppendChild(objRoot)
		
		Set objHead = xmlList.CreateElement("head")
		objRoot.AppendChild(objHead)
		
		Set objMeta = xmlList.CreateElement("meta")
		objMeta.SetAttribute "name", "Generator"
		objMeta.SetAttribute "content", "PlaylistBackupper"
		objHead.AppendChild(objMeta)
		
		Set objMeta = xmlList.CreateElement("meta")
		objMeta.SetAttribute "name", "Author"
		objMeta.SetAttribute "content", "PlaylistBackupper"
		objHead.AppendChild(objMeta)
		
		Set objTitle = xmlList.CreateElement("title")
		objTitle.Text = strFileName
		objHead.AppendChild(objTitle)
		
		Set objBody = xmlList.CreateElement("body")
		objRoot.AppendChild(objBody)
		
		Set objSeq = xmlList.CreateElement("seq")
		objBody.AppendChild(objSeq)
		
	Else
		xmlList.Load(strNewPlaylist)
		Set newFiles = xmlList.SelectNodes(strQuery)
		newFiles.RemoveAll
		
		Set objSeq = xmlList.SelectSingleNode("/smil/body/seq")
	End If
	
	xmlDoc.Load(file.Path)
	Set colNodes = xmlDoc.SelectNodes(strQuery)
	
	For Each objNode in colNodes
		strFileName = objNode.GetAttribute("src")
		strFilePath = strBasePath & "\" & strFileName
		
		strSimpleFileName = Mid(strFileName, InStrRev(strFileName, "\") + 1)
		objLog.WriteLine "   Handling " & strSimpleFileName
	
		If objFSO.FileExists(strFilePath) Then
			If Not objFSO.FileExists(strNewLoc & strSimpleFileName) Then
				objFSO.CopyFile strFilePath, strNewLoc
			Else
				objLog.WriteLine "   File already exists, skipping"
			End If
			
			Set objMedia = xmlList.CreateElement("media")
			objMedia.SetAttribute "src", ".\" & strSimpleFileName
			objSeq.AppendChild(objMedia)
		Else
			objLog.WriteLine " Err: File does not exist, skipping"
		End If
	Next
	
	xmlList.Save strNewPlaylist
Next

objLog.Close
WScript.Echo "Done"