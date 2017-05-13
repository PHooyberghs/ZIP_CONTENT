option explicit

'''Credits''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' I've learned a lot from scripts by Rob van der Woude
' Rob van der Woude's Scripting Pages
' In particular, these pages
' http://www.robvanderwoude.com/vbstech_databases_access.php
' http://www.robvanderwoude.com/vbstech_files_zip.php
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Arguments''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''procedure''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Variable declaration
dim sArchive,sDestination,blnNew,sSourcePath,sSourceFile,sSourceExtension,otsSource,sOutputFile,otsOutput
dim objFileSystem,objRegExp
dim sZipLocalFileMark,sZipClDirFileMark,sZipClDirEndMark

'dim sPathToCheck

' Get command line arguments
Get_Arguments

' Set/initialize base objects
Base_Objects_Initialize

' Check destination folder/files
wscript.echo "CheckPath"
checkpath sDestination
wscript.echo "Check File"
sOutputFile=sDestination & "\" & sSourcefile & "_content.txt"
wscript.echo "blnNew: " & cstr(blnNew)
sOutputFile=Check_File(sOutputFile,blnNew)
wscript.echo "sOutputFile: " & sOutputFile
set otsOutput=objFileSystem.CreateTextFile(sOutputFile)

' Check source file
SourceFile_Check

' Get content 
Get_Content
sub Get_Content
	dim sLine,iLineNr,sNonPrintable,sMark,sRead
	
	sNonPrintable="[^\w\s\d\.\\\/]"
	objRegExp.pattern=sNonPrintable

	set otsSource=objFileSystem.OpenTextFile(sArchive,1,-1)
	sMark=""
	sLine=""
	iLineNr=0
	
	do while (otsSource.AtEndOfStream<>true) and (instr(sMark,sZipClDirFileMark)=0)
		sRead=otsSource.read(1)
		sMark=right(sMark & sRead,4)
	loop	
	wscript.echo "sMark : " & sMark
	sLine=sMark
	sMark="    "
	do while (instr(sMark,sZipClDirEndMark)=0) and (otsSource.AtEndOfStream<>true) 
		do while (otsSource.AtEndOfStream<>true) and (instr(sMark,sZipClDirEndMark)=0) and (instr(sMark,sZipClDirFileMark)=0)
			sRead=otsSource.read(1)
			sMark=right(sMark & sRead,4)
			sLine=sLine & sRead
		loop
		iLineNr=iLineNr+1
		sLine=left(sLine,len(sLine)-4)
		wscript.echo "LineNr: " & cstr(iLineNr) & ":" & objRegExp.replace(sLine," ")
		wscript.echo "sMark: " & sMark
		wscript.echo "LineNr: " & cstr(iLineNr) & ":" & objRegExp.replace(sLine," ")
		otsOutput.writeline(cstr(iLineNr) & ":")
		otsOutput.writeline(objRegExp.replace(sLine," "))
		sLine=sMark
		sMark=""
	loop
	
	'do while otsSource.AtEndOfStream<>true and instr(sLine,sZipClDirFileMark)=0
	'	iLineNr=iLineNr+1
	'	sLine=right(sLine) & otsSource.read(1)
	'	wscript.echo "LineNr: " & cstr(iLineNr)
	'	if instr(sLine,sZipClDirFileMark)>0 then
	'		wscript.echo "HOERA!!!!!!!!!!" & objRegExp.replace(sLine," ")
	'	end if
	'loop
	'otsOutput.writeline(cstr(iLineNr) & ":")
	'otsOutput.writeline(objRegExp.replace(sLine," "))
	'do while (otsSource.AtEndOfStream<>true) and (instr(sLine,sZipClDirEndMark)=0)
	'	iLineNr=iLineNr+1
	'	sLine=otsSource.readline
	'	wscript.echo "LineNr: " & cstr(iLineNr) & ":" & objRegExp.replace(sLine," ")
	'	otsOutput.writeline(cstr(iLineNr) & ":")
	'	otsOutput.writeline(objRegExp.replace(sLine," "))
	'loop
end sub


' Leave
Normal_Exit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''subs'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Normal exit
function Normal_Exit
	set objRegExp = nothing	
	Set objFileSystem = nothing
	'Set objShell = nothing
	wscript.quit(0)
End function

' Get command line arguments
sub Get_Arguments
	dim iSplit
	with wscript.arguments
		'wscript.echo .unnamed.count
		if .unnamed.count<>2 then
			wscript.echo "Arguments not correct"
		else
			sArchive = ucase(trim(.unnamed(0)))
			sDestination = ucase(trim(.unnamed(1)))
		end if
		blnNew=0
		if .named.count>0 then
			if .Named.Exists("/NEW") then
				blnNew=1
			end if
		end if
	end with
	iSplit=InStrRev(sArchive,"\")
	sSourcePath=left(sArchive,iSplit)
	sSourcefile=right(sArchive,len(sArchive)-iSplit)
	iSplit=InStrRev(sSourcefile,".")
	sSourceExtension=right(sSourcefile,len(sSourcefile)-iSplit)
	sSourcefile=left(sSourcefile,iSplit-1)
	if right(sDestination,1)="\" then sDestination=left(sDestination,len(sDestination)-1)
	'sTmpPath=sDestination & "\tmp"
  
	sZipLocalFileMark=chr("&H" & "50") & chr("&H" & "4B") & chr("&H" & "03") & chr("&H" & "04")
	sZipClDirFileMark=chr("&H" & "50") & chr("&H" & "4B") & chr("&H" & "01") & chr("&H" & "02")
	sZipClDirEndMark=chr("&H" & "50") & chr("&H" & "4B") & chr("&H" & "05") & chr("&H" & "06")
end sub

'set/intialize base objects
sub Base_Objects_Initialize
	Set objFileSystem = CreateObject( "Scripting.FileSystemObject" )
	set objRegExp = new regexp
	objRegExp.IgnoreCase=-1
	objRegExp.Global=-1  
end sub

' check/create folder
Sub CheckPath(sPathToCheck)
	objRegExp.pattern="(.+)(\\[\w\d\s]+)$"
	do while not (objFileSystem.folderexists(sPathToCheck))
		sFolderShort1=sPathToCheck
		do while not (objFileSystem.folderexists(sFolderShort1))
			sFolderShort2=objRegExp.Replace(sFolderShort1,"$1")
			if (objFileSystem.folderexists(sFolderShort2)) then
				objFileSystem.createfolder(sFolderShort1)
				wscript.echo "folder '" & sFolderShort1 & "' created"
				sFolderShort1=sPathToCheck
				exit do
			else
				sFolderShort1=sFolderShort2
			end if
		loop
	loop
end sub

function Check_File(sFile,blnOverwrite)
	dim sPath,sName,sExtension,iSplit,iIdx,sCF
	iSplit=InStrRev(sFile,"\")
	sPath=left(sFile,iSplit)
	sName=right(sFile,len(sFile)-iSplit)
	iSplit=InStrRev(sName,".")
	sExtension=right(sName,len(sName)-iSplit)
	sName=left(sName,iSplit-1) 
	sCF=sFile
	if objFileSystem.FileExists(sCF) then
		wscript.echo "blnOverwrite: " & cstr(blnOverwrite)
		if blnOverwrite>0 then
			iIdx=1
			sCF=sPath & "\" & sName & "_" & cstr(iIdx) & "." & sExtension
			do while objFileSystem.FileExists(sCF)
				iIdx=iIdx+1
				sCF=sPath & "\" & sName & "_" & cstr(iIdx) & "." & sExtension
			loop
			Check_File=sCF
		else
			wscript.echo "deleting file: " & sfile
			objFileSystem.DeleteFile cstr(sFile),true
			Check_File=sFile
		end if
	else
		Check_File=sFile
	end if
end function

'check sourcefile
sub SourceFile_Check
	dim sLine
	if objFileSystem.Fileexists(sArchive)=false then
		wscript.echo "Sourcefile not found:"
		WScript.Quit(2)
	else
		set otsSource=objFileSystem.OpenTextFile(sArchive,1,-1)
		sLine=otsSource.read(4)
		wscript.echo "sLine: " & sLine
		if strcomp(sLine,sZipLocalFileMark)<>0 then
			wscript.echo "Sourcefile seems not being an archive"
			wscript.quit(3)
		end if
	end if
end sub
