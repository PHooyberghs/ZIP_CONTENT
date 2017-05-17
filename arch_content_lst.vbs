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
dim vbsFileReverse
dim sArchive,sArchiveRev,sDestination,blnNew,sSourcePath,sSourceFile,sSourceExtension,otsSource,sOutputFile,otsOutput,sCmd
dim objShell,objFileSystem,objRegExp
dim sZipLocalFileMark,sZipClDirFileMark,sZipClDirEndMark

'dim sPathToCheck

' Get command line arguments
Get_Arguments

' Set/initialize base objects
Base_Objects_Initialize

' Retrieve other scripts for execution
GetScripts

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

' Reverse file
objRegExp.Pattern="(.+\\[\w\d\s]+)(\.{1})([\w\d]+)$"
sArchiveRev=objRegExp.Replace(sArchive,"$1$2REV_$3")
wscript.echo "sArchiveRev: " & sArchiveRev
wscript.echo "vbsFileReverse: " & vbsFileReverse
wscript.echo "host path: " & wscript.path
sCmd=wscript.fullname & " "  & vbsFileReverse & " " & chr(34) & sArchive & chr(34) & " " & chr(34) & sArchiveRev & chr(34) 
wscript.echo "sCmd: " & sCmd
wscript.echo "start reversing"
objShell.run sCmd,0,-1
wscript.quit

' Get content 
Get_Content

' Leave
Normal_Exit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''subs'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Normal exit
function Normal_Exit
	set objRegExp = nothing	
	Set objFileSystem = nothing
	Set objShell = nothing
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
	'Set objShell = CreateObject("Shell.Application")
	Set objShell = CreateObject("WScript.Shell")
	Set objFileSystem = CreateObject( "Scripting.FileSystemObject" )
	set objRegExp = new regexp
	objRegExp.IgnoreCase=-1
	objRegExp.Global=-1  
end sub

' Prepare other scripts for execution
Sub GetScripts
	dim tCurrentDir

	objRegExp.pattern="(.+\\)([\w\d]+\.{1}[\w\d]+)$"
	tCurrentDir=objRegExp.replace(wscript.ScriptFullName,"$1")
	
	' Reverse_Asc_File.vbs
	vbsFileReverse=tCurrentDir & "Reverse_Asc_File.vbs"
	
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

' Get content 
sub Get_Content
	dim sLine,iLineNr,sNonPrintable,sMark,sRead,sHexByte,sHex,sDec,sBin,sTxt
	dim sVersionMadeBy,sMajor,sMinor,sVersionNeeded,sGeneralPurposeFlag,sCompressionMethod,sLastModTime,sLastModDate,sCRC32,sSizeCompressed,sSizeUnCompressed
	dim sFileNameLength,sExtraFieldLength,sFileCommentLength,sFileName,sExtraField,sFileComment
	
	sNonPrintable="[^\w\s\d\.\-\+\?\*\{\}\[\]\(\)\^\$\|\\&/]"
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
									
		sHexByte="&H" & hex(asc(mid(sLine,5,1)))
		sDec=Cint(sHexByte)
		sHexByte="&H" & hex(asc(mid(sLine,6,1)))
		sDec=sDec & "-" & (CDbl(sHexByte)/10)
		sVersionMadeBy=sDec	'4
		sHexByte="&H" & hex(asc(mid(sLine,7,1)))
		sTxt=CInt(sHexByte)
		sHexByte="&H" & hex(asc(mid(sLine,8,1)))
		sTxt=sTxt & "." & CInt(sHexByte)
		sVersionNeeded=sTxt	'2
		sGeneralPurposeFlag=chrb("K")	'2
		'sGeneralPurposeFlag=mid(sLine,9,2)	'2
		'sCompressionMethod=mid(sLine,11,2)	'2
		'sLastModTime=mid(sLine,13,2)	'2
		'sLastModDate=mid(sLine,15,2)	'2
		'sCRC32=mid(sLine,17,2)	'4
		'sSizeCompressed=mid(sLine,21,2)	'4
		'sSizeUnCompressed=mid(sLine,25,2)	'4
		'sFileNameLength=mid(sLine,29,2)	'2
		'sExtraFieldLength=mid(sLine,31,2)	'2
		'sFileCommentLength=mid(sLine,33,2)	'2
		'sFileName	
		'sExtraField
		'sFileComment
									
		otsOutput.writeline(cstr(iLineNr) & ":")
		otsOutput.writeline(objRegExp.replace(sLine," "))
		
		otsOutput.writeline("sVersionMadeBy: " & sVersionMadeBy)
		otsOutput.writeline("sVersionNeeded: " & sVersionNeeded)
		otsOutput.writeline("sGeneralPurposeFlag: " & sGeneralPurposeFlag)
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))
		'otsOutput.writeline(objRegExp.replace(sLine," "))

									
		sLine=sMark
		sMark=""
	loop
end sub
							

							
