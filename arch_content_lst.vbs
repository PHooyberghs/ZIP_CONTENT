option explicit

'''Credits''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' I've learned a lot from scripts by Rob van der Woude
' Rob van der Woude's Scripting Pages
' In particular, these pages
' http://www.robvanderwoude.com/vbstech_databases_access.php
' http://www.robvanderwoude.com/vbstech_files_zip.php
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' scripts required in same folder as this script:
'	- Reverse_Asc_File.vbs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''sample call and Arguments''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
' C:\Windows\System32\cscript [script-path]\zip_content.vbs "archive" "output folder" [/NEW]
'
' archive: full qualified path (including filename) to compressed file (fileheader &H504B0304 in byte 1 to 4 required)
' output folder: fulle qualified path to folder where outputfile has to be created
' /NEW: optional, when specified a new file (with index suffix) is created when there's already a file with the default name in the output folder
'
' output:
'	CSV-file
'	name: base name archive, suffix '_content' added, extension '.csv'
'	content: for each folder/file in the compressed file: 
'		index (actually the number the item appears in the central directory of the compressed file, but in reversed order)
'		filename/foldername
'		filesize (uncompressed, 0 when folder)
'		date & time last modification (of course, before compressing)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''procedure''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' Variable declaration
dim vbsFileReverse
dim sArchive,sArchiveRev,sDestination,blnNew,sSourcePath,sSourceFile,sSourceExtension,otsSource,sOutputFile,otsOutput,sCmd
dim objShell,objFileSystem,objRegExp
dim sZipLocalFileMark,sZipClDirFileMark,sZipClDirEndMark
dim iError,sError


' Get command line arguments
Get_Arguments

' Set/initialize base objects
Base_Objects_Initialize

' Retrieve other scripts for execution
GetScripts

' Check source file
wscript.echo "Check source file: " & sArchive
SourceFile_Check

' Check destination folder/files, set output textstream
wscript.echo "Check/create destination path '" & sDestination & "'"
checkpath sDestination,1
sOutputFile=sDestination & "\" & sSourcefile & "_content.csv"
wscript.echo "Check/create output file '" & sOutputFile & "'"
sOutputFile=Check_File(sOutputFile,blnNew)
wscript.echo "Finale name output file: " & sOutputFile
set otsOutput=objFileSystem.CreateTextFile(sOutputFile)


' Reverse file
objRegExp.Pattern="(.+\\[\w\d\s]+)(\.{1})([\w\d]+)$"
sArchiveRev=objRegExp.Replace(sArchive,"$1$2REV_$3")
'wscript.echo "sArchiveRev: " & sArchiveRev
'wscript.echo "vbsFileReverse: " & vbsFileReverse
'wscript.echo "host path: " & wscript.path
sCmd=wscript.fullname & " "  & vbsFileReverse & " " & chr(34) & sArchive & chr(34) & " " & chr(34) & sArchiveRev & chr(34) 
'wscript.echo "sCmd: " & sCmd
wscript.echo "start reversing"
objShell.run sCmd,0,-1
wscript.echo "reversing completed"

' Get content 
Get_Content

' Leave


Normal_Exit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''subs'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Normal exit
sub Normal_Exit
'	set objRegExp = nothing
'	Set objFileSystem = nothing
'	Set objShell = nothing
	wscript.echo
	wscript.echo
	wscript.echo "--------------------------------------------------------------------------------"
	wscript.echo "Content list archive "
	wscript.echo sArchive
	wscript.echo "stored in : " 
	wscript.echo sOutputFile
	wscript.echo"--------------------------------------------------------------------------------"
	wscript.echo
	CleanUp
	wscript.quit(0)
End sub

' Error-exit
sub Error_Exit
	wscript.echo
	wscript.echo
	wscript.echo "--------------------------------------------------------------------------------"
	wscript.echo "!!! ERROR"
	if iError>0 then wscript.echo "Error " & iError & ": " & sError
	if iError>5 and iError<50 then
		wscript.echo "Content list archive "
		wscript.echo sArchive
		wscript.echo "could not be retrieved" 
		wscript.echo sOutputFile
	end if
	wscript.echo"--------------------------------------------------------------------------------"
	wscript.echo
	cleanup
	wscript.quit(1)
end sub
	
' Clean up
sub CleanUp
	if IsObject(objFileSystem) then
		if objFileSystem.FileExists(sArchiveRev)then objFileSystem.deletefile(sArchiveRev)
		set objFileSystem=nothing
	end if
	if IsObject(objRegExp) then set objRegExp=nothing
	if IsObject(objShell) then set objShell=nothing
end sub

' Get command line arguments
sub Get_Arguments
	dim iSplit
	with wscript.arguments
		'wscript.echo .unnamed.count
		if .unnamed.count<>2 then
			iError=1
			sError="Arguments not correct (source archive and/or destination path not specified)"
			Error_Exit
		else
			sArchive = ucase(trim(.unnamed(0)))
			sDestination = ucase(trim(.unnamed(1)))
		end if
		blnNew=0
		if iError=0 and .named.count>0 then
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
	if  objFileSystem.Fileexists(vbsFileReverse)=false then
		iError=2
		sError="Required script 'Reverse_Asc_File.vbs' not found in same folder as this script."
		Error_exit
	end if
	
end sub

'check sourcefile
sub SourceFile_Check
	dim sLine
	if  objFileSystem.Fileexists(sArchive)=false then
		iError=3
		sError="Sourcefile not found: " & vbCrLf & sArchive
		Error_exit
	else
		iError=0
		sError=""
		On Error resume next	
		set otsSource=objFileSystem.OpenTextFile(sArchive,1,-1)
		if Err.Number <> 0 then
			iError=Err.number
			sError=Err.Description 
			Error_exit
		end if 
		on error goto 0
		sLine=otsSource.read(4)
		'wscript.echo "sLine: " & sLine
		if strcomp(sLine,sZipLocalFileMark)<>0 then
			iError=4
			sError="Source file does not appear to be a compressed file"
			Error_Exit
		end if
	end if
end sub

' check/create folder
Sub CheckPath(sPathToCheck,blnCreate)
	dim sFolderShort1,sFolderShort2
	iError=0
	sError=""
	On Error resume next
	objRegExp.pattern="(.+)(\\[\w\d\s]+)$"
	do while not (objFileSystem.folderexists(sPathToCheck)) and blnCreate=1
		sFolderShort1=sPathToCheck
		do while not (objFileSystem.folderexists(sFolderShort1))
			sFolderShort2=objRegExp.Replace(sFolderShort1,"$1")
			if (objFileSystem.folderexists(sFolderShort2)) then
				objFileSystem.createfolder(sFolderShort1)
				if Err.Number <> 0 then
					iError=Err.number
					sError=Err.Description & " (path: " & sFolderShort1 & " )"
					Error_exit
				end if 
				'wscript.echo "folder '" & sFolderShort1 & "' created"
				sFolderShort1=sPathToCheck
				exit do
			else
				sFolderShort1=sFolderShort2
			end if
		loop
	loop
	on error goto 0
	if not (objFileSystem.folderexists(sPathToCheck)) then
		iError=5
		sError="Destination path is not found/not created"
		Error_Exit
	end if
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
		'wscript.echo "blnOverwrite: " & cstr(blnOverwrite)
		if blnOverwrite>0 then
			iIdx=1
			sCF=sPath & "\" & sName & "_" & cstr(iIdx) & "." & sExtension
			do while objFileSystem.FileExists(sCF)
				iIdx=iIdx+1
				sCF=sPath & "\" & sName & "_" & cstr(iIdx) & "." & sExtension
			loop
			Check_File=sCF
		else
			'wscript.echo "deleting file: " & sfile
			objFileSystem.DeleteFile cstr(sFile),true
			Check_File=sFile
		end if
	else
		Check_File=sFile
	end if
end function

' Get content 
sub Get_Content
	dim sNonPrintable,sRead,iIdx,sLine,sLineHexCD,sLineHEX,sLineOUT,iLineNr,iOffset,iOffsetAsc,sMark,sHexByte,sHex,sDec,sBin,sTxt
	dim iNDisks,iDiskStart,iDiskNRecords,iTotNRecords,iCentrDirSize,iCentrDirStart,iCommentLength,sComment
	dim sVersionMadeBy,sMajor,sMinor,sVersionNeeded,sGeneralPurposeFlag,sCompressionMethod,sLastModTime,sLastModDate,sCRC32,sSizeCompressed,sSizeUnCompressed
	dim sFileNameLength,sExtraFieldLength,sFileCommentLength,iStartDisk,iFileAttributesInt,lFileAttributesExt,lLocFileHeadStart,sFileName,sExtraField,sFileComment
	
	sNonPrintable="[^\w\s\d\.\-\+\?\*\{\}\[\]\(\)\^\$\|\\&/]"
	objRegExp.pattern=sNonPrintable

	set otsSource=objFileSystem.OpenTextFile(sArchiveRev,1,-1)
	if err.number<>0 then
		iErr=Err.Number
		sError=Err.description
		Error_Exit
	end if
	
	sMark=""
	sLine=""
	sLineHex=""
	sLineHexCD=""
	iLineNr=0
	
	'End of central directory record (EOCD)
	
		'locate
		iIdx=0
		do while (otsSource.AtEndOfStream<>true) and (instr(sLine,sZipClDirEndMark)=0) and iIdx<256
			iIdx=iIdx+1
			sRead=otsSource.read(1)
			sLine=sRead & sLine
			sLineHexCD=AscToHex(sRead) & " " & sLineHexCD
			'wscript.echo "sLine: " & iIdx & " - " & objRegExp.Replace(sLine,"")
		loop
		'wscript.echo "sLineHexCD: " & sLineHexCd
		'wscript.echo
		
		' parse fields
		sLineHEX="&H" & mid(sLineHexCD,16,2) &  mid(sLineHexCD,13,2)
		iNDisks=CInt(sLineHEX)
		sLineHEX="&H" & mid(sLineHexCD,22,2) &  mid(sLineHexCD,19,2)  'AscToHex(mid(sLine,8,1)) &  AscToHex(mid(sLine,7,1))
		iDiskStart=CInt(sLineHEX)
		sLineHEX="&H" & mid(sLineHexCD,28,2) &  mid(sLineHexCD,25,2)  'AscToHex(mid(sLine,10,1)) &  AscToHex(mid(sLine,9,1))
		iDiskNRecords=CInt(sLineHEX)
		sLineHEX="&H" & mid(sLineHexCD,34,2) &  mid(sLineHexCD,31,2)  'AscToHex(mid(sLine,12,1)) &  AscToHex(mid(sLine,11,1))
		iTotNRecords=CInt(sLineHEX)
		sLineHEX="&H" & mid(sLineHexCD,46,2) &  mid(sLineHexCD,43,2) & mid(sLineHexCD,40,2) &  mid(sLineHexCD,37,2) 'AscToHex(mid(sLine,16,1)) &  AscToHex(mid(sLine,15,1)) & AscToHex(mid(sLine,14,1)) &  AscToHex(mid(sLine,13,1))
		iCentrDirSize=CLng(sLineHEX)
		sLineHEX="&H" & mid(sLineHexCD,58,2) &  mid(sLineHexCD,55,2) & mid(sLineHexCD,52,2) &  mid(sLineHexCD,49,2) 'AscToHex(mid(sLine,20,1)) &  AscToHex(mid(sLine,19,1)) & AscToHex(mid(sLine,18,1)) &  AscToHex(mid(sLine,17,1))
		iCentrDirStart=Clng(sLineHEX)
		sLineHEX="&H" & mid(sLineHexCD,64,2) &  mid(sLineHexCD,61,2)  'AscToHex(mid(sLine,22,1)) &  AscToHex(mid(sLine,21,1)) 
		iCommentLength=Clng(sLineHEX)
		sComment=mid(sLine,23,iCommentLength)

		'wscript.echo "iNDisks: " & iNDisks
		'wscript.echo "iDiskStart: " & iDiskStart
		'wscript.echo "iDiskNRecords: " & iDiskNRecords
		'wscript.echo "iTotNRecords: " & iTotNRecords
		'wscript.echo "iCentrDirSize: " & iCentrDirSize
		'wscript.echo "iCentrDirStart: " & iCentrDirStart
		'wscript.echo "iCommentLength: " & iCommentLength
		'wscript.echo "sComment: " & sComment
		

		
	' Central directory records
		' get records
		'wscript.echo "sLine: "
		'wscript.echo objRegExp.Replace(sLine," ") & vbCrLf & "*************"
		'wscript.echo
		'wscript.echo
		'wscript.echo
		'wscript.echo "next"
		'sLine="[******************]" & sLine
		'wscript.echo "iCentrDirSize: " & iCentrDirSize & " - iIdx: " & iIdx
		'iCentrDirSize=iCentrDirSize+iIdx
		for iIdx=1 to iCentrDirSize
			sRead=otsSource.read(1)
			sLine=sRead & sLine
			sLineHexCD=AscToHex(sRead) & " " & sLineHexCD
		next	
		'wscript.echo "sLine next: " 
		'wscript.echo objRegExp.Replace(sLine," ") & vbCrLf & "*************"
		'wscript.echo "length: " & cstr(len(sLine))
		'wscript.echo
		'wscript.echo "sLineHexCD: " & sLineHexCD
		'wscript.echo "sLine:"
		'wscript.echo "[start sLine]" & objRegExp.replace(sLine," ") & "[end sLine]"
		'wscript.echo

		iOffset=1
		iOffsetAsc=1
		' parse fields per record
		'Prepare output file
		sLineOut="Index,FileName,Size,Last modification"
		otsOutput.writeline(sLineOut)

		for iIdx=1 to iTotNRecords
			sLineOUT=""
			
			'wscript.echo "iOffset: " & iOffset
			'wscript.echo "iOffsetAsc: " & iOffsetAsc
			'wscript.echo
			
			sMark="&H" & mid(sLineHexCD,iOffset,2) &  mid(sLineHexCD,iOffset+3,2) & mid(sLineHexCD,iOffset+6,2) &  mid(sLineHexCD,iOffset+9,2)
			'wscript.echo "sMark " & iIdx & " : " & sMark
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+12,2)
			sVersionMadeBy=cstr(clng(sLineHEX)/10)
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+15,2)
			sVersionMadeBy=cstr(cint(sLineHEX)) & "-" & sVersionMadeBy & " (" & sLineHEX & ")"
			'wscript.echo "sVersionMadeBy " & iIdx & " : " & sVersionMadeBy
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+21,2) & mid(sLineHexCD,iOffSet+18,2)
			sVersionNeeded=cstr(cdbl(sLineHEX)) 'cstr(cdbl(sLineHEX)/10)
			'wscript.echo "sVersionNeeded " & iIdx & " : " & sVersionNeeded & " (" & sLineHEX & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+27,2) & mid(sLineHexCD,iOffSet+24,2)
			sGeneralPurposeFlag=HexToBin(sLineHEX)
			'wscript.echo "sGeneralPurposeFlag " & iIdx & " : " & sGeneralPurposeFlag & " (" & sLineHEX & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+33,2) & mid(sLineHexCD,iOffSet+30,2)
			sCompressionMethod=cint(sLineHEX)
			'wscript.echo "sCompressionMethod " & iIdx & " : " & sCompressionMethod & " (" & sLineHEX & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+39,2) & mid(sLineHexCD,iOffSet+36,2)
			sLastModTime=HexToDosTime(sLineHex)  'cdate(clng(sLineHex))
			'wscript.echo "sLastModTime " & iIdx & " : " & sLastModTime & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+45,2) & mid(sLineHexCD,iOffSet+42,2)
			sLastModDate=HexToDosDate(sLineHex) 'clng(sLineHex) 'cdate(clng(sLineHex)) + sLastModDate
			'wscript.echo "sLastModDate " & iIdx & " : " & sLastModDate & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+57,2) & mid(sLineHexCD,iOffSet+54,2) & mid(sLineHexCD,iOffSet+51,2) & mid(sLineHexCD,iOffSet+48,2)
			sCRC32=sLineHex
			'wscript.echo "sCRC32 " & iIdx & " : " & sCRC32 & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+69,2) & mid(sLineHexCD,iOffSet+66,2) & mid(sLineHexCD,iOffSet+63,2) & mid(sLineHexCD,iOffSet+60,2)
			sSizeCompressed=clng(sLineHex)
			'wscript.echo "sSizeCompressed " & iIdx & " : " & sSizeCompressed  & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+81,2) & mid(sLineHexCD,iOffSet+78,2) & mid(sLineHexCD,iOffSet+75,2) & mid(sLineHexCD,iOffSet+72,2)
			sSizeUnCompressed=clng(sLineHex)
			'wscript.echo "sSizeUnCompressed " & iIdx & " : " & sSizeUnCompressed  & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+87,2) & mid(sLineHexCD,iOffSet+84,2)
			sFileNameLength=clng(sLineHex)
			'wscript.echo "sFileNameLength " & iIdx & " : " & sFileNameLength & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+93,2) & mid(sLineHexCD,iOffSet+90,2)
			sExtraFieldLength=clng(sLineHex)
			'wscript.echo "sExtraFieldLength " & iIdx & " : " & sExtraFieldLength  & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+99,2) & mid(sLineHexCD,iOffSet+96,2)
			sFileCommentLength=clng(sLineHex)
			'wscript.echo "sFileCommentLength " & iIdx & " : " & sFileCommentLength  & " (" & sLineHEX  & ")"
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+105,2) & mid(sLineHexCD,iOffSet+102,2)
			iStartDisk=cint(sLineHex)
			'wscript.echo "iStartDisk " & iIdx & " : " & iStartDisk
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+111,2) & mid(sLineHexCD,iOffSet+108,2)
			iFileAttributesInt=HexToBin(sLineHEX)
			'wscript.echo "iFileAttributesInt " & iIdx & " : " & iFileAttributesInt
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+123,2) & mid(sLineHexCD,iOffSet+120,2) & mid(sLineHexCD,iOffSet+117,2) & mid(sLineHexCD,iOffSet+114,2)
			lFileAttributesExt=sLineHex 'clng(sLineHex)
			'wscript.echo "lFileAttributesExt " & iIdx & " : " & lFileAttributesExt
			sLineHEX="&H" & mid(sLineHexCD,iOffSet+135,2) & mid(sLineHexCD,iOffSet+132,2) & mid(sLineHexCD,iOffSet+129,2) & mid(sLineHexCD,iOffSet+126,2)
			lLocFileHeadStart=clng(sLineHex)
			'wscript.echo "lLocFileHeadStart " & iIdx & " : " & lLocFileHeadStart
			sFileName=mid(sLine,iOffsetAsc+46,sFileNameLength)
			'wscript.echo "sFileName " & iIdx & " : " & sFileName
			sExtraField=mid(sLine,iOffsetAsc+46+sFileNameLength,sExtraFieldLength)
			'wscript.echo "sExtraField " & iIdx & " : " & sExtraField
			sFileComment=mid(sLine,iOffsetAsc+46+sFileNameLength+sExtraFieldLength,sFileCommentLength)
			'wscript.echo "sFileComment " & iIdx & " : " & sFileComment
			iOffset=iOffset+((46+sFileNameLength+sExtraFieldLength+sFileCommentLength)*3)
			'wscript.echo "Next start at (hex): " & iOffset
			iOffsetAsc=iOffsetAsc+46+sFileNameLength+sExtraFieldLength+sFileCommentLength
			'wscript.echo "Next start at (asc): " & iOffsetAsc
			'wscript.echo
			sLineOut=cstr(iIdx) & "," & sFileName & "," & sSizeUnCompressed & "," & sLastModDate & " " & sLastModTime
			otsOutput.writeline(sLineOut)
		next
			
	otsSource.close
	
end sub
					
function DecToBin(iDecByte)
	dim sBits,sMod,iRemainder,idx
	if iDecByte=0 then
		sBits="00000000"
	else 
		sBits=""
		iRemainder=iDecByte
		idx=0
		'do while idx<8
		do while iRemainder>0
			idx=idx+1
			sBits=cstr(iRemainder mod 2) & sBits
			iRemainder=int(iRemainder/2)
			if (idx mod 8)=0 then sBits= " " & sBits
		loop
		sBits=right("00000000" & sBits,8)
	end if
	DecToBin=sBits
end function

function HexToBin(HexString)
	dim nBytes,sByte,sBinLocal,idx
	
	nBytes=int(len(HexString)/2)-1
	idx=2
	sBinLocal=""
	idx=0
	do while idx<nBytes
		idx=idx+1
		sByte=mid(HexString,(idx*2)+1,2)
		sBinLocal=sBinLocal & " " & DecToBin(CInt("&H" & sByte))
	loop
	HexToBin=trim(sBinLocal)
end function

function AscToHex(sAsc)
	AscToHex=right("00" & Hex(Asc(sAsc)), 2)
end function

function BinToDec(BinString)
	dim iDec,iNBits,sBinString,iIdx
	
	BinToDec=0
	sBinString=strreverse(trim(BinString))
	iNBits=len(trim(sBinString))

	for iIdx=1 to iNBits
		BinToDec=BinToDec + (cint(mid(sBinString,iIdx,1))*(2^(iIdx-1)))
	next
	
end function

function HexToDosDate(HexString2Bytes)
	dim sHexStringCorr,sBin,iYear,iMonth,iDay
	
	sBin=replace(strreverse(HexToBin(HexString2Bytes))," ", "")
	
	iYear=1980 + BinToDec(strreverse(mid(sBin,10,7)))
	iMonth=BinToDec(strreverse(mid(sBin,6,4)))
	iDay=BinToDec(strreverse(mid(sBin,1,5)))
	
	HexToDosDate=Dateserial(iYear,iMonth,iDay)
	
end function

function HexToDosTime(HexString2Bytes)
	dim sHexStringCorr,sBin,iHour,iMinute,iSecond
	
	sBin=replace(strreverse(HexToBin(HexString2Bytes))," ", "")
	
	iHour=BinToDec(strreverse(mid(sBin,12,5)))
	iMinute=BinToDec(strreverse(mid(sBin,6,6)))
	iSecond=BinToDec(strreverse(mid(sBin,1,5)))
	
	HexToDosTime=TimeSerial(iHour,iMinute,iSecond)
	

end function
