''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'	> Creates a backwards copy of a file
'	> Makes reading the sourcefile from the end to the start possible with filesystemobject textobject
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Credits:
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' [..]\cscript[.exe] [path to script]\Reverse_Asc_File.vbs  ["]Source["] ["]Reversed["] [/Blocksize:#####]
'	Source: Full qualified path to the file for wich reverse is required
'	Reversed: Full qualified path to outputfile (!!! No overwrite check)
'	Blocksize: max number of bytes read by each textobject.read instruction (influence on process time and required resources)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



option explicit
dim sArchive,sArchiveRev,iBlockLen,objFileSystem,objRegExp,ofSource,otsTextStream
dim lngFSize,lngBlocks,iLastLineLen,sTmp(),sTmpRev(),idx,sNonPrintable

with wscript.arguments
	'wscript.echo .unnamed.count
	if .unnamed.count<>2 then
		wscript.echo "Arguments Missing"
		Usage
	else
		sArchive = .unnamed(0)
		sArchiveRev = .unnamed(1)
	end if
	iBlockLen=4096
	if .named.count>0 then
		if .Named.Exists("Blocksize") then
			iBlockLen=.named.item("Blocksize")
		end if
	end if
end with

wscript.echo "Source: " & sArchive
wscript.echo "Reversed file: " & sArchiveRev
wscript.echo "Blocksize: " & cstr(iBlockLen)

sNonPrintable="[^\w\s\d\.\-\+\?\*\{\}\[\]\(\)\^\$\|\\&/]"

Set objFileSystem = CreateObject( "Scripting.FileSystemObject" )
set objRegExp = new regexp
objRegExp.IgnoreCase=-1
objRegExp.Global=-1  
objRegExp.Pattern=sNonPrintable

set ofSource=objFileSystem.getfile(sArchive)
lngFSize=ofSource.Size
lngBlocks=int(lngFSize/iBlockLen)+1
iLastLineLen=lngFSize mod iBlockLen
wscript.Echo "Size: " & cstr(lngFSize)
wscript.Echo "Blocks: " & cstr(lngBlocks)
wscript.Echo "Lenght last Line: " & cstr(iLastLineLen)


' reading
set otsTextStream=objFileSystem.OpenTextFile(sArchive,1,-1)

for idx=0 to lngBlocks - 2
	redim preserve sTmp(idx)
	sTmp(idx)=StrReverse(otsTextStream.read(iBlockLen))
	wscript.echo "sTmp" & cstr(idx) & " : "
	wscript.echo objRegExp.replace(sTmp(idx)," ")
next
redim preserve sTmp(lngBlocks-1)
sTmp(lngBlocks-1)=StrReverse(otsTextStream.read(iLastLineLen))
wscript.echo "sTmp" & cstr(lngBlocks-1) & " : "
wscript.echo objRegExp.replace(sTmp(lngBlocks-1)," ")

otsTextStream.Close

' writing
set otsTextStream=objFileSystem.OpenTextFile(sArchiveRev,2,-1)

for idx=lngBlocks  to 1 step -1
	otsTextStream.write(sTmp(idx-1))
next

set otsTextStream=nothing
set objRegExp=nothing
set objFileSystem=nothing

wscript.quit
	
sub Usage
	dim sMsgUsage
	sMsgUsage=""
	
	sMsgUsage="Reverse_Asc_File.vbs usage: "
	sMsgUsage=sMsgUsage & vbCrLf & "[..]\cscript[.exe] [path to script]\Reverse_Asc_File.vbs  [""]Source[""] [""]Reversed[""] [/Blocksize:#####]"
	sMsgUsage=sMsgUsage & vbCrLf & "	Source: Full qualified path to the file for wich reverse is required"
	sMsgUsage=sMsgUsage & vbCrLf & "	Reversed: Full qualified path to outputfile (!!! No overwrite check)"
	sMsgUsage=sMsgUsage & vbCrLf & "	Blocksize: max number of bytes read by each textsobject.read instruction (influence on process time and required resources)"
	
	wscript.echo sMsgUsage
	wscript.quit(1)
	
end sub
