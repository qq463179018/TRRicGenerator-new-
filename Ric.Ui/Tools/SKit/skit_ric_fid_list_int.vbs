REM SKit: Converts a flat list of RICs or FIDs into a list parsable by the SKit tools

if wscript.arguments.count = 0 then
  wscript.stderr.write "Input RIC/FID list file not found, specify as first argument."
  wscript.quit
end if

on error resume next
set l_fso = CreateObject("Scripting.FileSystemObject")
Set l_fo = l_fso.GetFile(wscript.arguments(0))
if Err.Number <> 0 then
  wscript.stderr.write "Input RIC/FID list file not found, specify as first argument."
  wscript.quit
end if

set l_file = l_fso.OpenTextFile(wscript.arguments(0))
do while l_file.AtEndOfStream <> True
  l_strText = l_file.ReadLine
  if Len(l_strText) > 0 then
    if wscript.arguments(1) = "rics" then
       wscript.echo "-rics " + l_strText
     else
       wscript.echo "-fids " + l_strText
     end if
 end if
loop
