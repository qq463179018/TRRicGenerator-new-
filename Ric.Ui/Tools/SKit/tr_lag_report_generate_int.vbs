REM SKit TickRelate: Generate lag report file from TickRelate output

if wscript.arguments.count = 0 then
  wscript.stderr.write "Input TickRelate XML lag report not found, specify as first argument."
  wscript.quit
end if

on error resume next 
set l_fsObject = Wscript.CreateObject ("Scripting.FileSystemObject")
Set l_oFile = l_fsObject.GetFile(wscript.arguments(0))
if Err.Number <> 0 then
  wscript.stderr.write "Input TickRelate XML lag report not found, specify as first argument."
  wscript.quit
end if

set l_file = l_fsObject.OpenTextFile(wscript.arguments(0))
wscript.echo "<?xml-stylesheet type='text/xsl' href='tickrelate_lag_report.xsl'?>"
wscript.echo "<LagReport>"
do while l_file.AtEndOfStream <> True
   wscript.echo l_file.ReadLine
loop
wscript.echo "</LagReport>"
