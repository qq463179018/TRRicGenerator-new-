REM SKit Version Dump - Simon Chudley (simon.chudley@reuters.com) 

REM Output Tool versions
output_ver("Tick2XML")
output_ver("TickRelate")
output_ver("TASExtract")
output_ver("FeedProxy")
output_ver("BFSSLProxy")

REM Outputs the version of a SKit app
function output_ver(a_app)

   set l_objShell = WScript.CreateObject("WScript.Shell")
   set l_objExecObject = l_objShell.Exec(a_app + " -version")
   do while not l_objExecObject.StdOut.AtEndOfStream
      l_strText = l_objExecObject.StdOut.ReadLine()
      if Instr(l_strText, "Developed by") > 0 Then
         wscript.echo Left(l_strText, Instr(l_strText, ":") - 1)
         exit do
      end if
   loop
end function
