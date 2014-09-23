REM SKit Feed Check Tool - Simon Chudley (simon.chudley@reuters.com) 

REM Check feed connectivity on QDN and IDN
check_feed("QDN")
check_feed("IDN")

REM Checks whether a given FID for a RIC can be fetched from the specified network
function check_connectivity(a_feed, a_type, a_ric, a_fid)

   set l_objShell = WScript.CreateObject("WScript.Shell")
   set l_objExecObject = l_objShell.Exec("Tick2XML -o " + a_feed + "_" + a_type + ".cfg -rics " + a_ric + " -dbout -fids " + a_fid + " -quiet -no_feed_timeout -timeout 15 -dbout_inactives")
   check_connectivity = 0
   do while not l_objExecObject.StdOut.AtEndOfStream
       l_strText = l_objExecObject.StdOut.ReadLine()
       if (Instr(l_strText, a_fid) > 0) or (Instr(l_strText, "INACTIVE") > 0) Then
           check_connectivity = 1
           exit do
       end if
   loop

end function


REM Checks a given feed for connectivity, and details what may be wrong if no data is received
function check_feed(a_feed)

   REM Check for connectivity
   l_feed = lcase(a_feed)
   wscript.echo "Checking " + a_feed + " connectivity:"
   l_presult = check_connectivity(a_feed, "primary", "RTR.L", "DSPLY_NAME")
   l_sresult = check_connectivity(a_feed, "secondary", "RTR.L", "DSPLY_NAME")
   if l_presult = 0 or l_sresult = 0 then
      wscript.echo ""
      wscript.echo "  Failed to fetch data from " + a_feed + " feed!"
      wscript.echo ""
      if l_presult = 0 then
         wscript.echo "     The current configuration for " + a_feed + " as a primary feed (in skit/etc/" + l_feed + "_primary.cfg) is:"
         wscript.echo ""
         output_configuration a_feed, "primary"
         wscript.echo ""
      end if
      if l_sresult = 0 then
         wscript.echo "     The current configuration for " + a_feed + " as a secondary feed (in skit/etc/" + l_feed + "_secondary.cfg) is:"
         wscript.echo ""
         output_configuration a_feed, "secondary"
         wscript.echo ""
      end if
      wscript.echo "     1) Ensure the above configuration is correct to connect to your team's designated " + a_feed + " data provider." 
      wscript.echo ""
      wscript.echo "        If not, modify the file(s) listed above to contain the correct configuration. "
      wscript.echo ""
      wscript.echo "     2) Ensure you have connectivity through to the host(s)."
      wscript.echo ""
      wscript.echo "     3) Ensure you are using the correct " + a_feed + " user account(s). "
      wscript.echo ""
      wscript.echo "        By default, it will use the username you are logged into Windows with. If you wish to "
      wscript.echo "        use a different account, specify the username in the relevant files above using -pu in "
      wscript.echo "        " + l_feed + "_primary.cfg and -su in " + l_feed + "_secondary.cfg."
      wscript.echo ""
   else
      wscript.echo ""
      wscript.echo "  All looks good."
      wscript.echo ""
   end if

end function

REM Outputs the configuration file being used for a given feed
function output_configuration(a_feed, a_type)

   l_feed = lcase(a_feed)
   const l_ForReading = 1
   set l_wsh = CreateObject("WScript.Shell")
   set l_objFSO = CreateObject("Scripting.FileSystemObject")
   set l_objFile = l_objFSO.OpenTextFile(l_wsh.ExpandEnvironmentStrings("%SKitRoot%") + "\etc\" + l_feed + "_" + a_type + ".cfg", l_ForReading)
   do while l_objFile.AtEndOfStream = False
      l_strLine = Trim(l_objFile.ReadLine)
      if not Left(l_strLine, 1) = "#" and Len(l_strLine) > 0 then
         wscript.echo "        " + l_strLine
      end if
   loop

end function
