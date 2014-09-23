##############################################
##          SKit FID List Generator         ##
## Distributed as part of the SKit tool set ##
##     Developed by Simon Chudley (2007)    ##
##         simon.chudley@reuters.com        ##
##############################################
##     This tool converts an IDN .OUPUT     ##
##    template file into a SKit FID list.   ##
##############################################

my $g_input = $ARGV[0];
my $r_s_str = "";
my $r_s_num = 0;
my $l_str = "";
my $l_num = 0; 

# Usage
if ($g_input eq "" or !(-f $g_input))
{
   print STDERR "Usage:\n\n   template_to_fid_list.pl <idn_template>.output\n\n";
   print STDERR "   Tool expects an IDN template OUTPUT file as first parameter.\n";
   exit(1);
}

# Process file
open(INF, $g_input);
while(<INF>)
{
   if ($_ =~ m/^([A-Z0-9_]+)\s+([0-9]+)\s+/)
   {
      my $acr = $1;
      my $fid = $2;

      if ($r_s_str eq "") { if ($acr ne "DSO_ID") { $r_s_str = $acr; $r_s_num = $fid; } }
      else
      {
         if ($l_num + 1 != $fid or $acr eq "DSO_ID")
         {
           # End of range
           if ($r_s_str ne $l_str) { print "-fids $r_s_str-$l_str\n"; } else { print "-fids $l_str\n"; }
           if ($acr ne "DSO_ID") { $r_s_str = $acr; $r_s_num = $fid; } else { $r_s_str = $r_s_num = ""; }  
         }
      }
      $l_str = $acr;
      $l_num = $fid;
   }
}
close(INF);

if ($r_s_str ne "") { if ($r_s_str ne $l_str) { print "-fids $r_s_str-$l_str\n"; } else { print "-fids $r_s_str\n"; } }
