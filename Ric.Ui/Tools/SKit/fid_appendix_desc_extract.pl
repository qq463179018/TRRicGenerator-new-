##############################################
##         SKit FID Appendix Parser         ##
## Distributed as part of the SKit tool set ##
##     Developed by Simon Chudley (2007)    ##
##         simon.chudley@reuters.com        ##
##############################################
##  This tool parses the appendix_a file    ##
##     and extracts FID descriptions.       ##
##############################################

if (!(-f $ARGV[0]))
{
   print "Specify appendix_a file as first parameter.\n";
   exit(1);
}

my $l_buffer = "";
my @l_fids = ();
open(INF, $ARGV[0]);
while (<INF>)
{
   chomp;
   if ($_ =~ /^(.+?)\s+(.+)\s+([0-9]+)\s+(.+)\s+(.+)\s([0-9]+)$/) { push(@l_fids, $3); }      
   elsif ($_ =~ m/^!\s*$/ and $l_buffer ne "") 
   {
      $l_buffer =~ s/\"//g;
      foreach (@l_fids) { print $_.";".$l_buffer."\n"; }
      $l_buffer = "";
      @l_fids = ();
   }
   elsif ($_ =~ /^\!\s+(.+)\s*$/) { if ($l_buffer ne "") { $l_buffer .= " "; } $l_buffer .= $1; }
}
close(INF)

