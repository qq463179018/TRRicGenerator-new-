##############################################
##       SKit BridgeFeed FID Extractor      ##
## Distributed as part of the SKit tool set ##
##     Developed by Simon Chudley (2007)    ##
##         simon.chudley@reuters.com        ##
##############################################
##  This tool parses BFSSLProxy Bridge FID  ##
##     mappings file and generates a SKit   ##
##                FID file.                 ##
##############################################

my $l_file = $ENV{"SkitRoot"}."\\lib\\bridge_fid_mappings.dat";
open(INF, $l_file) or die "Unable to open BridgeFeed FID mappings file '$l_file'!.";
while (<INF>)
   { if ($_ !~ /^\s*#/ and $_ =~ /,(.+?)\s*$/) { print "$1\n"; } }
close(INF);
