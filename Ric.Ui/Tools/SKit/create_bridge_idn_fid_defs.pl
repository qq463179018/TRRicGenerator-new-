##############################################
##         SKit BDN FID List Creator        ##
## Distributed as part of the SKit tool set ##
##     Developed by Simon Chudley (2008)    ##
##         simon.chudley@reuters.com        ##
##############################################
##   This tool creates an IDN appendix_a    ##
##    file containing all BDN elements.     ##
##############################################

open(APP, ">local_appendix_a");
open(MAP, ">local_bridge_fid_mappings.dat");
open(ENUM, ">local_enumtype.def");
my $l_id = 1;
open(INF, $ENV{"SkitRoot"}."\\lib\\bridge_element_dictionary.dat");
while (<INF>)
{
   if ($_ =~ /^\s*([0-9]+)\s+BF_TE_(.+?)\s+(.+?)\(([0-9]+)\)\s+([0-9]+)\s+/)
   {
      my $l_fid = $2;
      my $l_typ = $3;
      my $l_siz = $5;

      # Work out IDN type
      my $l_type = "";
      my $l_size = 0;
      if ($l_typ eq "bfPrice" or 
          $l_typ eq "bfReal32" or 
          $l_typ eq "bfReal64") { $l_type = "PRICE"; $l_size = 17; }
      elsif ($l_typ eq "bfTime") { $l_type = "TIME"; $l_size = 5; }
      elsif ($l_typ eq "bfScalar") { $l_type = "INTEGER"; $l_size = 15; }
      elsif ($l_typ eq "bfDate") { $l_type = "DATE"; $l_size = 11; }
      elsif ($l_typ eq "bfString" or
             $l_typ eq "bfChar" or
             $l_typ eq "bfUTF8String") { $l_type = "ALPHANUMERIC"; $l_size = $l_siz; }

      if ($l_size != 0)
      {
         print APP sprintf("%-50s %-50s %5d NULL %-15s %3d\n",
                           $l_fid,
                           "\"".$l_fid."\"",
                           $l_id++,
                           $l_type,
                           $l_size);
         print MAP "BF_TE_$l_fid,$l_fid\n";
      }
      else { print STDERR "No type for '$l_fid' of type '$l_typ'.\n"; }

   }
}
close(INF);
close(APP);
close(MAP);
close(ENUM);

print "\nDone. Created local_appendix_a, local_bridge_fid_mappings.dat and local_enumtype.def.\n";
