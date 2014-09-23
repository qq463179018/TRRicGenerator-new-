# SKit: TickRelate/Tick2XML Log Slicer
# S. Chudley 2007

use POSIX qw(strftime);
use Time::Local;

# Options
my $l_file = $ARGV[0];
my $l_start = 0;
my $l_finish = 0;
my %l_all_rics = ();
my %l_all_types = ();
my $l_strip_timestamps = 0;

if (! -f $l_file)
{
   print STDERR "TickRelate/Tick2XML Log Slice\n\n";
   print STDERR "This utility allows you to slice up a TickRelate/Tick2XML log file based on RICs, time and update type.\n\n";
   print STDERR "Usage:\n\n   skit_log_slice <logfile> <options>\n\n";
   print STDERR "Options:\n";
   print STDERR "\t-start:<time>        : Log messages from time only\n";
   print STDERR "\t-finish:<time>       : Log messages until time only\n";
   print STDERR "\t-rics:<name,name>    : Log messages for given RIC(s) only\n";
   print STDERR "\t-types:<type,type>   : Log messages for given message types only\n";
   print STDERR "\t-strip_timestamps    : Strip all time/dates from output\n\n";

   print "Unable to open input XML log file given as first argument: '".$l_file."'\n";
   exit(1);
}

# Split options
foreach my $l_option (@ARGV)
{
   if ($l_option =~ m/^(\/|-)(.+?)(:.+|)$/)
   {
       my $l_opt = lc($2);
       my $l_val = $3; $l_val =~ s/://;
       if ($l_opt eq "start") { $l_start = &get_time($l_val); }
       elsif ($l_opt eq "finish") { $l_finish = &get_time($l_val); }
       elsif ($l_opt eq "rics")
       {
           my @l_rics = split(/,/, $l_val);
           foreach my $l_ric (@l_rics)
           {
               $l_ric =~ s/^\s+//; $l_ric =~ s/\s+$//;
               $l_all_rics{$l_ric} = 1;
           }
       }
       elsif ($l_opt eq "types")
       {
           my @l_types = split(/,/, $l_val);
           foreach my $l_type (@l_types)
           {
               $l_type =~ s/^\s+//; $l_type =~ s/\s+$//;
               $l_all_types{$l_type} = 1;
           }
       }
       elsif ($l_opt eq "strip_timestamps") { $l_strip_timestamps = 1; }
   }
}

print "<?xml-stylesheet type=\"text/xsl\" href=\"".$ENV{"SKitRoot"}."\\xslt\\skit_log_render.xsl\"?>\n";
print "<TickRelateLog>\n";

# Process file
my $l_tag_filter = ""; 
my $l_time_filter = ""; 
open(INF, $l_file);
while (<INF>)
{
   # Strip timestamps
   my $l_line = $_;
   if ($l_strip_timestamps)
   {
      $l_line =~ s/timestamp=".+?"//g;
      $l_line =~ s/time=".+?"//g;
      $l_line =~ s/date=".+?"//g;
   }

   # Time filtering
   if ($l_time_filter eq "" and $_ =~ m/^<(.+?)\s.*\s(time|timestamp)="(.+?)"/ and ($l_start != 0 or $l_finish != 0))
   {
      my $l_log_time = &get_time($3);
      if (($l_start != 0 and $l_log_time < $l_start) or ($l_finish != 0 and $l_log_time > $l_finish))
      {
         if ($_ !~ m/\/>$/) { $l_time_filter = $1; }
         next;
      }
   }
   elsif ($l_time_filter ne "" and $_ =~ m/<\/$l_time_filter>/) { $l_time_filter = ""; next; }
   elsif ($l_time_filter ne "") { next; }

   # Log message type filtering
   if ($l_time_filter eq "" and $_ =~ m/^<(.+?)\s.+/ and (keys %l_all_types > 0) and !defined $l_all_types{$1}) 
   {
      if ($_ !~ m/\/>$/) { $l_time_filter = $1; }
      next;
   }

   # RIC filtering
   if ($l_tag_filter eq "" and $_ =~ m/^<(.+?)\s.*(primary|ric)="(.+?)"/ and defined $l_all_rics{$3})
   {
      $l_tag_filter = $1;
      print $l_line;
      if ($_ =~ m/\/>\s*$/) {print "NNN\n"; $l_tag_filter = ""; }
   }
   elsif ($l_tag_filter ne "" and $_ =~ m/<\/$l_tag_filter>/) { $l_tag_filter = ""; print $l_line; }
   elsif (((keys %l_all_rics) == 0 or $l_tag_filter ne "") and $_ !~ /^\s*$/) { print $l_line; }
}
close(INF);

print "</TickRelateLog>\n";

# Works out time from user string
sub get_time
{
   my $a_time = shift;
   $a_time =~ m/([0-9]+):([0-9]+):([0-9]+)/;
   return timelocal($3, $2, $1, (gmtime)[3,4,5]);
}
