# SKit: Tick2XML XML to DBOUT converter
# S. Chudley 2007

if (! -f $ARGV[0])
{
   print "Unable to open input XML snap-shot file given as first argument: '".$ARGV[0]."'\n";
   exit(1);
}

print "KEY=RIC\n";
my $l_ric = "";
open(INF, $ARGV[0]);
while (<INF>)
{
   if ($_ =~ m/^<SnapShot ric="(.+?)"/) { $l_ric = $1; }
   elsif ($_ =~ m/^\s*<(.+?)>(.+?)<\// and $l_ric ne "")
      { print sprintf("%-18s %-18s %s\n", $l_ric, $1, $2); }
}
close(INF);
