chomp(my $l_ver = `$ARGV[0] -version 2>nul`);
if ($l_ver =~ m/^(.+v.+):/)
{
   print $1."\n";
   exit(0);
}
else { exit(1); }
