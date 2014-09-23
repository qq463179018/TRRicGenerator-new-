######################
# SKitLet Setup Script
#
# SKit: S. Chudley 2007

# Work out where Perl is installed
my $l_perl = (grep(/perl.bin/i, split(/;/, $ENV{"PATH"})))[0];
if ($l_perl eq "") 
{ 
   print STDERR "Can't locate perl installation?\n"; 
   exit(1); 
}
 
# Finalise location
$l_perl =~ s/\\$//;
$l_perl =~s/\/$//;
$l_perl .= "\\perl.exe";

# Set SKitLet file type
my $l_skit_root = $ENV{"SKitRoot"};
`ftype SkitLet=\"$l_perl\" -I\"$l_skit_root\\lib\" \"%1\" %*`;
if ($? != 0)
{
   print STDERR "Unable to set SKitLet file type!\n"; 
   exit(1);
}

# Set SKitLet file association
`ASSOC .skt=SKitLet`;
if ($? != 0)
{
   print STDERR "Unable to set SKitLet file association!\n"; 
   exit(1);
}

# Update PATHEXT if not already set
if ($ENV{"PATHEXT"} !~ m/SKT/i) 
{
   `\"$l_skit_root\\bin\\setx\" PATHEXT \"%PATHEXT%\";.SKT`;
   if ($? != 0)
   {
      print STDERR "Unable to update PATHEXT!\n"; 
      exit(1);
   }
}

# Update PATH if not already set
if ($ENV{"PATH"} !~ m/SKitLets/i) 
{
   `\"$l_skit_root\\bin\\setx\" PATH \"%PATH%;$l_skit_root\\SKitLets\"`;
   if ($? != 0)
   {
      print STDERR "Unable to update PATH!\n"; 
      exit(1);
   }
}
