#
#       This document contains information proprietary to
#       Reuters Limited and  may not be reproduced and/or
#       used in whole or  in  part  without  the  express
#       written  permission   of   Reuters  Limited.
#
#               Copyright 2003 Reuters Limited.
#
# ------------------------------------------------------------
#
# Title         : dbout_diff.pl
# Author        : Simon Chudley
# Creation Date : Fed 2006 (SKit)
#
# Description
# -----------
# Performs a DIFF of two DBOutput files, and intelligently tries to work out whats
# different between the two.
#
# This tool is part of the SKit data validation tool set.
#
#   http://pdt.uki.ime.reuters.com/teams/core/srcrtr.asp?a_page=skit.php
#
# Change History
# --------------
# 21-Feb-2006  S.R. Chudley      SKit : Created
# 17-Aug-2006  S.R. Chudley      SKit : Added support for acronym mapping and ric mangling
# 13-Dec-2006  S.R. Chudley      SKit : Added option to specify FIDs to compare
# 11-May-2007  S.R. Chudley      SKit : Fixed bug when comparing numeric vs blank text field
# 22-May-2007  S.R. Chudley      SKit : Performance & memory usage improvement, field tolerances

# Version
my $version = "1.1.6.0";

use POSIX qw(strftime);
use Time::Local;

# Required parameters
my $file_a = $ARGV[0];
my $file_b = $ARGV[1];

if ($file_a eq "-version") 
{
   print "DBOut_Diff v".$version.": Developed by S. Chudley\n";
   exit(0);
}

# Tool usage
if ($file_a eq "" or $file_b eq "" or !(-f $file_a) or !(-f $file_b))
{
    print STDERR "DBOut_Diff v".$version.": Developed by S. Chudley\n\n";
    print STDERR "Usage:\n\n   dbout_diff.pl <file_a> <file_b> <options>\n\n";
    if ($file_a eq "-help" or $file_a eq "/help" or $file_a eq "-?")
    {
        print STDERR "Optional (FID Filtering):\n\n";
        print STDERR "\t-fids:x,y,z           : Only compare FIDs given in the comma separated list\n";
        print STDERR "\t-fidsfile:file        : File containing list of FIDs to compare\n";
        print STDERR "\t-ignorefids:x,y,z     : Ignore any FID given in the comma separated list\n";
        print STDERR "\t-ignorefidslike:x,y,z : Ignore FIDs that match any of the regular expressions in the comma separated list\n";
        print STDERR "\nOptional (Missing RIC/FIDs):\n\n";
        print STDERR "\t-imfids               : Ignore differences due to missing FIDS\n";
        print STDERR "\t-imfids_a             : Ignore differences due to missing FIDS within the A file\n";
        print STDERR "\t-imfids_b             : Ignore differences due to missing FIDS within the B file\n";
        print STDERR "\t-imrics               : Ignore differences due to missing RICs\n";
        print STDERR "\nOptional (Output Formatting):\n\n";
        print STDERR "\t-colwidth:width       : Sets the column width for the outputs\n";
        print STDERR "\t-csv                  : Produce output in CSV format\n";
        print STDERR "\t-tee:file             : Tee output to specified file\n";
        print STDERR "\t-tee_append           : Append data to end of tee file rather than truncating\n";
        print STDERR "\t-isolaterics          : Isolate differences on individual RICs by surrounding them by blank lines\n";
        print STDERR "\nOptional (Tool):\n\n";
        print STDERR "\t-options:file         : Read in options from a file\n";
        print STDERR "\t-o:file               : Alias for above\n";
        print STDERR "\t-manglerule:sub-regex : Mangling regular subs exp from B RIC to A RIC (eg. ^x(.+)/\$1)\n";
        print STDERR "\t-acrmap:file          : Define acronym mappings/translations to make prior to compare\n";
        print STDERR "\t-seg_count:rics       : Number of RICs to read in during each segment comparison\n";
        print STDERR "\t-trim_fids            : Trim whitespace from he front and end of FID values before comparing\n";
        print STDERR "\t-no_footer            : Don't append footer summary\n";
        print STDERR "\t-version              : Output version identifier\n";
    }
    # Open input?
    elsif (($file_a ne "" and !(-f $file_a)) or
           ($file_b ne "" and !(-f $file_b)))
    {
       print STDERR "Unable to open both input files?\n";
       exit(1);
    }
    else
    {
       print STDERR "Specify -? or -help to display details on all supported options.\n";
    }

    exit(1);
}

# State
my %rics_a = ();
my %rics_b = ();
$ric_diff_count = 0;
$fid_diff_count = 0;
$miss_ric_count = 0;
$ric_count = 0;



###################
# PROCESS OPTIONS #
###################

# Default options
my $col_width = 60;
my $seg_ric_count = 100;
my $isolate_diffs = 0;
my $imfids = 0;
my $imfids_a = 0;
my $imfids_b = 0;
my $imrics = 0;
my $csv = 0;
my %ignore_fids = ();
my %ignore_fids_like = ();
my $mangle_rule = "";
my %acrmap = ();
my %comp_fids = ();
my $tee = "";
my $tee_append = 0;
my $no_footer = 0;
my $trim_fids = 0;

# Get options from command line
my @l_options = ();
for (my $i = 2;$i < @ARGV;$i++) { push(@l_options, $ARGV[$i]); }

# Split out options
foreach my $option (@l_options)
{
   if ($option =~ m/^(\/|-)(.+?)(:.+|)$/)
   {
       my $opt = lc($2);
       my $val = $3; $val =~ s/://g;
       if ($opt eq "imrics") { $imrics = 1; }
       elsif ($opt eq "imfids") { $imfids = 1; }
       elsif ($opt eq "imfids_a") { $imfids_a = 1; }
       elsif ($opt eq "imfids_b") { $imfids_b = 1; }
       elsif ($opt eq "colwidth") { $col_width = $val; }
       elsif ($opt eq "no_footer") { $no_footer = 1; }
       elsif ($opt eq "isolaterics") { $isolate_diffs = 1; }
       elsif ($opt eq "seg_count") { $seg_ric_count = $val; }
       elsif ($opt eq "csv") { $csv = 1; }
       elsif ($opt eq "tee") { $tee = $val; }
       elsif ($opt eq "tee_append") { $tee_append = 1; }
       elsif ($opt eq "trim_fids") { $trim_fids = 1; }
       elsif ($opt eq "manglerule") { $manglerule = $val; }
       elsif ($opt eq "version") { print "DBOut_Diff v".$version.": Developed by S. Chudley\n\n"; }
       elsif ($opt eq "options" or $opt eq "o") 
       {
          # Check options file 
          if (!-f $val)
          {
             print STDERR "Unable to read options file '$val'!\n";
             exit(1);
          }

          # Read options
          my $l_opt_str = "";
          open(INF, $val);
          while (<INF>) { if ($_ !~ m/^\s*#/) { my $l_data = $_; $l_data =~ s/#.*$//; push(@l_options, split(/\s+/, $l_data)); } }
          close(INF);
       }
       elsif ($opt eq "fids") 
       { 
           my @fids = split(/,/, $val);
           foreach my $fid (@fids)
           {
               $fid =~ s/^\s+//; $fid =~ s/\s+$//; 
               $fid = uc($fid);
               $comp_fids{$fid} = 1;
           }
       }
       elsif ($opt eq "fidsfile") 
       { 
          if (! -f $val)
          {
             print STDERR "Unable to open FIDs file '$val'!\n";
             exit(1);
          }

          open(INF, $val);
          while (<INF>)
          {
             my $l_line = $_;
             if ($l_line =~ m/^\s*(#|!)/) { next; }
             my @fids = split(/,|\s/, $l_line);
             foreach my $fid (@fids)
             {
                 $fid =~ s/^\s+//; $fid =~ s/\s+$//; 
                 $fid =~ s/\s.*$//g;
                 $fid = uc($fid);
                 if ($fid =~ m/^[A-Z][A-Z0-9_]*$/i) { $comp_fids{$fid} = 1; }
             }
          }
          close(INF);
       }
       elsif ($opt eq "ignorefids") 
       { 
           my @fids = split(/,/, $val);
           foreach my $fid (@fids)
           {
               $fid =~ s/^\s+//; $fid =~ s/\s+$//; 
               $fid = uc($fid);
               $ignore_fids{$fid} = 1;
           }
       }
       elsif ($opt eq "ignorefidslike") 
       { 
           my @fids = split(/,/, $val);
           foreach my $fid (@fids)
           {
               $fid =~ s/^\s+//; $fid =~ s/\s+$//; 
               $fid = uc($fid);
               push(@ignore_fids_like, $fid);
           }
       }
       elsif ($opt eq "acrmap")
       {
          if (! -f $val)
          {
             print STDERR "Unable to open acronym mapping file '$val'!\n";
             exit(1);
          }

          my @l_with = ();
          open(INF, $val);
          while (<INF>)
          {
             if ($_ =~ m/^\s*#/) { next; }
             elsif ($_ =~ m/^\s*with\s*(.+)\s*$/i)
                { @l_with = split(/,/, $1); }
             elsif ($_ =~ m/^\s*(.+?)\s*=\s*(.+)\s*$/)
             {
                my $val_a = $1;
                my $val_b = $2;
                foreach my $l_acr (@l_with)
                   { $l_acr =~ s/\s//g; $acrmap{$l_acr}{"values"}{$val_a} = $val_b; }
             }
             elsif ($_ =~ m/^\s*limit_size\s+([0-9]+)\s*$/) 
             {
                my $val = $1;
                foreach my $l_acr (@l_with)
                   { $l_acr =~ s/\s//g; $acrmap{$l_acr}{"limit_size"} = $val; }
             }
             elsif ($_ =~ m/^\s*numeric_tolerance\s+([0-9]+)\s*$/) 
             {
                my $val = $1;
                foreach my $l_acr (@l_with)
                   { $l_acr =~ s/\s//g; $acrmap{$l_acr}{"numeric_tolerance"} = $val; }
             }
             elsif ($_ =~ m/^\s*time_tolerance\s+([0-9:]+)\s*$/) 
             {
                my $val = $1;
                foreach my $l_acr (@l_with)
                   { $l_acr =~ s/\s//g; $acrmap{$l_acr}{"time_tolerance"} = $val; }
             }
          }
          close(INF);
       }
   }
}

# Tee?
if ($tee ne "")
{
   my $l_file = ">".$tee;
   if ($tee_append == 1) { $l_file = ">".$l_file; }
   if (!open(TEE, $l_file))
   {
      print STDERR "Unable to output output file '$tee'!\n";
      exit(1);
   }
}

# Output formatting
if ($col_width < 31) { $col_width = 31; }
my $output_format = " %-".$col_width."s | %-".$col_width."s \n";
my $ent_format = "%-19s %-12s %-".$l_width."s";

# Header
if ($csv == 0) { &logmsg(sprintf($output_format, $file_a, $file_b)); }
else { &logmsg("$file_a,,,$file_b\n"); }

my $dash = "";
for (my $i=0;$i < $col_width;$i++) { $dash .= "-"; }
if ($csv == 0) { &logmsg(sprintf($output_format, $dash, $dash)); } 



#######################
# ANALYSE DIFFERENCES #
#######################

# Open input files
open(FILE_A, $file_a);
open(FILE_B, $file_b);

# While we still have data to read
my $l_start = timelocal(gmtime());
while (&read_segment() > 0)
{
    # For all RICs in file segment A
    my @l_keys = sort keys %rics_a;
    foreach my $ric (@l_keys)
    {
        # If RIC is present in file segment B
        my $diff = 0;
        if (defined $rics_b{$ric})
        {
            # Fetch RIC FIDs 
            my %fids_a = %{$rics_a{$ric}};
            my %fids_b = %{$rics_b{$ric}};
            $ric_count++;

            # Compare each FID in file A for this RIC
            foreach my $fid (sort keys %fids_a)
            {
                my $val_a = $fids_a{$fid};
                my $val_b = $fids_b{$fid};
                if (defined $val_b)
                {
                    my $numeric = 0;
                    if ($val_a =~ m/^(\+|-){0,1}[\.0-9]+$/ and $val_b =~ m/^(\+|-){0,1}[\.0-9]+$/) 
                       { $numeric = 1; $val_a =~ s/\+//; $val_b =~ s/\+//; } 

                    my $l_is_diff = 0;
                    if ($numeric == 1)
                    {
                        # Numeric tolerances
                        if (defined $acrmap{$fid}{"numeric_tolerance"})
                        {
                            my $l_tol = $acrmap{$fid}{"numeric_tolerance"};
                            $l_is_diff = ($val_b < $val_a - $l_tol) || ($val_b > $val_a + $l_tol);
                        }
                        else { $l_is_diff = $val_a != $val_b; }
                    }
                    else
                    {
                        # Time based tolernaces
                        if (defined $acrmap{$fid}{"time_tolerance"} and
                            $val_a =~ m/^[0-9]+:[0-9]+/ and 
                            $val_b =~ m/^[0-9]+:[0-9]+/)
                        {
                            my $l_tol = &time_to_sec($acrmap{$fid}{"time_tolerance"});
                            my $l_val_a = &time_to_sec($val_a);
                            my $l_val_b = &time_to_sec($val_b);
                            $l_is_diff = ($l_val_b < $l_val_a - $l_tol) || ($l_val_b > $l_val_a + $l_tol);
                        }
                        else { $l_is_diff = $val_a ne $val_b; }
                    }

                    if ($l_is_diff == 1)
                       { $diff += &output_row($ric, $fid, $val_a, $val_b); }
                }
                else
                    { if ($imfids == 0 && $imfids_b == 0) { $diff += &output_row($ric, $fid, $val_a, "** NOT DEFINED **"); } }
            }
    
            # Check if the B file has any FIDs for this RIC that are not in the A file
            foreach my $fid (sort keys %fids_b)
            {
                my $val_a = $fids_a{$fid};
                my $val_b = $fids_b{$fid};
                if (!defined $val_a)
                    { if ($imfids == 0 && $imfids_a == 0) { $diff += &output_row($ric, $fid, "** NOT DEFINED **", $val_b); } }
            }
    
            # Remove RIC from A and B hash
            delete($rics_a{$ric});
            delete($rics_b{$ric});
        }
            
        # RIC separation
        if ($isolate_diffs == 1 and $diff >= 1) { if ($csv == 0) { &logmsg(sprintf($output_format, $dash, $dash)); } }
        if ($diff >= 1) { $ric_diff_count++; }
    }
}

# Output any RICs that weren't present in the A file
foreach my $ric (sort keys %rics_a)
{
    $ric_count++;
    if ($imrics == 0)
    {
        if ($csv == 0) { &logmsg(sprintf($output_format, &format_col($ric, "", ""), "** RIC NOT DEFINED **")); }
        else { &logmsg("$ric,,,** RIC NOT DEFINED **\n"); }
        $miss_ric_count++; $ric_diff_count++;
        if ($isolate_diffs == 1) { if ($csv == 0) { &logmsg(sprintf($output_format, $dash, $dash)); } }
    }
}

# Output any RICs that weren't present in the B file
foreach my $ric (sort keys %rics_b)
{
    $ric_count++;
    if ($imrics == 0)
    {
        if ($csv == 0) { &logmsg(sprintf($output_format, "** RIC NOT DEFINED **", &format_col($ric,  "", ""))); }
        else { &logmsg("** RIC NOT DEFINED **,,,$ric\n"); }
        $miss_ric_count++; $ric_diff_count++;
        if ($isolate_diffs == 1) { if ($csv == 0) { &logmsg(sprintf($output_format, $dash, $dash)); } }
    }
}

# Close input files
close(FILE_A);
close(FILE_B);

# Summary
if ($no_footer != 1)
{
    my $l_end = timelocal(gmtime());
    my $l_time = $l_end - $l_start; if ($l_time == 0) { $l_time = 1; }
    my $l_rate = $ric_count / $l_time; $l_rate =~ s/(\...).+$/$1/;
    &logmsg("\nCompared $ric_count RICS of which $ric_diff_count were different; $fid_diff_count unique FID difference(s); $miss_ric_count RICs missing from the input files.\n"); 
    &logmsg("\nComparison took ".strftime("%H:%M:%S", gmtime($l_end - $l_start))." at a rate of ".$l_rate." RICs/second.\n");
}
if ($tee ne "") { close(TEE); }



####################
# HELPER FUNCTIONS #
####################

# Formats a column entry
sub format_col
{
    my ($ric, $fid, $val) = @_;
    my $entry = sprintf($ent_format, $ric, $fid, $val); 
    if (length($entry) > $col_width) { $entry = substr($entry, 0, $col_width); }
    return $entry;
}

# Formats an output row
sub output_row
{
    my ($ric, $fid, $val_a, $val_b) = @_;

    if ($csv == 0) { &logmsg(sprintf($output_format, &format_col($ric, $fid, $val_a), &format_col($ric, $fid, $val_b))); }
    else { &logmsg("$ric,$fid,$val_a,$ric,$fid,$val_b\n"); }
    $fid_diff_count++; 
    return 1;
}

# Acronym mapping function
sub acronym_mapping
{
    local ($acr, *val) = @_;
    if (defined $acrmap{$acr})
    {
        if (defined $acrmap{$acr}{"limit_size"} and length($val) > $acrmap{$acr}{"limit_size"})
           { $val = substr($val, 0, $acrmap{$acr}{"limit_size"}); }
        if (defined $acrmap{$acr}{"values"}{$val}) { $val = $acrmap{$acr}{"values"}{$val}; }
    }
}

# Want to include this FID?
sub select_fid
{
    my $a_fid = shift;
    if  (!defined $ignore_fids{$a_fid} && (keys %comp_fids == 0 || defined $comp_fids{$a_fid}))
    {
        # Want to exclude FIDs like this?
        foreach my $l_fid_regexp (@ignore_fids_like)
            { if ($a_fid =~ m/$l_fid_regexp/) { return 0; } }

         # Include this FID
         return 1;
    }

    # Exclude this FID
    return 0;
}

# Reads a segment of RICs from both files
sub read_segment
{
    # Read from both files
    my $l_count = 0;
    $l_count += &read_rics(FILE_A, \%rics_a, 0);
    $l_count += &read_rics(FILE_B, \%rics_b, 1);
    return $l_count;
}

# Reads all FIDs for a set number of RICs from a single file
sub read_rics
{
    local (*a_fd, *a_hash, $a_mangle) = @_;

    my $l_ric_count = 0; 
    my $l_last_ric = ""; 
    my $l_last_pos = 0;
    while (<a_fd>)
    {
        if ($_ =~ m/^(.+?)\s+(.+?)(\s+.+?\s*){0,1}$/ and $_ !~ m/^!/) 
        { 
            my $ric = $1;
            my $acr = uc($2); $acr =~ s/\s//g;
            chomp(my $val = $3); if (!defined $val) { $val = ""; } $val =~ s/^\s//;
            if ($trim_fids == 1 or $val =~ m/^\s+$/) { $val =~ s/^\s+//; $val =~ s/\s+$//; }

            # Hit RIC count?
            if ($l_last_ric ne $ric and $l_ric_count++ == $seg_ric_count) 
            {
               # Process that data later
               seek(a_fd, $l_last_pos, 0);
               last;
            }
            $l_last_ric = $ric;

            # RIC mangling 
            if ($a_mangle == 1 && $manglerule)
            {
               my ($src, $dst) = split(/\//, $manglerule);
               eval "\$ric =~ s/$src/$dst/";
            }

            # Store this FID
            if (&select_fid($acr) == 0) { next; }
            &acronym_mapping($acr, \$val);
            $a_hash{$ric}{$acr} = $val; 
        }
        
        $l_last_pos = tell(a_fd);
    } 

    return $l_ric_count;
}

# Log a message to STDOUT and tee file 
sub logmsg
{
   my $a_msg = shift;
   print $a_msg;
   if ($tee ne "") { print TEE $a_msg; }
}

# Converts a time to a number of seconds
sub time_to_sec
{
    my @l_arr = split(/:/, shift);
    return $l_arr[0] * 60 * 60 + $l_arr[1] * 60 + $l_arr[2];
}
