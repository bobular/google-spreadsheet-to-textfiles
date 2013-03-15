#!/usr/bin/perl
#
#
# USAGE: ./google-spreadsheet-to-textfiles.pl -outdir my-dir -username myemail@gmail.com -title "Spreadsheet title"
#
# Prompts for your google password, accesses the desired spreadsheet
# and writes a text file for each worksheet in the directory specified.
#
# OPTIONS:
#
#   -suffix .tsv (default is .txt)
#   -delimiter , (default is tab)
#
# INSTALLATION NOTES:
#
# Net::Google::Spreadsheets build may fail with a problem using Net-OAuth2 if you have a version >= 0.50
# can be fixed by downgrading to 0.07 (not as severe as it sounds!) with:
#   cpan KGRENNAN/Net-OAuth2-0.07.tar.gz
#

require 5.10.0; # for the // operator (can be replaced with defined ? :)

use strict;
use warnings;
use Term::ReadKey;

use Getopt::Long;
use Net::Google::Spreadsheets;

use Text::CSV_XS;

my $outdir;
my $username;
my $title;
my $suffix = ".txt";
my $delimiter = "\t";

GetOptions( 
      "outdir=s"=>\$outdir,
      "username|googleaccount=s"=>\$username,
	    "title=s"=>\$title,
	    "suffix=s"=>\$suffix,
	    "delimiter=s"=>\$delimiter,
	  );

die "must provide spreadsheet's title (-title)\n" unless ($title);
die "must provide google username/email (-username)\n" unless ($username);
die "must provide output directory (-outdir)\n" unless ($outdir);
die "outdir exists but isn't a directory\n" if (-e $outdir && !-d $outdir);

mkdir $outdir unless (-d $outdir);

#
# get the google password
#
print "Enter your google account ('$username') password: ";
ReadMode('noecho'); # don't echo
chomp(my $password = <STDIN>);
ReadMode(0);        # back to normal
print "\n";

# log in
my $service = Net::Google::Spreadsheets->new
  (
   username => $username,
   password => $password,
  );
# a Net::Google module fails if can't log in
# so no pressing need to test for sucess here.


# find the spreadsheet by title
my $spreadsheet = $service->spreadsheet({ title => $title });
die "Cannot find spreadsheet doc '$title' for user '$username'\n" unless (defined $spreadsheet);
warn "Found spreadsheet '$title'\n";

# make a TSV formatter object thingy
my $tsv = Text::CSV_XS->new({
            binary => 1,
			      eol => $/,
			      sep_char => $delimiter,
			    });

# now loop over worksheets and write the files
foreach my $worksheet ($spreadsheet->worksheets) {
  my $w_title = $worksheet->title;
  my $outfile = "$outdir/$w_title$suffix";
  open(my $fh, ">:utf8", $outfile) || die "Help! can't open output file '$outfile'\n";

  # fill a 2D content array (sparsely) with values from the spreadsheet
  my @content; # [row][col] = cell->content
  map { $content[$_->row - 1][$_->col - 1] = $_->content } $worksheet->cells;

  # print each row tab delimited
  foreach my $row (@content) {
    $tsv->print($fh, $row // []);
  }

  close($fh);
  warn "wrote '$outfile'\n";
}
