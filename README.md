google-spreadsheet-to-textfiles
===============================

Uses the API to fetch and dump tab-delimited text files (configurable delimiter) for each worksheet on your local filesystem.

#
# USAGE: ./fetch-tsv-from-google-spreadsheet.pl -outdir my-dir -username myemail@gmail.com -title "Spreadsheet title"
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
