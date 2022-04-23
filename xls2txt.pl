#!/usr/bin/perl

# MIT License
#
# Copyright (c) 2022 Giovanni Bechis
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# xls2csv.pl - parses an excel spreadsheet and writes its text content to stdout
# receives the Excel filename as first arguement

use strict;
use warnings;

use Getopt::Std;
use Spreadsheet::ParseExcel;

my %opts;
my $minstrsize = 20;
my $filename;
my $onlyuri = 0;

getopts('i:s:u', \%opts);
if ( defined $opts{'i'} ) {
  $filename = $opts{'i'};
} else {
  print "Usage: ", $0 , " -i Excel_File [ -s \$min_cell_size -u ]\n";
  exit;
}
if ( defined $opts{'s'} ) {
  $minstrsize = $opts{'s'};
}
if ( defined $opts{'u'} ) {
  $onlyuri = 1;
}

my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse($filename);

if ( !defined $workbook ) {
  warn $parser->error(), ".\n";
  exit;
}

for my $worksheet ( $workbook->worksheets() ) {

  my ( $row_min, $row_max ) = $worksheet->row_range();
  my ( $col_min, $col_max ) = $worksheet->col_range();

  for my $row ( $row_min .. $row_max ) {
    for my $col ( $col_min .. $col_max ) {
      my $cell = $worksheet->get_cell( $row, $col );
      next unless $cell;
      next unless length($cell->value) > $minstrsize;
      if($onlyuri eq 1) {
        if($cell->value() !~ /https?:\/\/.{3,256}|www\.|\.[a-z0-9_-]{3,64}\.[a-z]{2,6}/) {
          next;
        }
      }
      print $cell->value() . "\n";
    }
  }
}
