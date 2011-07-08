#!perl -T

use strict;

use Test::More tests => 10;

# Load Module
# -------------------------------------------
BEGIN {
    use_ok('File::MSWord') || print "Bail out!\n";
}

diag("Testing File::MSWord $File::MSWord::VERSION, Perl $], $^X");

my $msword;

# Instantiate Module
# -------------------------------------------
{
    use FindBin qw/$Bin/;
    my $filename = "$Bin/word.doc";
    ok( $msword = File::MSWord->new($filename) );
    is( ref($msword), 'File::MSWord' );
}

# Streams, Trash and GUID
# -------------------------------------------
{
    my %streams = $msword->listStreams();
    diag("Stream: $_\n") foreach keys %streams;
    ok(%streams);

    my %trash = $msword->readTrash();
    diag( sprintf( "Trash: %-15s %-8s\n", "Trash Bin", "Size" ) );
    diag( sprintf( "Trash: %-15s %-8d\n", $_, $trash{$_}{size} ) ) foreach keys %trash;
    ok(%trash);

    my ( $id_l, $id_u ) = $msword->getGUID();
    diag( sprintf( "GUID: %x - %x\n", $id_l, $id_u ) );
    ok($id_l);
    ok($id_u);
}

# SummaryInfo
# -------------------------------------------
{
    my %summary = $msword->getSummaryInfo();
    diag("Summary: $_ => $summary{$_}\n") for keys %summary;
    ok(%summary);

    my %docinfo = $msword->getDocSummaryInfo();
    diag("DocInfo: $_ => $docinfo{$_}\n") for keys %docinfo;

    # ok(%docinfo); -- can be blank
}

# Last 10 Authors
# -------------------------------------------
{
    my ( $ofs, $size ) = $msword->getSavedBy();
    ok($size);
    diag sprintf( "Authors: 0x%x -> 0x%x\n", $ofs, $size );
    my $buff = $msword->readStreamTable( $ofs, $size );
    my %revlog = $msword->parseSTTBF( $buff, "author", "path" );
    foreach my $k ( sort { $a <=> $b } keys %revlog ) {
        diag(
            sprintf(
                "Authors: %-4s %-15s %-60s\n",
                $k, $revlog{$k}{author},
                $revlog{$k}{path}
            )
        );
    }
}

# BinaryData, Language
# -------------------------------------------
{
    my %binary = $msword->getDocBinaryData();
    ok(%binary);
    diag( sprintf( "Binary: $_ => 0x%x\n", $binary{$_} ) ) for keys %binary;
    diag( "LanguageID: " . $msword->getLangID( $binary{langid} ) . "\n" );
}

