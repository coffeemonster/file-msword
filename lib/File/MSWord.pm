package File::MSWord;

use strict;

our $VERSION = '0.2';

=head1 NAME

File::MSWord

=head1 SYNOPSIS

    my $msword = File::MSWord->new( '/path/to/word.doc' );

    my %streams = $msword->listStreams;
    my %trash   = $msword->readTrash;
    my %summary = $msword->getSummaryInfo;
    my %docinfo = $msword->getDocumentSummaryInfo;

    See METHODS section below for more.

    See t/00-file-msword.t for more usage examples.

=head1 AUTHOR

Harlan Carvey, C<< <keydet89@yahoo.com> >>

=cut

use OLE::Storage;
use OLE::PropertySet;
use Startup;
use Carp;

my $self;

#-------------------------------------------------------
# new()
#-------------------------------------------------------
sub new {
    my ($class, $filename) = @_;

    my $self = { filename => $filename };

    # $self->{filesize} = (stat($self->{file}))[7];
    if (open($self->{hFile},"<",$filename)) {
        binmode($self->{hFile});
        return bless($self, $class);
    }
    else {
        carp "Could not open $filename : $!\n";
    }
}

#-------------------------------------------------------
# getGUID()
# Input : Nothing (uses $self)
# Output: Returns lower and upper halves of GUID in little-endian
#         order
#-------------------------------------------------------
sub getGUID {
    $self = shift;
    my $record;
    seek($self->{hFile},0,0);
    read($self->{hFile},$record,8);
# Returns lower and upper portions of GUID in little-endian order
# Looks like e011cfd0 and e11ab1a1, respectively
    return unpack("VV",$record);
}

#----------------------------------------------------
# getDocBinaryData()
# parse binary data from the file
# Input  : Nothing (uses $self)
# Output : Hash containing binary data from the document file
# Ref    : File::MSWord::Struct
#----------------------------------------------------
sub getDocBinaryData {
    $self = shift;
    my $record;
    my %doc =();

# Go to beginning of FIB
    seek($self->{hFile},0x200,0);
    read($self->{hFile},$record,20);
# wIdent   = magic number
# nFib     = FIB version
# nProduct = product version written by
# langid   = lang ID stamp
# pnNext
# fDot
# nFibBack = backwards compatibility setting
# lkey     = encryption key, if file is encrypted
# envr     = creation environment; 0 = Win, 1 = Mac
    ($doc{wIdent},$doc{nFib},$doc{nProduct},$doc{langid},$doc{pnNext},
        $doc{fDot},$doc{nFibBack},$doc{lKey},$doc{envr},$doc{fMac}) = unpack("v7VC2",$record);

    ($doc{fDot} & 0x0200) ? ($doc{table} = "1Table") : ($doc{table} = "0Table");
# Save for later use
    $self->{table} = $doc{table};
# The fDot and fMac values can be parsed using flags
# fDot & 0x0001 = doc is a template
# fDot & 0x0002 = doc is a glossary
# fDot & 0x0004 = doc is in complex, fast-saved format
# fDot & 0x0008 = file contains one or more pictures
# fDot & 0x0100 = file is encrypted
# fDot & 0x0200 = which table stream is valid
# fDot & 0x0400 = user has recommended that file be read-only
# fDot & 0x0800 = file is write reserved
# fDot & 0x8000 = file is encrypted
#
# fMac & 0x01 = file last saved on a Mac
# fMac & 0x10 = file last saved on Word97

  return %doc;
}

#-------------------------------------------------------
# getMagicIDs()
# Input : None
# Output: Returns values for creator/revisor IDs
# Magic ID for Word is 0x6a62
# 0x6a62 => MS Word 97
# 0x626a => Word 98 Mac
# 0xa5dc => Word 6.0/7.0
# 0xa5ec => Word 8.0
#-------------------------------------------------------
sub getMagicIDs {
    $self = shift;
    my $record;
    seek($self->{hFile},0x200 + 0x22,0);
    read($self->{hFile},$record,4);
    return unpack("vv",$record);
}

#-------------------------------------------------------
# getBuildDates()
# Input : None
# Output: Returns values for build date of creator and modifier
#         products, respectively
#-------------------------------------------------------
sub getBuildDates {
    $self = shift;
    return $self->getTwoDWORDs(0x44);
}

#-------------------------------------------------------
# getSavedBy()
# Input : None
# Output:    Returns offset and size of structure in table stream
#         recording names of users who have saved this document
#-------------------------------------------------------
sub getSavedBy {
    $self = shift;
    return $self->getTwoDWORDs(0x02d2);
}

#-------------------------------------------------------
# getDocUndo()
# Input : None
# Output:    Returns offset and size of structure in table stream
#         for "undocumented undo/versioning data"
#-------------------------------------------------------
sub getDocUndo {
    $self = shift;
    return $self->getTwoDWORDs(0x0302);
}

#-------------------------------------------------------
# getUndocOCX()
# Input : None
# Output:    Returns offset and size of structure in table stream
#         of "undocumented OCX data"
#-------------------------------------------------------
sub getUndocOCX {
    $self = shift;
    return $self->getTwoDWORDs(0x0342);
}

#-------------------------------------------------------
# getLastModified()
# Input : None
# Output: Two DWORDs containing the FILETIME object for the
#         last modified date in little-endian order; can be
#         translated using Math::BigInt and gmtime()
#-------------------------------------------------------
sub getLastModified {
    $self = shift;
    return $self->getTwoDWORDs(0x0352);
}

#-------------------------------------------------------
# getRoutingSlip()
# Input : None
# Output: Two DWORDs containing the offset to the mailer
#         routing slip in the table stream, and it's size
#-------------------------------------------------------
sub getRoutingSlip {
    $self = shift;
    return $self->getTwoDWORDs(0x02ca);
}

#-------------------------------------------------------
# getTwoDWORDs()
# Input : Offset within the file where two DWORDs start
# Output: Two DWORD values, in little endian order
#-------------------------------------------------------
sub getTwoDWORDs {
    $self = shift;
    my $ofs = shift;
    my $record;
    seek($self->{hFile},0x200 + $ofs,0);
    read($self->{hFile},$record,8);
    return unpack("VV",$record);
}

#----------------------------------------------------
# parseSTTBF()
# parse the STTBF
# Input : Buffer containing STTBF
# Output: Hash of hashes
#----------------------------------------------------
sub parseSTTBF {
    $self = shift;
    my $buff = shift;
    my $part1 = shift || "part1";
    my $part2 = shift || "part2";
    my %revLog = ();
    my $num_str = unpack("v",substr($buff,2,2));
    my $cursor = 6;
    my ($size,$str);
    foreach my $i (1..($num_str/2)) {
        $size = unpack("v",substr($buff,$cursor,2));
        $cursor += 2;
        $str  = substr($buff,$cursor,$size*2);
        $str =~ s/\00//g;
        $revLog{$i}{$part1} = $str;
        $cursor += $size*2;

        $size = unpack("v",substr($buff,$cursor,2));
        $cursor += 2;
        $str  = substr($buff,$cursor,$size*2);
        $str =~ s/\00//g;
        $revLog{$i}{$part2} = $str;
        $cursor += $size*2;
    }
    return %revLog;
}


#-------------------------------------------------------
# listStreams()
# Input : None
# Output: Hash of hashes with the names of each stream as
#         the key.  Second-level keys are temp names (leading
#         non-ASCII char removed), size, and the actual stream
#         (buffer), if size > 0.
#-------------------------------------------------------
sub listStreams {
    $self = shift;
    my ($size,$tempname);
    my %streams = ();

    my $var = OLE::Storage->NewVar();
    my $startup = new Startup;
    my $doc = OLE::Storage->open($startup,$var,$self->{filename});
    my @pps = $doc->dirhandles(0);
    foreach my $pps (sort {$a <=> $b} @pps) {
        my $name = $doc->name($pps)->string();
        $tempname = $name;
        $tempname = (split(/^\W/,$name))[1] if ($name =~ m/^\W/);
        my $buff;
        $doc->read($pps,\$buff);
        my $size = length($buff);
        $streams{$name}{tempname} = $tempname;
        $streams{$name}{size}     = $size;
        $streams{$name}{buffer}   = $buff if ($size > 0);

        if ($doc->is_file($pps)) {
            $streams{$name}{type}     = 'File';
        }
        elsif ($doc->is_directory($pps)) {
            $streams{$name}{type}     = 'Dir';
        }
        else {

        }
    }
    return %streams;
}

#-------------------------------------------------------
# getSummaryInfo()
# Input : None
# Output: Hash containing contents of SummaryInformation stream
#-------------------------------------------------------
sub getSummaryInfo {
    $self = shift;
    my ($size,$tempname);
    my %suminfo = ();

    my $var = OLE::Storage->NewVar();
    my $startup = new Startup;
    my $doc = OLE::Storage->open($startup,$var,$self->{filename});
    my @pps = $doc->dirhandles(0);
    foreach my $pps (sort {$a <=> $b} @pps) {
        my $name = $doc->name($pps)->string();
        next unless ($name eq "\05SummaryInformation");
        if (my $prop = OLE::PropertySet->load($startup,$var,$pps,$doc)) {
            ($suminfo{title},$suminfo{subject},$suminfo{authress},$suminfo{lastauth},
            $suminfo{revnum},$suminfo{appname},$suminfo{created},$suminfo{lastsaved},
            $suminfo{lastprinted}) = string {$prop->property(2,3,4,8,9,18,12,13,11)};
        }
    }
    return %suminfo;
}

#-------------------------------------------------------
# getDocSummaryInfo()
# Input : None
# Output: Hash containing contents of DocumentSummaryInformation stream
#-------------------------------------------------------
sub getDocSummaryInfo {
    $self = shift;
    my ($size,$tempname);
    my %docsuminfo = ();

    my $var = OLE::Storage->NewVar();
    my $startup = new Startup;
    my $doc = OLE::Storage->open($startup,$var,$self->{filename});
    my @pps = $doc->dirhandles(0);
    foreach my $pps (sort {$a <=> $b} @pps) {
        my $name = $doc->name($pps)->string();
        next unless ($name eq "\05DocumentSummaryInformation");
        if (my $prop = OLE::PropertySet->load($startup,$var,$pps,$doc)) {
                    $docsuminfo{org} = string {$prop->property(15)};
        }
    }
    return %docsuminfo;
}

#-------------------------------------------------------
# readStreamTable()
# Input : offset within stream table, and number of bytes to read
# Output: Buffer containing portion of the stream table
#-------------------------------------------------------
sub readStreamTable {
    $self    = shift;
    my $ofs  = shift;
    my $size = shift;
    my ($record,$table,$buff,$fDot);
    my $buff;
# Get table to read
    seek($self->{hFile},0x200 + 0xA,0);
    read($self->{hFile},$record,2);
    $fDot = unpack("v",$record);
    ($fDot & 0x0200) ? ($table = "1Table") : ($table = "0Table");

    my $var = OLE::Storage->NewVar();
    my $startup = new Startup;
    my $doc = OLE::Storage->open($startup,$var,$self->{filename});
    my @pps = $doc->dirhandles(0);
    foreach my $pps (sort {$a <=> $b} @pps) {
        my $name = $doc->name($pps)->string();
        print "$name\n";
        next unless ($name eq $table);
        $doc->read($pps,\$buff,$ofs,$size);
        return $buff;
    }
}
#-------------------------------------------------------
# readTrash()
# Input : None
# Output: Hash of hashes, with trash bin names as keys; the
#         sizes of each bin and the actual content (buffer) are the
#         second level keys.
#-------------------------------------------------------
sub readTrash {
    $self = shift;
    my %trash_streams = ();
    my %trash = (1 => "BigBlocks",
                   2 => "SmallBlocks",
                   4 => "FileEndSpace",
                   8 => "SystemSpace");

    my $var = OLE::Storage->NewVar();
    my $startup = new Startup;
    my $doc = OLE::Storage->open($startup,$var,$self->{filename});

    foreach my $type (sort {$a <=> $b} keys %trash) {
        my $buff;
        $doc->read_trash($type,\$buff);
        my $size = length($buff);
        $trash_streams{$trash{$type}}{size} = $size;
        $trash_streams{$trash{$type}}{buffer} = $buff if ($size > 0);
    }
    return %trash_streams;
}

#----------------------------------------------------------
# getLangID()
# Input : Language ID (hex)
# Output: Language ID (readable)
#----------------------------------------------------------
sub getLangID {
    $self = shift;
    my $id = shift;
    my %langID = (
        0x0400 => "None",
        0x0401 => "Arabic",
        0x0402 => "Bulgarian",
        0x0403 => "Catalan",
        0x0404 => "Traditional Chinese",
        0x0804 => "Simplified Chinese",
        0x0405 => "Czech",
        0x0406 => "Danish",
        0x0407 => "German",
        0x0807 => "Swiss German",
        0x0408 => "Greek",
        0x0409 => "English (US)",
        0x0809 => "British English",
        0x0c09 => "Australian English",
        0x040a => "Castilian Spanish",
        0x080a => "Mexican Spanish",
        0x040b => "Finnish",
        0x040c => "French",
        0x080c => "Belgian French",
        0x0c0c => "Canadian French",
        0x100c => "Swiss French",
        0x040d => "Hebrew",
        0x040e => "Hungarian",
        0x040f => "Icelandic",
        0x0410 => "Italian",
        0x0810 => "Swiss Italian",
        0x0411 => "Japanese",
        0x0412 => "Korean",
        0x0413 => "Dutch",
        0x0813 => "Belgian Dutch",
        0x0414 => "Norwegian (Bokmal)",
        0x0814 => "Norwegian (Nynorsk)",
        0x0415 => "Polish",
        0x0416 => "Brazilian Portuguese",
        0x0816 => "Portuguese",
        0x0417 => "Rhaeto-Romanic",
        0x0418 => "Romanian",
        0x0419 => "Russian",
        0x041a => "Croato-Serbian (Latin)",
        0x081a => "Serbo-Croatian (Cyrillic)",
        0x041b => "Slovak",
        0x041c => "Albanian",
        0x041d => "Swedish",
        0x041e => "Thai",
        0x041f => "Turkish",
        0x0420 => "Urdu",
        0x0421 => "Bahasa",
        0x0422 => "Ukrainian",
        0x0423 => "Byelorussian",
        0x0424 => "Slovenian",
        0x0425 => "Estonian",
        0x0426 => "Latvian",
        0x0427 => "Lithuanian",
        0x0429 => "Farsi",
        0x042D => "Basque",
        0x042F => "Macedonian",
        0x0436 => "Afrikaans",
        0x043E => "Malaysian "
    );

    (exists $langID{$id}) ? (return $langID{$id}) : (return "Unknown");
}

#-------------------------------------------------------
# close()
#-------------------------------------------------------
sub close {close($self->{hFile});}


=head1 DESCRIPTION

    Perl module to parse MSWord OLE compound documents without relying on the MS
    API.  Neither MSOffice nor MSWord need to be installed to use this module.  The
    intent of this module is to provide a cross-platform method for retrieving
    metadata from MSWord documents.  This module parses binary information in the
    file headers, and lists/dumps the various streams and 'trash'.

    All methods return binary values in little-endian order, unless otherwise
    specified.


=head1 METHODS

=head2 my $word = File::MSWord::new()

    Creates a new $word object.

=head2 @guid = $word->getGUID();

    Returns a 2-element list containing the halves of the GUID, in little-endian order.

=head2 %doc = $word->getDocBinaryData();

    Returns a hash containing various elements of binary header information located in the file.

=head2 @ids = $word->getMagicIDs()

    Returns IDS for creator/reviser apps (ie, Word version)

=head2 @dates = $word->getBuildDates()

    Returns 2 DWORDS holding the build dates of the creator/reviser apps

=head2 @list = $word->getSavedBy()

    Returns 2 DWORDS corresponding to the offset within the table stream of the list of names (of
    users who have saved this document, alternating with the path the file was saved to), and the
    size of the buffer.

=head2 @list = $word->getDocUndo()

    Returns 2 DWORDS (offset, size) of undocumented undo information saved in the table stream.
    This is one of several "undocumented" areas of undo/versioning information listed in the
    primary reference.

=head2 @list = $word->getUndocOCX()

    Returns 2 DWORDS (offset,size) of undocumented OCX data within the table stream.

=head2 @list = $word->getLastModified()

    Returns 2 DWORDS corresponding to the last modified FILETIME object.  This information
    can be fed to a routine using Math::BigInt and gmtime() to return something readable.

=head2 @list = $word->getRoutingSlip()

    Returns 2 DWORDS (offset,size) corresponding to the routing slip information maintained in
    the table stream.

=head2 @list = $word->getTwoDWORDs($offset)

    Takes an offset within the file information block (FIB) and returns 2 DWORDS located in
    8 bytes starting at that offset.

=head2 %hash = $word->listStreams()

    Returns a hash of hashes containing the names of the streams of the OLE/compound/structured storage
    document as the keys.

=head2 %hash = $word->getSummaryInfo()

    Returns a hash containing elements of the SummaryInformation stream

=head2 %hash = $word->getDocSummaryInfo()

    Returns a hash containing elements of the DocumentSummaryInformation stream

=head2 $buffer = $word->readStreamTable($offset,$size)

    Takes in an offset and size of a buffer within the table stream, and returns the
    contents of the buffer.

=head2 %hash = $word->parseSTTBF($buffer[,$name1,$name2])

    Takes a buffer (extracted from the table stream) and parses it out into a hash of hashes,
    whose keys are the order (1,2,3...) of the entries.  Optionally, you can pass in the names
    of the subkeys.

=head2 $landid = $word->getLangID($id)

    Translates the language id from the FIB into something readable.

=head2 %hash = $word->readTrash()

    Reads the trash bins in an OLE/compound/structured storage document.  Returns a hash of
    hashes with the names of the trash bins as keys, and the size and contents of the bins
    as subkeys.

=cut


=head1 REFERENCES

    The primary reference for this module is wv's convert-to-struct/demo.txt

    See: perldoc File::MSWord::Struct

    Metadata in MSWord documents has been an issue for quite a while. See:

    L<http://www.computerbytesman.com/privacy/blair.htm>
    L<http://blogs.washingtonpost.com/securityfix/2005/12/document_securi.html>
    L<http://www.forbes.com/2005/12/13/microsoft-word-merck_cx_de_1214word.html>

    MS KB 290945: How to minimize metadata in Word 2002
    L<http://support.microsoft.com/kb/290945>

=head1 BUGS

    Please report any bugs and feature requests to C<< <keydet89@yahoo.com> >>.

=head1 COPYRIGHT AND LICENSE

    Copyright (C) 2011 by Harlan Carvey

    This library is free software; you can redistribute it and/or modify
    it as you like.  However, please be sure to provide proper credit where
    it is due.

=cut

1; # End of File::MSWord

