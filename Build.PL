use 5.006;
use strict;
use warnings;
use Module::Build;

my $builder = Module::Build->new(
    module_name         => 'File::MSWord',
    license             => 'perl',
    dist_author         => q{Harlan Carvey, <keydet89@yahoo.com>},
    dist_abstract       => 'Grab metadata from MSWord documents',
    dist_version_from   => 'lib/File/MSWord.pm',
    build_requires => {
        'Test::More'    => 0,
    },
    requires => {
        'perl'             => 5.006,
        'Carp'             => 0,
        'OLE::Storage'     => 0,
        'OLE::PropertySet' => 0,
        'Startup'          => 0,
        'Unicode::Map'     => 0,
      },
    add_to_cleanup      => [ 'File-MSWord-*' ],
);

$builder->create_build_script();
