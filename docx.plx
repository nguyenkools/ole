use lib ".";
use Win32::OLE qw(in);
use Win32::OLE::Const 'Microsoft Office 15.0 Object Library';
use lstmp qw(propertieslist methodslist dumpobject);
use v5.10;


my $word = Win32::OLE->new( 'Word.Application', 'Quit' );
my $doc  = $word->Documents->Open( 'C:\Users\Nhat\Desktop\docx\t.docx' ) || die 'Unable to open document: ', Win32::OLE->LastError;

my $paragraphs = Win32::OLE::Enum->new( $doc->Paragraphs );

while ( defined( my $paragraph = $paragraphs->Next ) ) {
    my $words = Win32::OLE::Enum->new( $paragraph->{Range}->{Words} );

    while ( defined( my $word = $words->Next ) ) {
        print $word->{Text};
    }
}
# methodslist($doc->Paragraphs);

foreach my $sec (in $doc->Sections){
	foreach $hder (in $sec->Headers){
		$hder->Range->Delete;
		# propertieslist($hder);
		# methodslist($hder);
	}
	
	foreach $ftr (in $sec->Footers){
		$ftr->Range->Delete;
		# propertieslist($ftr);
		# methodslist($ftr);
	}
}

say "count";
say $doc->TablesOfContents->Count;
my $c;
foreach my $tbct (in $doc->TablesOfContents){

			$tbct->Delete;
			
			say "aaaa".$c;
			# methodslist($tbct);

}

propertieslist ($doc->ActiveWindow->Panes(1));
say ($doc->ActiveWindow->Panes(1)->Pages(1));
# propertieslist($doc);
# methodslist($doc);
$doc->Save;
$doc->Close;




