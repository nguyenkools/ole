use Win32::OLE qw(in);
use Win32::OLE::Const 'Microsoft Office 15.0 Object Library';
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
# methodlist($doc->Paragraphs);

foreach my $sec (in $doc->Sections){
	foreach $hder (in $sec->Headers){
		$hder->Range->Delete;
		# printproperty($hder);
		# methodlist($hder);
	}
	
	foreach $ftr (in $sec->Footers){
		$ftr->Range->Delete;
		# printproperty($ftr);
		# methodlist($ftr);
	}
}

say "count";
say $doc->TablesOfContents->Count;
my $c;
foreach my $tbct (in $doc->TablesOfContents){

			$tbct->Delete;
			
			say "aaaa".$c;
			# methodlist($tbct);

}

printproperty ($doc->ActiveWindow->Panes(1));
say ($doc->ActiveWindow->Panes(1)->Pages(1));
# printproperty($doc);
# methodlist($doc);
$doc->Save;
$doc->Close;

sub printproperty{
    my $Object = shift;
	print "OLE object's properties:\n";
	foreach my $Key (sort keys %$Object) {
		my $Value;

		eval {$Value = $Object->{$Key} };
		$Value = "***Exception***" if $@;
		say $Value;
		$Value = "<undef>" unless defined $Value;
		
		$Value = '['.Win32::OLE->QueryObjectType($Value).']' 
		  if UNIVERSAL::isa($Value,'Win32::OLE');

		$Value = '('.join(',',@$Value).')' if ref $Value eq 'ARRAY';

		printf "%s %s %s\n", $Key, '.' x (40-length($Key)), $Value;
}
}


sub methodlist{
my $OleObject =shift;
my $typeinfo = $OleObject->GetTypeInfo();
my $attr = $typeinfo->_GetTypeAttr();
for (my $i = 0; $i< $attr->{cFuncs}; $i++) {
    my $desc = $typeinfo->_GetFuncDesc($i);
    # the call conversion of method was detailed in %$desc
    my $funcname = @{$typeinfo->_GetNames($desc->{memid}, 1)}[0];
    say $funcname;
}

sub dumpMedandPros{

	my $Obj = shift;
	my $depth = shift;
	local $Data::Dumper::Maxdepth = $depth;#2;
	print Data::Dumper->Dump( [ $my_ole_object ] )

}


}


