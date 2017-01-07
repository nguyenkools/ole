
package lstmp;

use strict;
use warnings;
use Win32::OLE qw(in);
use Win32::OLE::Const 'Microsoft Office 15.0 Object Library';
use Data::Dumper;
use v5.10;
use Exporter qw(import);
 
our @EXPORT_OK = qw(propertieslist methodslist dumpobject);


sub propertieslist{

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


sub methodslist{

	my $OleObject =shift;
	my $typeinfo = $OleObject->GetTypeInfo();
	my $attr = $typeinfo->_GetTypeAttr();
	for (my $i = 0; $i< $attr->{cFuncs}; $i++) {
		my $desc = $typeinfo->_GetFuncDesc($i);
		# the call conversion of method was detailed in %$desc
		my $funcname = @{$typeinfo->_GetNames($desc->{memid}, 1)}[0];
		say $funcname;
	}
}
sub dumpobject{

	my $Obj = shift;
	my $depth = shift;
	local $Data::Dumper::Maxdepth = $depth;#2;
	print Data::Dumper->Dump( [ $Obj ] )

}


1;