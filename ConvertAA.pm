
package Spreadsheet::ConvertAA;

use 5.006;
use strict;
use warnings;
use Carp ;

require Exporter;
use AutoLoader qw(AUTOLOAD);

our @ISA = qw(Exporter);

our %EXPORT_TAGS = ( 'all' => [ qw() ] );

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } );

our @EXPORT = qw( FromAA ToAA);
our $VERSION = '0.03';

#----------------------------------------------------

my @AA_divider = (1,   26,   676,  17576,   456976) ;

#----------------------------------------------------

sub FromAA
{
my $baseAA = shift ;
my ($AA_bit, $base10) = (0, 0) ;

confess "Invalid baseAA '$baseAA'" if($baseAA =~ /[^A-Za-z@]/) ;
confess "BaseAA '$baseAA' is too big" if(length($baseAA) > 4) ;

for (split //, (reverse $baseAA))
	{
	$base10 += ((ord(uc($_)) - ord('A')) + 1) * $AA_divider[$AA_bit] ;
	$AA_bit++ ;
	}

return($base10) ;
}

#----------------------------------------------------

sub ToAA
{
my $base10 = shift ;
confess "Base10 '$base10' is too big" if($base10 > 475254) ;
confess "Base10 '$base10' is negative" if($base10 < 0) ;

return('@') if $base10 == 0 ;

my $baseAA = '' ;

for (reverse 0 .. 4)
	{
	my $result = $base10 / $AA_divider[$_] ;
	my $rest   = $base10 % $AA_divider[$_] ;
	
	if($result > 26)
		{
		$baseAA .= 'Z' ;
		$base10 = $AA_divider[$_] + $rest ;
		}
	else
		{
		$baseAA .= chr( (ord('A') - 1) + $result) ;
		$base10 = $rest ;
		}
	}

while($baseAA =~ /@/g)
	{
	#~ print "=>$baseAA\n" ;
	next if($baseAA =~ s/^@//) ;
	
	my $match_position   = pos($baseAA) - 1 ;
	my $char_before_zero = substr($baseAA, $match_position - 1 , 1) ;
	
	substr($baseAA, $match_position, 1, 'Z') ;
	substr($baseAA, $match_position - 1 , 1, chr(ord($char_before_zero) - 1)) ;
	#~ print "===>$baseAA\n" ;
	}

return($baseAA) ;
}

1;
__END__
=head1 NAME

Spreadsheet::ConvertAA - Perl extension for Converting Spreadsheet column name to/from  decimal

=head1 SYNOPSIS

  use Spreadsheet::ConvertAA ;

  my $baseAA = ToAA(475255) ;
  my $base10 = FromAA('AAAZ') ;

=head1 DESCRIPTION

This module allows you to convert from Spreadsheet column notation ( 'A', 'AZ', 'BC', ..) to decimal and back.
The Spreadsheet column notation is base 26 _without_ zero. 'A' is 1 and 'AA' is 27. I named the base 'AA' because
I found no better name.

ToAA taked a base 10 numbers in the range [0 .. 475254]. 0 (zero) is special and converted to '@'.
FromAA takes a string in the range [A-Za-z]{1,4}.

Spreadsheet::ConvertAA 'confess' on invalid input.

Convertion is done in the range: 'A' .. 'ZZZZ' which corresponds to 1 .. 475254.

=head2 EXPORT

ToAA and FromAA

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2004 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=head1 SEE ALSO

L<Spreadsheet::Perl>.

=cut

