use warnings;
use strict;
#use PDL::Stats;
use Statistics::Basic;
use Text::Iconv;
my $converter = Text::Iconv->new("utf-8", "windows-1251");
# Text::Iconv is not really required. This can be any object with the convert method. Or nothing.
use Spreadsheet::XLSX;
my $excel = Spreadsheet::XLSX->new('IES2024.xlsx', $converter);
my $sheet = (@{$excel->{Worksheet}})[0]; 

# totals and totals with itemizations
my %budgetlinetotal;
my %budgetlinepervol;
my %budgetlineperrep;
#
my %volunteertotal;
my %volunteerperpur;
my %volunteerperrep;

# Note: rows here start at 0 and not 1 as in the Office software
foreach my $row (13 .. $sheet->{MaxRow} ) {
    my $repid  = $sheet->{Cells}[$row][0]->{Val};
    my $flname = $sheet->{Cells}[$row][3]->{Val}  . " "   . $sheet->{Cells}[$row][4]->{Val};
    my $purpos = $sheet->{Cells}[$row][15]->{Val} . " - " . $sheet->{Cells}[$row][16]->{Val};
    my $amount = $sheet->{Cells}[$row][34]->{Val};
    my $uschck = $sheet->{Cells}[$row][35]->{Val};

    ($uschck eq "USD") or printf("WARNING: Excel line %d is not in USD\a\n", $row+1);

    #########
    defined( $budgetlinetotal{$purpos} ) or $budgetlinetotal{$purpos} = 0;
    $budgetlinetotal{$purpos} += $amount;

    defined( $budgetlinepervol{$purpos} ) or $budgetlinepervol{$purpos} = {};
    defined( $budgetlinepervol{$purpos}{$flname} ) or $budgetlinepervol{$purpos}{$flname} = 0;
    $budgetlinepervol{$purpos}{$flname} += $amount;

    defined( $budgetlineperrep{$purpos} ) or $budgetlineperrep{$purpos} = {};
    defined( $budgetlineperrep{$purpos}{$repid} ) or $budgetlineperrep{$purpos}{$repid} = 0;
    $budgetlineperrep{$purpos}{$repid} += $amount;

    ##########
    defined( $volunteertotal{$flname} ) or $volunteertotal{$flname} = 0;
    $volunteertotal{$flname} += $amount;

}

if (1) {
        foreach my $p ( sort ( keys(%budgetlinetotal) ) ) {
        printf("%-50s %10.3fK\n",  $p, $budgetlinetotal{$p}/1000.0 );
    }
    print "\n\n\n";
}

if (1) {
    foreach my $p ( sort ( keys(%budgetlinepervol) ) ) {
        printf("%-50s %10.3fK TOTAL\n",  $p, $budgetlinetotal{$p}/1000.0 );
        my @data = ();
        foreach my $n ( sort ( keys %{ %budgetlinepervol{$p} } ) ) {
            my $item = $budgetlinepervol{$p}{$n}/1000.0;
            # printf("     %-45s %10.3fK\n",  $n, $item );
            push( @data, $item );
        }
        @data = sort (@data);
        print "    PER PERSON: @data \n";
        printf("%-50s %10.3fK TOTAL\n",  $p, $budgetlinetotal{$p}/1000.0 );
        if ($#data>=9) {
            print "    90th % is " . $data[0.9*$#data] . "\n";
        }
        printf("    STATS: CNT:%3d AVG: %.3f MED: %.3f STD: %.3f\n\n", 
            $#data+1,
            Statistics::Basic::mean(@data), 
            Statistics::Basic::median(@data), 
            Statistics::Basic::stddev(@data) 
            );
    }
    print "\n\n\n";
}

if (1) {
    foreach my $p ( sort ( keys(%budgetlineperrep) ) ) {
        printf("%-50s %10.3fK TOTAL\n",  $p, $budgetlinetotal{$p}/1000.0 );
        my @data = ();
        foreach my $r ( sort ( keys %{ %budgetlineperrep{$p} } ) ) {
            my $item = $budgetlineperrep{$p}{$r}/1000.0;
            # printf("     %-45s %10.3fK\n",  $r, $item );
            push( @data, $item );
        }
        @data = sort (@data);
        print "    PER TRIP: @data \n";
        printf("%-50s %10.3fK TOTAL\n",  $p, $budgetlinetotal{$p}/1000.0 );
        if ($#data>=9) {
            print "    90th % is " . $data[0.9*$#data] . "\n";
        }
        printf("    STATS: CNT:%3d AVG: %.3f MED: %.3f STD: %.3f\n\n", 
            $#data+1,
            Statistics::Basic::mean(@data), 
            Statistics::Basic::median(@data), 
            Statistics::Basic::stddev(@data) 
            );
    }
    print "\n\n\n";
}


if (0) {
    foreach my $p ( sort ( keys(%volunteertotal) ) ) {
        printf("%-50s %10.3fK\n",  $p, $volunteertotal{$p}/1000.0 );
    }
    print "\n\n\n";
}

if (0) {
    foreach my $p ( sort { $volunteertotal{$a}<=>$volunteertotal{$b} } keys(%volunteertotal) ) {
        printf("%-50s %10.3fK\n",  $p, $volunteertotal{$p}/1000.0 );
    }
    print "\n\n\n";
}
