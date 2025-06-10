use Text::Iconv;
my $converter = Text::Iconv->new("utf-8", "windows-1251");
# Text::Iconv is not really required. This can be any object with the convert method. Or nothing.
use Spreadsheet::XLSX;
my $excel = Spreadsheet::XLSX->new('IES2024.xlsx', $converter);
my $sheet = (@{$excel->{Worksheet}})[0]; 

my %budgetline = {};
my %volunteer  = {};

foreach my $row (13 .. $sheet->{MaxRow}) {
    my $flname = $sheet->{Cells}[$row][3]->{Val}  . " "   . $sheet->{Cells}[$row][4]->{Val};
    my $purpos = $sheet->{Cells}[$row][15]->{Val} . " - " . $sheet->{Cells}[$row][16]->{Val};
    my $amount = $sheet->{Cells}[$row][34]->{Val};
    my $uschck = $sheet->{Cells}[$row][35]->{Val};

    defined( $budgetline{$purpos} ) or $budgetline{$purpos} = 0;
    $budgetline{$purpos} = $budgetline{$purpos} + $amount;

    defined( $volunteer{$flname} ) or $volunteer{$flname} = 0;
    $volunteer{$flname} = $volunteer{$flname} + $amount;
}

foreach my $p ( sort ( keys(%budgetline) ) ) {
    printf("%-50s %10.3fK\n",  $p, $budgetline{$p}/1000.0 );
}

print "\n\n\n";

foreach my $p ( sort ( keys(%volunteer) ) ) {
    printf("%-50s %10.3fK\n",  $p, $volunteer{$p}/1000.0 );
}

print "\n\n\n";

foreach my $p ( sort { $volunteer{$a}<=>$volunteer{$b} } keys(%volunteer) ) {
    printf("%-50s %10.3fK\n",  $p, $volunteer{$p}/1000.0 );
}
