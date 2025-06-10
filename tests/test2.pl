use Text::Iconv;
my $converter = Text::Iconv->new("utf-8", "windows-1251");
# Text::Iconv is not really required. This can be any object with the convert method. Or nothing.
use Spreadsheet::XLSX;
my $excel = Spreadsheet::XLSX->new('IES2024.xlsx', $converter);
my $sheet = (@{$excel->{Worksheet}})[0]; 
#$sheet->{MaxRow} ||= $sheet->{MinRow}; 
foreach my $row (12 .. $sheet->{MaxRow}) {
    #$sheet->{MaxCol} ||= $sheet->{MinCol};
    my $flname = $sheet->{Cells}[$row][3]->{Val}  . " "   . $sheet->{Cells}[$row][4]->{Val};
    my $purpos = $sheet->{Cells}[$row][15]->{Val} . " - " . $sheet->{Cells}[$row][16]->{Val};
    my $amount = $sheet->{Cells}[$row][34]->{Val};
    my $uschck = $sheet->{Cells}[$row][35]->{Val};

    printf("%s,%s,%s %s\n", $flname, $purpos, $amount, $uschck ); 
}
