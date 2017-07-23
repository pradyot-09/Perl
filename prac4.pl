use Spreadsheet::WriteExcel;
use Spreadsheet::Read;
 
my $workbook  = Spreadsheet::WriteExcel->new( 'Output.xls' );
my $worksheet = $workbook->add_worksheet();

$format1=$workbook->add_format(); 
$format1->set_size('11');
$format1->set_font('Calibri');

$format2 = $workbook->add_format();
$format2->set_size('7.5');
$format2->set_align("right");
$format2->set_font('Verdana');
$format2->set_bg_color('green');
$format2->set_border(1);

$format3 = $workbook->add_format();
$format3->set_size('7.5');
$format3->set_font('Verdana');
$format3->set_border(1);

$format4 = $workbook->add_format();
$format4->set_size('7.5');
$format4->set_font('Verdana');
$format4->set_bg_color("orange");
$format4->set_border(1);

my $data = ReadData ('Input.xlsx');

my @row1 = Spreadsheet::Read::row($data->[1], 1);
my @row2 = Spreadsheet::Read::row($data->[1], 2);


$worksheet->write("A1","State",$format1);
$worksheet->write("B1","Time",$format1);
$worksheet->write("C1","Value",$format1); 
$s=1 ;
$e=6 ;
$k=3;
for($j=1;$j<31;$j=$j+1)
{
	$q=1;
	for($i=$s;$i<$e;$i=$i+1)
	{
		$worksheet->write($i,0,$row2[0],$format3);
		$worksheet->write($i,1,$row1[$q],$format2);
		if($i%2==0)
		{
			$worksheet->write($i,2,$row2[$q],$format3);
		}
		else
		{
			$worksheet->write($i,2,$row2[$q],$format4);
		}
		$q=$q+1 ;
	}
	$s=$s+5;
	$e=$e+5 ;
	 @row2 = Spreadsheet::Read::row($data->[1], $k);
	$k=$k+1
}
$workbook->close;