<?php

namespace App\Http\Controllers;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Http\Request;

class ExcelController extends Controller
{
    function __construct(){
        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->alphabet = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
    }
    private function update_alignment(){
        $this->sheet->getStyle('B')->getAlignment()->setHorizontal('left');
        $this->sheet->getStyle('E2:CJ5')->getAlignment()->setHorizontal('center');
        $this->sheet->getStyle('C6:D999')->getAlignment()->setHorizontal('center');
        $this->sheet->getStyle('A:CJ')->getAlignment()->setVertical('center');
    }
    private function get_column($i){
        if ($i <= 26){
            return $this->alphabet[$i-1];
        }else{
            $first = floor($i/26);
            $second = $i - 26 * $first ;
            if($second ==0){
                return $this->alphabet[$first-2 ].''.$this->alphabet[25];
            };
            return $this->alphabet[$first-1].''.$this->alphabet[$second-1];
        }
    }
    private function update_row_height(){
        $this->sheet->getRowDimension('2')->setRowHeight(22.9);
    }
    private function update_borders($columns){
        $outtline = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
                'inside' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE,
                ]
            ],
        ];
        $this->sheet->getStyle($columns)->applyFromArray($outtline);
    }
    private function set_borders(){
        $right = [
            'borders' => [
                'right' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $bottom = [
            'borders' => [
                'bottom' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        for ($j = 5; $j <= 12; $j++)
        for ($i = 5; $i <= 87; $i = $i +4 ){
            $start = $this->get_column($i);
            $end = $this->get_column($i + 3);
            $this->sheet->getStyle("$end$j")->applyFromArray($right);
            $this->sheet->getStyle("$start$j:$end$j")->applyFromArray($bottom);
        }
        $inside = [
            'borders' => [
                'inside' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $this->sheet->getStyle('C2:CJ4')->applyFromArray($inside);
        $this->sheet->getStyle('B6:D12')->applyFromArray($inside);
    }
    private function add_outline_borders(){
        $outtline = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
                ],
            ],
        ];
        $this->sheet->getStyle('C2:CJ4')->applyFromArray($outtline);
        $this->sheet->getStyle('B6:D12')->applyFromArray($outtline);
        $this->sheet->getStyle('E5:CJ12')->applyFromArray($outtline);
        $this->sheet->getStyle('B5:CJ5')->applyFromArray($outtline);
    }
    private function make_bg($columns, $color){
        $this->sheet->getStyle($columns)
        ->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()
        ->setARGB($color);
    }
    private function update_bg(){
        $this->make_bg('A1:D414','ffffff');
        $this->make_bg('E5:FO5','ffffff');
        $this->make_bg('E13:FO414','ffffff');
        $this->make_bg('CK1:FO12','ffffff');
        $this->make_bg('A1:CK1','ffffff');
        $this->make_bg('C2:CG2','E7E6E6');
        $this->make_bg('C3:CG3','E2EFDA');
        $this->make_bg('C4:CG4','FCE4D6');
        $this->make_bg('B5:CJ5','E7E6E6');
    }
    private function update_width(){
        $this->sheet->getColumnDimension('B')->setWidth(40.21);
        $this->sheet->getColumnDimension('C')->setWidth(21);
        $this->sheet->getColumnDimension('D')->setWidth(6.9);
        $this->sheet->getColumnDimension('A')->setWidth(2.4);
        $j = 8;
        $i = 5;
        while ($i <= 88){
            $column = $this->get_column($i);
            if ($i === $j){
                $this->sheet->getColumnDimension($column)->setWidth(2.4);
                $j = $j + 4;
            }else{
                $this->sheet->getColumnDimension($column)->setWidth(2.2);
            };
            $i++;
        };
    }
    public function to_excel_time($time){ 
        $timestamp = strtotime($time);
        $excelTimestamp = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($timestamp);
        return $excelTimestamp;
    }
    private function merge_cells(){
        $this->sheet->mergeCells("B5:D5");
        $this->sheet->mergeCells("C2:D2");
        $this->sheet->mergeCells("C3:D3");
        $this->sheet->mergeCells("C4:D4");
        $this->sheet->mergeCells("B5:D5");
        $i = 5;
        while ($i <= 88){
            $start = $this->get_column($i) ;
            $end = $this->get_column($i+3);
            $this->sheet->mergeCells($start.'2:'.$end.'2');
            $this->sheet->mergeCells($start.'3:'.$end.'3');
            $this->sheet->mergeCells($start.'4:'.$end.'4');
            $this->sheet->mergeCells($start.'5:'.$end.'5');
            $i = $i + 4;
        };
    }
    private function set_text($column, $text, $bold=true, $size=14){
        $this->sheet->getCell($column)->setValue($text);
        $this->sheet->getStyle($column)
        ->getFont()
        ->setBold($bold)
        ->setSize($size);
    }
    private function add_date(){
        $this->set_text('B2', "MERCREDI");
        $this->set_text('B3', "22 JUILLET 2021 - S29");
        $this->sheet->getStyle('B2')->getAlignment()->setHorizontal('left');
        $this->sheet->getStyle('B3')->getAlignment()->setHorizontal('left');
    }
    private function add_headers(){
        $this->sheet->setCellValue('C2', "Planifié");
        $this->sheet->setCellValue('C3', "Chiffre prévisionel");
        $this->sheet->setCellValue('C4', "Nombre de collaborateurs");
        $this->sheet->setCellValue('B5', "Collaborateurs");
        $this->sheet->getStyle('C2')->getAlignment()->setHorizontal('right');
        $this->sheet->getStyle('C3')->getAlignment()->setHorizontal('right');
        $this->sheet->getStyle('C4')->getAlignment()->setHorizontal('right');
        $this->sheet->getStyle('B5')->getAlignment()->setHorizontal('left');
    }
    public function add_data($data, $row, $is_date){
        $j = 5;
        for ($i = 0; $i <= 20; $i++){
            $start = $this->get_column($j);
            $end = $this->get_column($j+3);
            if ($is_date){
                $this->sheet->setCellValue(
                    "$start$row", 
                    $this->to_excel_time($data[$i])
                );
                $this->sheet->getStyle("$start$row:$end$row")
                ->getNumberFormat()
                ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME3);
            }else{
                $this->sheet->setCellValue("$start$row", $data[$i]);
            }
            $j= $j + 4; 
        }
    }
    public function add_planned(){
        $data = array('00:00:00','00:30:00','03:00:00','04:00:00','04:00:00','04:00:00','04:00:00','04:00:00','03:30:00','05:00:00','05:00:00','03:30:00','03:00:00','03:00:00','03:00:00','03:00:00','03:00:00','03:00:00','02:00:00','01:30:00','00:30:00');
        $this->add_data($data, '2', true);
    }
    public function add_figures(){
        $data = array('0','100','200','300','1430','1630','1900','2140','1800','900','400','500','800','1400','1600','2800','2600','1400','1100','200','0');
        $this->add_data($data, '3', false);
    }
    public function add_collaborators_time(){
        $data = array('06:00:00','07:00:00','08:00:00','09:00:00','10:00:00','11:00:00','12:00:00','13:00:00','14:00:00','15:00:00','16:00:00','17:00:00','18:00:00','19:00:00','20:00:00','21:00:00','22:00:00','23:00:00','00:00:00','01:00:00','02:00:00');
        $this->add_data($data, '5', true);
        $this->sheet->getStyle('B5:CJ5')
        ->getFont()
        ->setSize(12);
    }
    public function add_collaborators_number(){
        $data = [0, [1,1],[1,1],[3,5],[4,4],[4,4],[5,5],[4,4],[8,6],[5,6],[2,4],[5,6],[3,4],[3,4],[3,3],[3,3],[3,3],[3,3],[3,3],[3,3],[3,3]];
        $j = 5;
        for ($i=0; $i<=20; $i++ ){
            $column = $this->get_column($j);
            if ($data[$i] == 0){
                $this->sheet->setCellValue($column.'4',$data[$i]);
                $j = $j + 4;
                continue;
            }
            $value = $data[$i][0].'/'.$data[$i][1];
            if ($data[$i][0] != $data[$i][1]){
                $this->sheet->setCellValue($column.'4',$value);
                $this->make_bg($column.'4',"F4B084");
                $j = $j + 4;
                continue;
            }
            $this->sheet->setCellValue($column.'4',$value);
            $j = $j + 4;
        }
    }
    private function get_index($column){
        $length = strlen($column);
        if ($length === 1 ){
            return array_search($column, $this->alphabet) ;
        }
        $arr = array();
        for ($i=0; $i < $length; $i++){
            array_push($arr, array_search($column[$i], $this->alphabet));
        }
        $index = ($arr[0] + 1) * 26 + $arr[1] ;
        return $index;
    } 
    public function colorize_working_hours($start, $len, $row, $color, $text=null){
        $start_index = $this->get_index($start);
        $end_column = $this->get_column($start_index + $len);
        $this->make_bg("$start$row:$end_column$row", $color);
        $this->update_borders("$start$row:$end_column$row");
        if ($text){
            for ($i=0; $i < strlen($text); $i++){
                $char_column = $this->get_column($start_index+1 + $i);
                $this->set_text(
                    "$char_column$row",
                    $text[$i],
                    false
                );
            }
        }
    } 
    public function add_working_hours(){
        $this->colorize_working_hours("K", 10, 6, "4472C4", "CT");
        $this->colorize_working_hours("U", 16, 6, "D9E1F2", "AB");
        $this->colorize_working_hours("AK", 2, 6, "FFE699", "FT");
        $this->colorize_working_hours("M", 36, 7, "A5A5A5");
        $this->colorize_working_hours("M", 38, 8, "A5A5A5");
        $this->colorize_working_hours("Q", 15, 9, "FFE699","C");
        $this->colorize_working_hours("AF", 5, 9, "92D050","P");
        $this->colorize_working_hours("AK", 16, 9, "E2EFDA","D");
        $this->colorize_working_hours("AO", 36, 10, "A5A5A5");
        $this->colorize_working_hours("AO", 42, 11, "A5A5A5");
        $this->colorize_working_hours("BA", 34, 12, "A5A5A5");
    }
    public function add_collaborators(){
        $collaborators = [
            ["NOM PRENOM","07:30 - 14:30",7],
            ["Salah Saadaoui","08:00 - 17:00",9],
            ["Caddy Dz","08:00 - 17:30",9.5],
            ["John Doe","09:00 - 17:00",9],
            ["SALIM DJERBOUH","15:00 - 23:00",8],
            ["Elon Musk","15:00 - 01:30",10.5],
            ["Jane Doe","18:00 - 02:30",8.5],
        ];
        for ($i = 0; $i <= count($collaborators) - 1 ; $i++ ){
            $row = $i + 6;
            $this->set_text("B$row", $collaborators[$i][0], false, 12);
            $this->set_text("C$row", $collaborators[$i][1], false, 12);
            $this->set_text("D$row", $collaborators[$i][2], false, 12);
        }
    }
    public function index($company, $day){
        $this->update_width();
        $this->update_alignment();
        $this->add_date();
        $this->merge_cells();
        $this->add_headers();
        $this->set_borders();
        $this->update_bg();
        $this->add_planned();        
        $this->add_figures();        
        $this->add_collaborators_number();        
        $this->add_collaborators_time();        
        $this->add_collaborators();        
        $this->add_working_hours();    
        $this->add_outline_borders();    
        $this->sheet->getDefaultRowDimension()->setRowHeight(24.45);
        $writer = new Xlsx($this->spreadsheet);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="'. urlencode("file.xlsx").'"');
        $writer->save('php://output');
    }
}
