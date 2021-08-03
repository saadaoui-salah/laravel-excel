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
        $this->sheet->getDefaultRowDimension()->setRowHeight(24);
        $this->sheet->getRowDimension(1)->setRowHeight(9.2);
        $this->sheet->getStyle('A:CJ')->getAlignment()->setHorizontal('center');
        $this->sheet->getStyle('A:CJ')->getAlignment()->setVertical('center');
    }
    private function get_column($i){
        $alphabet = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
        if ($i <= 26){
            return $alphabet[$i-1];
        }else{
            $first = floor($i/26);
            $second = $i - 26 * $first ;
            if($second ==0){
                return $alphabet[$first-2 ].''.$alphabet[25];
            };
            return $alphabet[$first-1].''.$alphabet[$second-1];
        }
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
        $outtline = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
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
    private function set_bold_text($column, $text){
        $this->sheet->getCell($column)->setValue($text);
        $this->sheet->getStyle($column)
        ->getFont()
        ->setBold(true)
        ->setSize(14);
    }
    private function add_date(){
        $this->set_bold_text('B2', "MERCREDI");
        $this->set_bold_text('B3', "22 JUILLET 2021 - S29");
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
    public function index(){
        $this->update_width();
        $this->add_date();
        $this->merge_cells();
        $this->add_headers();
        $this->set_borders();
        $this->update_bg();
        
        $writer = new Xlsx($this->spreadsheet);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="'. urlencode("file.xlsx").'"');
        $writer->save('php://output');
    }
}
