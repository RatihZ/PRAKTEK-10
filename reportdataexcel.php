<?php

include 'koneksi.php';
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet    = new Spreadsheet();
$sheet          = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','No');
$sheet->setCellValue('B1','JP');
$sheet->setCellValue('C1','TMS');
$sheet->setCellValue('D1','NIS');
$sheet->setCellValue('E1','NPU');
$sheet->setCellValue('F1','PAUD');
$sheet->setCellValue('G1','TK');
$sheet->setCellValue('H1','SKHUN');
$sheet->setCellValue('I1','IJAZAH');
$sheet->setCellValue('J1','HOBI');
$sheet->setCellValue('K1','CITA-CITA');
$sheet->setCellValue('L1','NAMA');
$sheet->setCellValue('M1','JK');
$sheet->setCellValue('N1','NISN');
$sheet->setCellValue('O1','NIK');
$sheet->setCellValue('P1','TmL');
$sheet->setCellValue('Q1','TgL');
$sheet->setCellValue('R1','AGAMA');
$sheet->setCellValue('S1','KEBUTUHAN KHUSUS');
$sheet->setCellValue('T1','ALAMAT');
$sheet->setCellValue('U1','RT');
$sheet->setCellValue('V1','RW');
$sheet->setCellValue('W1','DUSUN');
$sheet->setCellValue('X1','KECAMATAN');
$sheet->setCellValue('Y1','KODE POS');
$sheet->setCellValue('Z1','TT');
$sheet->setCellValue('AA1','MT');
$sheet->setCellValue('AB1','HP');
$sheet->setCellValue('AC1','TELPON');
$sheet->setCellValue('AD1','EMAIL');
$sheet->setCellValue('AE1','KPS');
$sheet->setCellValue('AF1','Kwn');

$sql = mysqli_query($conn,"SELECT * FROM peserta, pribadi");
$i  = 2;
$no = 1;
while ($row = mysqli_fetch_array($sql)) {
    $sheet->setCellValue('A'.$i,$no++);
    $sheet->setCellValue('B'.$i,$row['jp']);
    $sheet->setCellValue('C'.$i,$row['tanggal']);
    $sheet->setCellValue('D'.$i,$row['nis']);
    $sheet->setCellValue('E'.$i,$row['no_ujian']);
    $sheet->setCellValue('F'.$i,$row['paud']);
    $sheet->setCellValue('G'.$i,$row['tk']);
    $sheet->setCellValue('H'.$i,$row['skhun']);
    $sheet->setCellValue('I'.$i,$row['ijazah']);
    $sheet->setCellValue('J'.$i,$row['hobi']);
    $sheet->setCellValue('K'.$i,$row['cita']);
    $sheet->setCellValue('L'.$i,$row['nama']);
    $sheet->setCellValue('M'.$i,$row['jk']);
    $sheet->setCellValue('N'.$i,$row['nisn']);
    $sheet->setCellValue('O'.$i,$row['nik']);
    $sheet->setCellValue('P'.$i,$row['tempatL']);
    $sheet->setCellValue('Q'.$i,$row['tanggalL']);
    $sheet->setCellValue('R'.$i,$row['agama']);
    $sheet->setCellValue('S'.$i,$row['khusus']);
    $sheet->setCellValue('T'.$i,$row['alamat']);
    $sheet->setCellValue('U'.$i,$row['rt']);
    $sheet->setCellValue('V'.$i,$row['rw']);
    $sheet->setCellValue('W'.$i,$row['dusun']);
    $sheet->setCellValue('X'.$i,$row['kecamatan']);
    $sheet->setCellValue('Y'.$i,$row['kode']);
    $sheet->setCellValue('Z'.$i,$row['tempatT']);
    $sheet->setCellValue('AA'.$i,$row['transpor']);
    $sheet->setCellValue('AB'.$i,$row['hp']);
    $sheet->setCellValue('AC'.$i,$row['telp']);
    $sheet->setCellValue('AD'.$i,$row['email']);
    $sheet->setCellValue('AE'.$i,$row['no_KPS']);
    $sheet->setCellValue('AF'.$i,$row['negara']);
    $i++;
}

$styleArray = [
    'borders'=>[
        'allBorders'=>[
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$i = $i - 1;
$sheet->getStyle('A1:AF'.$i)->applyFromArray($styleArray);

$writer         = new Xlsx($spreadsheet);
$writer->save('Report Registrasi Peserta Didik Baru.Xlsx');
?>