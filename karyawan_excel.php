<?php
include('koneksi.php');
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;



$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();


if(isset($_GET['dari']) && isset($_GET['sampai'])){
 $dari = $_GET['dari'];
 $sampai = $_GET['sampai'];
 $sheet->setCellValue('A1', 'No');
 $sheet->setCellValue('B1', 'NAMA LENGKAP');
 $sheet->setCellValue('C1', 'ALAMAT');
 $sheet->setCellValue('D1', 'JENIS KELAMIN');
 $sheet->setCellValue('E1', 'EMAIL');
 $sheet->setCellValue('F1', 'KONTAK');
 $sheet->setCellValue('G1', 'TANGGAL BERGABUNG');

 $data = mysqli_query($koneksi,"select * from karyawan where karyawan_tanggal_bergabung between '$dari' and '$sampai'");
 $i = 2;
 $no = 1;
 while($d = mysqli_fetch_array($data))
 {
    $sheet->setCellValue('A'.$i, $no++);
    $sheet->setCellValue('B'.$i, $d['karyawan_nama']);
    $sheet->setCellValue('C'.$i, $d['karyawan_alamat']);
    $sheet->setCellValue('D'.$i, $d['karyawan_kelamin']);
    $sheet->setCellValue('E'.$i, $d['karyawan_email']);    
    $sheet->setCellValue('F'.$i, $d['karyawan_kontak']);    
    $sheet->setCellValue('G'.$i, date('d-m-Y', strtotime($d['karyawan_tanggal_bergabung'])));
    $i++;
}

$writer = new Xlsx($spreadsheet);
$writer->save('Data karyawan.xlsx');
echo "<script>window.location = 'Data karyawan.xlsx'</script>";

}else{
    

}

?>