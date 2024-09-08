<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hasil Input Nilai</title>
    <!-- Memuat file CSS eksternal -->
    <link rel="stylesheet" href="style.css">
    <!-- Memuat font Google -->
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
</head>
<body>

<div class="result-container">
<?php
require 'vendor/autoload.php'; // Include PhpSpreadsheet via Composer

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['submit'])) {
    // Ambil data dari form
    $nama = $_POST['nama'];
    $nim = $_POST['nim'];
    $nilai = $_POST['nilai'];

    // Fungsi untuk mengkonversi nilai angka ke nilai huruf
    function konversiNilaiHuruf($nilai) {
        if ($nilai >= 80) {
            return "A";
        } elseif ($nilai >= 71) {
            return "B+";
        } elseif ($nilai >= 65) {
            return "B";
        } elseif ($nilai >= 60) {
            return "C+";
        } elseif ($nilai >= 55) {
            return "C";
        } elseif ($nilai >= 50) {
            return "D+";
        } elseif ($nilai >= 40) {
            return "D";
        } else {
            return "E";
        }
    }

    $nilai_huruf = konversiNilaiHuruf($nilai);

    // Cek apakah file Excel sudah ada
    $filePath = 'nilai_mahasiswa.xlsx';

    if (file_exists($filePath)) {
        // Jika file sudah ada, buka file tersebut
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
    } else {
        // Jika file belum ada, buat file baru
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Nama');
        $sheet->setCellValue('B1', 'NIM');
        $sheet->setCellValue('C1', 'Nilai Angka');
        $sheet->setCellValue('D1', 'Nilai Huruf');
    }

    // Dapatkan sheet aktif
    $sheet = $spreadsheet->getActiveSheet();

    // Cari baris terakhir yang terisi
    $highestRow = $sheet->getHighestRow();

    // Tambahkan data baru ke baris berikutnya
    $sheet->setCellValue('A' . ($highestRow + 1), $nama);
    $sheet->setCellValue('B' . ($highestRow + 1), $nim);
    $sheet->setCellValue('C' . ($highestRow + 1), $nilai);
    $sheet->setCellValue('D' . ($highestRow + 1), $nilai_huruf);

    // Simpan file Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);

    // Tampilkan hasil dengan style yang lebih rapi
    echo "
    <div class='result'>
        <h3>Hasil Input Nilai</h3>
        <div class='result-content'>
            <p>Nama: <strong>$nama</strong></p>
            <p>NIM: <strong>$nim</strong></p>
            <p>Nilai Angka: <strong>$nilai</strong></p>
            <p>Nilai Huruf: <strong>$nilai_huruf</strong></p>
        </div>
        <a href='index.html' class='back-button'>Kembali</a>
    </div>
    ";
}
?>

</div>

</body>
</html>

