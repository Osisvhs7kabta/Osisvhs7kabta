<?php
require 'vendor/autoload.php'; // Memuat autoloader Composer

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Mengecek apakah data sudah dikirimkan dengan metode POST
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Mengambil data dari form
    $nama = $_POST['nama'];
    $kelas = $_POST['kelas'];
    $divisi = $_POST['divisi'];

    // Membuat spreadsheet baru
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Menambahkan header jika file Excel belum ada
    $file = 'absensi_osis.xlsx';
    if (!file_exists($file)) {
        $sheet->setCellValue('A1', 'Nama');
        $sheet->setCellValue('B1', 'Kelas');
        $sheet->setCellValue('C1', 'Divisi');
    }

    // Menambahkan data absensi ke spreadsheet
    $highestRow = $sheet->getHighestRow() + 1; // Menentukan baris kosong berikutnya
    $sheet->setCellValue('A' . $highestRow, $nama);
    $sheet->setCellValue('B' . $highestRow, $kelas);
    $sheet->setCellValue('C' . $highestRow, $divisi);

    // Menulis file Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save($file);

    // Menampilkan pesan terimakasih dengan desain
    echo "
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: Arial, sans-serif;
            background-color: #f0f4f8;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
        }
        .thank-you {
            background-color: #e8f5e9;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
            text-align: center;
            width: 100%;
            max-width: 500px;
        }
        .thank-you h2 {
            color: #4CAF50;
            font-size: 28px;
            margin-bottom: 15px;
        }
        .thank-you p {
            font-size: 18px;
            margin-bottom: 20px;
        }
        .thank-you a {
            color: #4CAF50;
            text-decoration: none;
            font-weight: bold;
            padding: 10px 20px;
            border: 2px solid #4CAF50;
            border-radius: 4px;
            transition: background-color 0.3s ease, color 0.3s ease;
        }
        .thank-you a:hover {
            background-color: #4CAF50;
            color: white;
        }
    </style>
    <div class='thank-you'>
        <h2>Terimakasih Adikku, Absennya Sudah Terdata</h2>
        <p>Data absensi telah berhasil disimpan. Kamu bisa mengunduh arsip absensi berikut:</p>
        <p><a href='$file' download>Unduh Arsip Absensi (Excel)</a></p>
    </div>";
}
?>
