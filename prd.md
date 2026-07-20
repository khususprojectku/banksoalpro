# Product Requirements Document (PRD)

## BankSoalPro – Sistem Manajemen Bank Soal Berbasis Web

**Versi:** 1.0
**Status:** Draft
**Platform:** Web (Responsive)
**Target Pengguna:** Sekolah, Guru, MGMP, Bimbingan Belajar, Yayasan

---

# 1. Ringkasan Produk

BankSoalPro adalah aplikasi web yang memungkinkan guru menyusun, mengelola, mencari, dan membagikan bank soal secara terpusat. Sistem mendukung kolaborasi banyak guru, penyusunan paket ujian secara otomatis, analisis soal, serta integrasi dengan CBT dan LMS.

Aplikasi dapat diakses melalui browser tanpa perlu instalasi.

---

# 2. Tujuan

* Menjadi pusat penyimpanan soal sekolah.
* Mempermudah guru membuat soal berkualitas.
* Mengurangi soal duplikat.
* Memudahkan penyusunan ujian.
* Mendukung kolaborasi antar guru.
* Menyediakan analisis kualitas soal.
* Siap digunakan oleh banyak sekolah (multi-tenant).

---

# 3. Target Pengguna

### Super Admin

Mengelola seluruh sistem.

### Admin Sekolah

Mengelola data sekolah, guru, kelas, mata pelajaran.

### Guru

Membuat dan mengelola bank soal.

### Ketua MGMP

Mengelola bank soal tingkat sekolah atau kabupaten.

### Siswa (Opsional)

Mengerjakan latihan atau ujian.

---

# 4. Ruang Lingkup MVP

Versi pertama mencakup:

* Login
* Dashboard
* Manajemen Guru
* Mata Pelajaran
* Kelas
* Bank Soal
* Paket Soal
* Import Excel
* Export Word/PDF
* Pencarian Soal
* Backup Database

---

# 5. Fitur Utama

## 5.1 Autentikasi

* Login Email
* Login Google
* Reset Password
* Verifikasi Email
* OTP (opsional)
* Two Factor Authentication (opsional)

---

## 5.2 Dashboard

Menampilkan statistik:

* Jumlah Guru
* Jumlah Mata Pelajaran
* Jumlah Soal
* Jumlah Paket
* Jumlah Ujian
* Soal Terbaru
* Aktivitas Guru

Grafik:

* Soal per Mata Pelajaran
* Soal per Guru
* Soal per Tingkat Kesulitan

---

## 5.3 Manajemen Sekolah

Data:

* Nama
* NPSN
* Logo
* Alamat
* Email
* Nomor Telepon
* Tahun Ajaran

---

## 5.4 Manajemen Guru

CRUD Guru

Field:

* Nama
* NIP
* Email
* Password
* Mata Pelajaran
* Status

---

## 5.5 Mata Pelajaran

CRUD

Contoh:

* Informatika
* Matematika
* Fisika

---

## 5.6 Kelas

* X
* XI
* XII

---

## 5.7 Bab

Misalnya:

* Algoritma
* Jaringan
* Basis Data
* AI

---

## 5.8 Topik

Sub bab.

---

# 6. Modul Bank Soal

## Jenis Soal

* Pilihan Ganda
* Pilihan Ganda Kompleks
* Benar Salah
* Menjodohkan
* Isian Singkat
* Esai
* Uraian
* Numerik

---

## Setiap Soal Memiliki

* Nomor
* Mata Pelajaran
* Kelas
* Bab
* Topik
* Tingkat Kesulitan
* Kompetensi
* Tujuan Pembelajaran
* Pertanyaan
* Gambar
* Audio
* Video
* Rumus LaTeX
* Pilihan Jawaban
* Jawaban Benar
* Pembahasan
* Referensi
* Tag
* Status
* Pembuat
* Reviewer
* Tanggal Dibuat

---

## Editor Soal

Editor WYSIWYG

Mendukung:

* Gambar
* Tabel
* Rumus Matematika
* Diagram
* Drag and Drop Upload
* Copy Paste dari Word
* Preview

---

# 7. Import Soal

Import dari:

* Excel (.xlsx)
* CSV
* Word (.docx)

Fitur:

* Mapping kolom otomatis
* Validasi
* Preview
* Deteksi duplikat

---

# 8. Export

Export menjadi:

* PDF
* Word
* Excel

Pilihan:

* Dengan Jawaban
* Tanpa Jawaban
* Dengan Pembahasan

---

# 9. Paket Soal

Jenis Paket:

* UTS
* UAS
* PAS
* PAT
* Try Out
* Latihan

---

Pembuatan Paket

Manual

Acak

Berdasarkan:

* Bab
* Topik
* Tingkat Kesulitan
* Kompetensi
* Tag

---

# 10. Generator Soal Acak

Input:

* Jumlah soal
* Bab
* Mata Pelajaran
* Tingkat Kesulitan
* Kompetensi

Output:

Paket soal siap cetak.

---

# 11. Pencarian Soal

Filter:

* Kata Kunci
* Mata Pelajaran
* Guru
* Bab
* Topik
* Tingkat Kesulitan
* Tag
* Tahun

---

# 12. Favorit

Guru dapat memberi bookmark soal.

---

# 13. Riwayat

Menampilkan:

* Soal terbaru
* Soal yang diedit
* Aktivitas pengguna

---

# 14. Analisis Soal

Menampilkan:

* Jumlah penggunaan
* Tingkat kesulitan
* Tingkat diskriminasi
* Persentase benar
* Statistik jawaban

---

# 15. CBT (Versi Berikutnya)

* Jadwal ujian
* Token ujian
* Acak soal
* Acak jawaban
* Timer
* Auto Submit
* Nilai otomatis

---

# 16. Hak Akses

## Super Admin

Seluruh sistem.

## Admin Sekolah

Mengelola sekolah.

Guru.

Mata Pelajaran.

Bank Soal.

Paket.

## Guru

CRUD soal.

Import.

Export.

Paket.

## Reviewer

Approve soal.

Revisi soal.

## Siswa

Latihan.

Ujian.

---

# 17. Notifikasi

* Email
* WhatsApp (opsional)
* In-App Notification

---

# 18. Audit Log

Mencatat:

* Login
* Logout
* Tambah soal
* Edit soal
* Hapus soal
* Export
* Import

---

# 19. Teknologi

## Frontend

* React.js + Vite
* TypeScript
* Tailwind CSS
* Shadcn/UI
* TanStack Query
* React Hook Form
* Chart.js / ApexCharts

## Backend

* Laravel 12
* PHP 8.4
* REST API
* Laravel Sanctum

## Database

* MySQL 8 (struktur dibuat kompatibel untuk migrasi ke PostgreSQL)

## Storage

* Local Storage
* Amazon S3 / Cloudflare R2 (opsional)

## Cache

* Redis

## Queue

* Redis Queue

---

# 20. Non Functional Requirements

* Responsive (Desktop, Tablet, Mobile)
* Multi School (Multi-Tenant)
* Multi User
* HTTPS
* Backup Otomatis
* Restore Database
* Waktu respon halaman < 2 detik untuk operasi umum
* Mendukung minimal 500 pengguna aktif bersamaan (dapat ditingkatkan dengan horizontal scaling)

---

# 21. Roadmap

### Versi 1.0

* Login & manajemen pengguna
* Bank soal
* Paket soal
* Import/Export
* Pencarian

### Versi 1.5

* AI pembuat soal dari materi
* AI pembuat pembahasan
* OCR soal dari foto/PDF
* Deteksi soal duplikat

### Versi 2.0

* CBT lengkap
* Analisis butir soal
* Dashboard kepala sekolah
* Integrasi LMS
* API publik

---

## Catatan Arsitektur

Untuk kebutuhan yang pernah Anda jelaskan (banyak guru, kemungkinan di-host di Hostinger, dan ingin mudah dikembangkan), arsitektur yang disarankan adalah:

* **Frontend:** React + TypeScript
* **Backend:** Laravel 12 (REST API)
* **Database:** MySQL (dengan skema yang kompatibel untuk migrasi ke PostgreSQL)
* **Autentikasi:** Laravel Sanctum
* **Deployment:** Nginx + PHP-FPM + MySQL
* **Penyimpanan berkas:** Local Storage (MVP), dapat ditingkatkan ke S3/R2 di masa depan

Arsitektur ini memisahkan frontend dan backend sehingga lebih mudah dikembangkan, lebih fleksibel untuk integrasi aplikasi lain di masa depan, dan memudahkan migrasi database dari MySQL ke PostgreSQL jika diperlukan.
