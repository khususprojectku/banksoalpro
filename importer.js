// importer.js - Excel and CSV Importer using SheetJS (XLSX)
// Implemented as an ES Module

// Parse Excel or CSV file to JSON array
export function parseExcelOrCSV(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error('Pustaka parsing Excel (SheetJS) belum dimuat. Periksa koneksi internet Anda atau muat ulang halaman.'));
      return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = window.XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert sheet to JSON array of objects
        const rawRows = window.XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        resolve(processRawRows(rawRows));
      } catch (err) {
        reject(new Error(`Gagal membaca berkas Excel: ${err.message}`));
      }
    };

    reader.onerror = function() {
      reject(new Error('Gagal membaca berkas.'));
    };

    reader.readAsArrayBuffer(file);
  });
}

// Map columns and validate rows
function processRawRows(rawRows) {
  const parsedQuestions = [];
  const errors = [];
  const duplicateCodes = new Set();

  rawRows.forEach((row, index) => {
    const rowNum = index + 2; // Row number in sheet (1-based + 1 for header)
    
    // Normalise column names
    const rawCode = String(row['Kode Soal'] || '').trim();
    const rawType = String(row['Tipe Soal'] || '').trim().toUpperCase();
    const rawDifficulty = String(row['Tingkat Kesulitan'] || '').trim().toUpperCase();
    const rawQuestion = String(row['Pertanyaan'] || '').trim();
    const rawAnswer = String(row['Jawaban Benar'] || '').trim();
    const rawDiscussion = String(row['Pembahasan'] || '').trim();
    const rawCompetence = String(row['Kompetensi'] || '').trim();
    const rawGoal = String(row['Tujuan Pembelajaran'] || '').trim();
    const rawRef = String(row['Referensi'] || '').trim();
    const rawTag = String(row['Tag'] || '').trim();

    // Skip empty rows
    if (!rawType && !rawDifficulty && !rawQuestion && !rawAnswer) {
      return;
    }

    const rowErrors = [];

    // Validasi Kolom Wajib
    if (!rawType) rowErrors.push('Tipe Soal wajib diisi.');
    if (!rawDifficulty) rowErrors.push('Tingkat Kesulitan wajib diisi.');
    if (!rawQuestion) rowErrors.push('Teks Pertanyaan wajib diisi.');
    if (!rawAnswer) rowErrors.push('Jawaban Benar wajib diisi.');

    // Validasi Tipe Soal
    const validTypes = ['PG', 'PG_KOMPLEKS', 'BENAR_SALAH', 'ISIAN_SINGKAT', 'ESSAY', 'URAIAN', 'NUMERIK'];
    if (rawType && !validTypes.includes(rawType)) {
      rowErrors.push(`Tipe Soal '${rawType}' tidak valid. Harus salah satu dari: ${validTypes.join(', ')}`);
    }

    // Validasi Tingkat Kesulitan
    const validDiffs = ['MUDAH', 'SEDANG', 'SULIT'];
    if (rawDifficulty && !validDiffs.includes(rawDifficulty)) {
      rowErrors.push(`Tingkat Kesulitan '${rawDifficulty}' tidak valid. Harus salah satu dari: ${validDiffs.join(', ')}`);
    }

    // Validasi & Mapping Pilihan Jawaban (khusus PG dan PG_KOMPLEKS)
    let choices = {};
    if (rawType === 'PG' || rawType === 'PG_KOMPLEKS') {
      const optA = String(row['Pilihan A'] || '').trim();
      const optB = String(row['Pilihan B'] || '').trim();
      const optC = String(row['Pilihan C'] || '').trim();
      const optD = String(row['Pilihan D'] || '').trim();
      const optE = String(row['Pilihan E'] || '').trim();

      if (!optA || !optB) {
        rowErrors.push('Untuk pilihan ganda, minimal Pilihan A dan Pilihan B wajib diisi.');
      } else {
        choices.A = optA;
        choices.B = optB;
        if (optC) choices.C = optC;
        if (optD) choices.D = optD;
        if (optE) choices.E = optE;
      }
    } else if (rawType === 'BENAR_SALAH') {
      choices = {
        A: 'BENAR',
        B: 'SALAH'
      };
      if (rawAnswer !== 'BENAR' && rawAnswer !== 'SALAH' && rawAnswer !== 'A' && rawAnswer !== 'B') {
        rowErrors.push("Jawaban benar untuk BENAR_SALAH harus berupa 'BENAR' atau 'SALAH'.");
      }
    }

    // Normalisasi format jawaban benar
    let correctAnswer = rawAnswer;
    if (rawType === 'PG_KOMPLEKS') {
      // Split comma-separated list of keys, e.g. "A, C, D" -> ["A", "C", "D"]
      correctAnswer = rawAnswer.split(/[,;\s]+/).map(a => a.trim().toUpperCase()).filter(Boolean);
      if (correctAnswer.length === 0) {
        rowErrors.push('Jawaban Benar untuk Pilihan Ganda Kompleks tidak boleh kosong.');
      }
    } else if (rawType === 'BENAR_SALAH') {
      if (rawAnswer === 'A') correctAnswer = 'BENAR';
      if (rawAnswer === 'B') correctAnswer = 'SALAH';
    }

    // Simpan data jika tidak ada error fatal
    if (rowErrors.length > 0) {
      errors.push({
        row: rowNum,
        code: rawCode || `Baris ${rowNum}`,
        messages: rowErrors
      });
    } else {
      // Deteksi duplikat kode di file yang sama
      if (rawCode) {
        if (duplicateCodes.has(rawCode)) {
          errors.push({
            row: rowNum,
            code: rawCode,
            messages: [`Duplikasi kode soal '${rawCode}' terdeteksi di dalam file yang diunggah.`]
          });
        } else {
          duplicateCodes.add(rawCode);
        }
      }

      parsedQuestions.push({
        code: rawCode || `Q-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        type: rawType,
        difficulty: rawDifficulty,
        questionText: rawQuestion.startsWith('<') ? rawQuestion : `<p>${rawQuestion}</p>`,
        choices,
        correctAnswer,
        discussion: rawDiscussion ? (rawDiscussion.startsWith('<') ? rawDiscussion : `<p>${rawDiscussion}</p>`) : '',
        competence: rawCompetence,
        learningGoal: rawGoal,
        references: rawRef,
        tag: rawTag,
        status: 'DRAFT' // Default status for imported questions
      });
    }
  });

  return {
    questions: parsedQuestions,
    errors: errors
  };
}

// 2. Generate and download template Excel for Import
export function downloadExcelTemplate() {
  if (!window.XLSX) {
    alert('Pustaka Excel (SheetJS) belum siap. Pastikan Anda terhubung ke internet.');
    return;
  }

  // Header column names
  const headers = [
    'Kode Soal', 'Tipe Soal', 'Tingkat Kesulitan', 'Pertanyaan', 
    'Pilihan A', 'Pilihan B', 'Pilihan C', 'Pilihan D', 'Pilihan E', 
    'Jawaban Benar', 'Pembahasan', 'Kompetensi', 'Tujuan Pembelajaran', 'Referensi', 'Tag'
  ];

  // Example data rows
  const sampleData = [
    [
      'INF-X-101', 'PG', 'MUDAH', 
      'Apakah fungsi simbol berbentuk persegi panjang dalam flowchart?',
      'Proses perhitungan', 'Pengambilan keputusan', 'Input/Output data', 'Mulai/Selesai', '',
      'A', 'Simbol persegi panjang merepresentasikan operasi aritmatika atau manipulasi data.',
      'Memahami simbol flowchart', 'Siswa dapat menentukan kegunaan simbol persegi panjang',
      'Buku Informatika Kurmer Kelas X Hlm 45', 'flowchart,algoritma'
    ],
    [
      'INF-X-102', 'PG_KOMPLEKS', 'SEDANG', 
      'Manakah yang termasuk komponen utama CPU? (Pilih semua yang benar)',
      'ALU (Arithmetic Logic Unit)', 'CU (Control Unit)', 'RAM (Random Access Memory)', 'Register', 'SSD',
      'A, B, D', 'CPU terdiri dari ALU, Control Unit, dan Register. RAM dan SSD adalah media penyimpanan eksternal dari CPU.',
      'Memahami arsitektur komputer', 'Siswa dapat membedakan unit di dalam CPU',
      'Modul Hardware Kelas X Hlm 12', 'hardware,cpu'
    ],
    [
      'INF-X-103', 'BENAR_SALAH', 'MUDAH', 
      'Bahasa pemrograman Python memerlukan kompilator (compiler) terpisah untuk menghasilkan file biner .exe sebelum dapat dijalankan.',
      '', '', '', '', '',
      'SALAH', 'Python adalah bahasa berjenis interpreter, kode dijalankan baris per baris tanpa kompilasi ke file .exe terlebih dahulu.',
      'Memahami konsep interpreter', 'Siswa mengerti cara kerja interpreter Python',
      'Panduan Python Kelas X', 'python,interpreter'
    ],
    [
      'INF-X-104', 'ISIAN_SINGKAT', 'SEDANG', 
      'Apakah nama protokol jaringan standar yang digunakan untuk mentransfer halaman web secara aman?',
      '', '', '', '', '',
      'HTTPS', 'HTTPS (Hypertext Transfer Protocol Secure) merupakan protokol terenkripsi menggunakan SSL/TLS.',
      'Memahami protokol jaringan', 'Siswa mengenal protokol HTTPS',
      'Keamanan Jaringan Hlm 90', 'jaringan,keamanan'
    ],
    [
      'INF-X-105', 'ESSAY', 'SULIT', 
      'Jelaskan cara kerja teknik pencarian Binary Search dan apa syarat mutlak kumpulan data agar algoritma ini dapat diterapkan!',
      '', '', '', '', '',
      'Syarat mutlak data harus sudah terurut. Cara kerjanya dengan membagi data menjadi dua bagian, lalu membandingkan nilai tengah dengan data yang dicari, lalu membuang setengah bagian yang tidak memenuhi secara berulang.',
      'Binary Search cara kerjanya kompleks.',
      'Menganalisis efisiensi algoritma', 'Siswa dapat mengimplementasikan binary search',
      'Buku Algoritma Pemrograman Hlm 120', 'algoritma,pencarian'
    ]
  ];

  // Combine headers and sample data
  const sheetData = [headers, ...sampleData];

  // Create workbook and sheet
  const ws = window.XLSX.utils.aoa_to_sheet(sheetData);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, 'Template Impor');

  // Adjust column widths
  ws['!cols'] = [
    { wch: 12 }, // Kode Soal
    { wch: 12 }, // Tipe Soal
    { wch: 15 }, // Kesulitan
    { wch: 45 }, // Pertanyaan
    { wch: 20 }, // Pilihan A
    { wch: 20 }, // Pilihan B
    { wch: 20 }, // Pilihan C
    { wch: 20 }, // Pilihan D
    { wch: 20 }, // Pilihan E
    { wch: 15 }, // Jawaban Benar
    { wch: 35 }, // Pembahasan
    { wch: 25 }, // Kompetensi
    { wch: 25 }, // Tujuan Pembelajaran
    { wch: 20 }, // Referensi
    { wch: 15 }  // Tag
  ];

  // Download
  window.XLSX.writeFile(wb, 'template_impor_soal_banksoalpro.xlsx');
}

