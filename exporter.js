// exporter.js - Export libraries for Word, Excel, and PDF
// Implemented as an ES Module

// Helper to strip HTML tags for simple text output
function stripHtml(html) {
  if (!html) return '';
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return doc.body.textContent || "";
}

// 1. Export Question Package to MS Word (.doc)
export function exportToWord(pkg, questions, school, options = {}) {
  const { includeAnswer = false, includeDiscussion = false } = options;
  const filename = `${pkg.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}`;
  
  let contentHtml = `
    <div class="kop-surat">
      <div style="font-size: 16pt; font-weight: bold;">${school.name.toUpperCase()}</div>
      <div style="font-size: 11pt;">NPSN: ${school.npsn} | Telp: ${school.phone} | Email: ${school.email}</div>
      <div style="font-size: 11pt;">${school.address}</div>
      <div style="margin-top: 10px; border-bottom: 3px double #000; width: 100%;"></div>
    </div>
    
    <table style="width: 100%; margin-top: 15px; margin-bottom: 25px; font-size: 11pt; border: none;">
      <tr>
        <td style="width: 15%; font-weight: bold;">MATA PELAJARAN</td>
        <td style="width: 35%;">: ${pkg.subjectName || ''}</td>
        <td style="width: 15%; font-weight: bold;">HARI / TANGGAL</td>
        <td style="width: 35%;">: .........................</td>
      </tr>
      <tr>
        <td style="font-weight: bold;">KELAS / SEMESTER</td>
        <td>: ${pkg.className || ''} / Ganjil</td>
        <td style="font-weight: bold;">WAKTU</td>
        <td>: ......................... Menit</td>
      </tr>
      <tr>
        <td style="font-weight: bold;">JENIS UJIAN</td>
        <td>: ${pkg.type}</td>
        <td style="font-weight: bold;">TAHUN AJARAN</td>
        <td>: ${school.academicYear}</td>
      </tr>
    </table>
    
    <div style="border: 1px solid #000; padding: 10px; margin-bottom: 20px; font-size: 10pt;">
      <strong>PETUNJUK UMUM:</strong><br>
      1. Tuliskan nama dan nomor peserta Anda pada lembar jawaban yang disediakan.<br>
      2. Periksa dan bacalah soal-soal dengan teliti sebelum Anda menjawabnya.<br>
      3. Dahulukan menjawab soal-soal yang Anda anggap mudah.<br>
      4. Laporkan kepada pengawas ujian apabila terdapat lembar soal yang kurang jelas, rusak, atau tidak lengkap.
    </div>

    <h3 style="text-align: center; text-transform: uppercase;">SOAL UJIAN</h3>
  `;

  questions.forEach((q, idx) => {
    contentHtml += `
      <div class="question" style="margin-bottom: 20px; page-break-inside: avoid;">
        <table style="width: 100%; border: none;">
          <tr style="vertical-align: top;">
            <td style="width: 3%;">${idx + 1}.</td>
            <td style="width: 97%;">
              <div>${q.questionText}</div>
    `;

    // Render Choices if Multiple Choice or true/false
    if (q.type === 'PG' && q.choices) {
      contentHtml += `<table style="width: 100%; margin-top: 8px; border: none; margin-left: 15px;">`;
      Object.entries(q.choices).forEach(([key, text]) => {
        contentHtml += `
          <tr style="vertical-align: top;">
            <td style="width: 4%; font-weight: bold;">${key}.</td>
            <td>${text}</td>
          </tr>
        `;
      });
      contentHtml += `</table>`;
    } else if (q.type === 'PG_KOMPLEKS' && q.choices) {
      contentHtml += `
        <div style="font-size: 9pt; color: #555; margin-top: 5px; font-style: italic;">(Pilihan Ganda Kompleks - Pilihlah jawaban-jawaban yang benar)</div>
        <table style="width: 100%; margin-top: 8px; border: none; margin-left: 15px;">;
      `;
      Object.entries(q.choices).forEach(([key, text]) => {
        contentHtml += `
          <tr style="vertical-align: top;">
            <td style="width: 4%; font-weight: bold;">[  ] ${key}.</td>
            <td>${text}</td>
          </tr>
        `;
      });
      contentHtml += `</table>`;
    } else if (q.type === 'BENAR_SALAH') {
      contentHtml += `
        <div style="margin-top: 8px; margin-left: 15px;">
          [ &nbsp; ] BENAR &nbsp;&nbsp;&nbsp;&nbsp; [ &nbsp; ] SALAH
        </div>
      `;
    } else if (q.type === 'MENJODOHKAN' && q.choices) {
      contentHtml += `
        <div style="font-size: 9pt; color: #555; margin-top: 5px; font-style: italic;">(Menjodohkan - Pasangkanlah pernyataan kiri dengan jawaban kanan)</div>
        <table style="width: 80%; margin-top: 8px; border-collapse: collapse; margin-left: 15px; border: 1px solid #000;">
          <thead>
            <tr style="background: #e0e0e0;">
              <th style="border: 1px solid #000; padding: 5px; text-align: left;">Pernyataan</th>
              <th style="border: 1px solid #000; padding: 5px; text-align: left;">Pilihan Jawaban</th>
            </tr>
          </thead>
          <tbody>
      `;
      const maxLen = Math.max(q.choices.premises?.length || 0, q.choices.responses?.length || 0);
      for (let i = 0; i < maxLen; i++) {
        contentHtml += `
          <tr>
            <td style="border: 1px solid #000; padding: 5px;">${q.choices.premises?.[i] || ''}</td>
            <td style="border: 1px solid #000; padding: 5px;">${q.choices.responses?.[i] || ''}</td>
          </tr>
        `;
      }
      contentHtml += `
          </tbody>
        </table>
      `;
    } else if (q.type === 'ISIAN_SINGKAT') {
      contentHtml += `
        <div style="margin-top: 8px; margin-left: 15px;">
          Jawab: ....................................................................................................
        </div>
      `;
    } else if (q.type === 'ESSAY' || q.type === 'URAIAN') {
      contentHtml += `
        <div style="margin-top: 15px; height: 100px; border-bottom: 1px dashed #777; width: 100%;"></div>
      `;
    } else if (q.type === 'NUMERIK') {
      contentHtml += `
        <div style="margin-top: 8px; margin-left: 15px;">
          Jawaban Angka: [ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ]
        </div>
      `;
    }

    // Answers and discussions if chosen
    if (includeAnswer) {
      let formattedAnswer = q.correctAnswer;
      if (Array.isArray(q.correctAnswer)) {
        formattedAnswer = q.correctAnswer.join(', ');
      } else if (typeof q.correctAnswer === 'object') {
        formattedAnswer = JSON.stringify(q.correctAnswer);
      }
      contentHtml += `
        <div style="margin-top: 10px; color: #2c5282; font-weight: bold; font-size: 10pt;">
          Kunci Jawaban: ${formattedAnswer}
        </div>
      `;
    }
    if (includeDiscussion && q.discussion) {
      contentHtml += `
        <div style="margin-top: 5px; background-color: #f7fafc; border-left: 3px solid #4a5568; padding: 5px 10px; font-style: italic; font-size: 10pt;">
          <strong>Pembahasan:</strong> ${q.discussion}
        </div>
      `;
    }

    contentHtml += `
            </td>
          </tr>
        </table>
      </div>
    `;
  });

  const completeHtml = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <title>${pkg.name}</title>
      <!--[if gte mso 9]>
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
      </xml>
      <![endif]-->
      <style>
        @page Section1 {
          size: 595.3pt 841.9pt; /* A4 size */
          margin: 1.2in 1.0in 1.2in 1.0in;
          mso-header-margin: .5in;
          mso-footer-margin: .5in;
          mso-paper-source: 0;
        }
        div.Section1 { page: Section1; }
        body {
          font-family: 'Times New Roman', Times, serif;
          font-size: 11pt;
          line-height: 1.3;
          color: #000;
        }
        .kop-surat {
          text-align: center;
          margin-bottom: 10px;
        }
      </style>
    </head>
    <body>
      <div class="Section1">
        ${contentHtml}
      </div>
    </body>
    </html>
  `;

  // Create file blob and download
  const blob = new Blob(['\ufeff' + completeHtml], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${filename}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// 2. Export Package Questions to Excel using SheetJS (XLSX)
export function exportToExcel(pkg, questions) {
  const filename = `${pkg.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}`;
  
  // Format question rows
  const rows = questions.map((q, idx) => {
    let choicesText = '';
    if (q.choices && q.type === 'PG') {
      choicesText = Object.entries(q.choices).map(([k, v]) => `${k}. ${stripHtml(v)}`).join(' | ');
    } else if (q.choices && q.type === 'PG_KOMPLEKS') {
      choicesText = Object.entries(q.choices).map(([k, v]) => `${k}. ${stripHtml(v)}`).join(' | ');
    } else if (q.choices && q.type === 'MENJODOHKAN') {
      const premises = q.choices.premises?.join(', ') || '';
      const responses = q.choices.responses?.join(', ') || '';
      choicesText = `Kiri: [${premises}] | Kanan: [${responses}]`;
    }

    let answerText = q.correctAnswer;
    if (Array.isArray(q.correctAnswer)) {
      answerText = q.correctAnswer.join(', ');
    } else if (typeof q.correctAnswer === 'object') {
      answerText = JSON.stringify(q.correctAnswer);
    }

    return {
      'No': idx + 1,
      'Kode Soal': q.code,
      'Mata Pelajaran': pkg.subjectName || '',
      'Kelas': pkg.className || '',
      'Bab': q.chapterName || '',
      'Tipe': q.type,
      'Tingkat Kesulitan': q.difficulty,
      'Pertanyaan': stripHtml(q.questionText),
      'Pilihan Jawaban (PG)': choicesText,
      'Jawaban Benar / Kunci': answerText,
      'Pembahasan': stripHtml(q.discussion),
      'Kompetensi': q.competence || '',
      'Tujuan Pembelajaran': q.learningGoal || '',
      'Referensi': q.references || '',
      'Tag': q.tag || '',
      'Pembuat': q.creatorName || ''
    };
  });

  // Check if XLSX library is loaded globally
  if (window.XLSX) {
    const ws = window.XLSX.utils.json_to_sheet(rows);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, 'Bank Soal');
    
    // Auto-fit column widths
    const maxCols = Object.keys(rows[0] || {}).length;
    ws['!cols'] = Array(maxCols).fill({ wch: 15 });
    ws['!cols'][7] = { wch: 40 }; // Column Pertanyaan
    ws['!cols'][8] = { wch: 30 }; // Column Choices
    ws['!cols'][9] = { wch: 25 }; // Column Answer
    ws['!cols'][10] = { wch: 30 }; // Column Discussion

    window.XLSX.writeFile(wb, `${filename}.xlsx`);
  } else {
    // Fallback to CSV export
    console.warn('XLSX library not loaded. Falling back to CSV export.');
    exportToCSV(rows, filename);
  }
}

// Helper to export CSV
function exportToCSV(rows, filename) {
  if (rows.length === 0) return;
  const headers = Object.keys(rows[0]);
  const csvContent = [
    headers.join(','),
    ...rows.map(row => 
      headers.map(header => {
        let val = row[header] === null || row[header] === undefined ? '' : String(row[header]);
        // Escape double quotes
        val = val.replace(/"/g, '""');
        // Wrap in quotes if it contains comma, quotes, or newline
        if (val.includes(',') || val.includes('"') || val.includes('\n')) {
          val = `"${val}"`;
        }
        return val;
      }).join(',')
    )
  ].join('\n');

  const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${filename}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// 3. Export Package Questions to PDF via Browser Print
export function exportToPDF(pkg, questions, school, options = {}) {
  const { includeAnswer = false, includeDiscussion = false } = options;
  
  // Create a clean new window/tab for printing
  const printWindow = window.open('', '_blank');
  
  let questionsHtml = '';
  questions.forEach((q, idx) => {
    let choicesHtml = '';
    
    if (q.type === 'PG' && q.choices) {
      choicesHtml += `<table class="choices-table">`;
      Object.entries(q.choices).forEach(([key, text]) => {
        choicesHtml += `
          <tr>
            <td class="choice-key">${key}.</td>
            <td class="choice-val">${text}</td>
          </tr>
        `;
      });
      choicesHtml += `</table>`;
    } else if (q.type === 'PG_KOMPLEKS' && q.choices) {
      choicesHtml += `
        <div class="tip-soal">(Pilihlah beberapa jawaban yang benar)</div>
        <table class="choices-table">
      `;
      Object.entries(q.choices).forEach(([key, text]) => {
        choicesHtml += `
          <tr>
            <td class="choice-key">[ &nbsp; ] ${key}.</td>
            <td class="choice-val">${text}</td>
          </tr>
        `;
      });
      choicesHtml += `</table>`;
    } else if (q.type === 'BENAR_SALAH') {
      choicesHtml += `
        <div class="bs-choices">
          [ &nbsp;&nbsp; ] BENAR &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; [ &nbsp;&nbsp; ] SALAH
        </div>
      `;
    } else if (q.type === 'MENJODOHKAN' && q.choices) {
      choicesHtml += `
        <div class="tip-soal">(Pasangkanlah pernyataan kiri dengan pilihan jawaban kanan)</div>
        <table class="match-table">
          <thead>
            <tr>
              <th>Pernyataan</th>
              <th>Pilihan Jawaban</th>
            </tr>
          </thead>
          <tbody>
      `;
      const maxLen = Math.max(q.choices.premises?.length || 0, q.choices.responses?.length || 0);
      for (let i = 0; i < maxLen; i++) {
        choicesHtml += `
          <tr>
            <td>${q.choices.premises?.[i] || ''}</td>
            <td>${q.choices.responses?.[i] || ''}</td>
          </tr>
        `;
      }
      choicesHtml += `
          </tbody>
        </table>
      `;
    } else if (q.type === 'ISIAN_SINGKAT') {
      choicesHtml += `
        <div class="isian-blank">
          Jawab: ........................................................................................................................
        </div>
      `;
    } else if (q.type === 'ESSAY' || q.type === 'URAIAN') {
      choicesHtml += `
        <div class="essay-box"></div>
      `;
    } else if (q.type === 'NUMERIK') {
      choicesHtml += `
        <div class="isian-blank">
          Jawaban Angka: [ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ]
        </div>
      `;
    }

    let ansHtml = '';
    if (includeAnswer) {
      let formattedAnswer = q.correctAnswer;
      if (Array.isArray(q.correctAnswer)) {
        formattedAnswer = q.correctAnswer.join(', ');
      } else if (typeof q.correctAnswer === 'object') {
        formattedAnswer = JSON.stringify(q.correctAnswer);
      }
      ansHtml += `
        <div class="kunci-jawaban">
          Kunci Jawaban: ${formattedAnswer}
        </div>
      `;
    }
    if (includeDiscussion && q.discussion) {
      ansHtml += `
        <div class="pembahasan">
          <strong>Pembahasan:</strong> ${q.discussion}
        </div>
      `;
    }

    questionsHtml += `
      <div class="question-container">
        <div class="q-num">${idx + 1}.</div>
        <div class="q-body">
          <div class="q-text">${q.questionText}</div>
          ${choicesHtml}
          ${ansHtml}
        </div>
      </div>
    `;
  });

  printWindow.document.write(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>${pkg.name}</title>
      <meta charset="utf-8">
      <!-- MathJax for rendering LaTeX formulas -->
      <script>
        window.MathJax = {
          tex: {
            inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
            displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']]
          },
          svg: {
            fontCache: 'global'
          }
        };
      </script>
      <script type="text/javascript" id="MathJax-script" async
        src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js">
      </script>
      <style>
        body {
          font-family: 'Times New Roman', Times, serif;
          margin: 0;
          padding: 30px;
          color: #000;
          font-size: 12pt;
          line-height: 1.4;
        }
        .kop-surat {
          text-align: center;
          border-bottom: 3px double #000;
          padding-bottom: 10px;
          margin-bottom: 15px;
        }
        .kop-title {
          font-size: 16pt;
          font-weight: bold;
          text-transform: uppercase;
        }
        .kop-info {
          font-size: 10pt;
          margin-top: 4px;
        }
        .info-table {
          width: 100%;
          margin-bottom: 20px;
          font-size: 11pt;
        }
        .info-table td {
          padding: 3px 0;
        }
        .instructions {
          border: 1px solid #000;
          padding: 10px;
          font-size: 10pt;
          margin-bottom: 25px;
        }
        .sheet-title {
          text-align: center;
          font-weight: bold;
          font-size: 14pt;
          text-transform: uppercase;
          margin-bottom: 20px;
        }
        .question-container {
          display: flex;
          margin-bottom: 20px;
          page-break-inside: avoid;
        }
        .q-num {
          width: 25px;
          font-weight: bold;
        }
        .q-body {
          flex: 1;
        }
        .q-text {
          margin-bottom: 8px;
        }
        .choices-table {
          width: 100%;
          border-collapse: collapse;
          margin-left: 10px;
          margin-top: 5px;
        }
        .choices-table td {
          padding: 3px 0;
          vertical-align: top;
        }
        .choice-key {
          width: 30px;
          font-weight: bold;
        }
        .choice-val {
          padding-left: 5px;
        }
        .tip-soal {
          font-size: 9pt;
          color: #444;
          font-style: italic;
          margin-top: 3px;
          margin-left: 10px;
        }
        .bs-choices {
          margin-top: 8px;
          margin-left: 10px;
        }
        .match-table {
          width: 90%;
          border-collapse: collapse;
          margin: 8px 0 8px 10px;
        }
        .match-table th, .match-table td {
          border: 1px solid #000;
          padding: 6px 10px;
          text-align: left;
        }
        .match-table th {
          background-color: #f2f2f2;
        }
        .isian-blank {
          margin-top: 8px;
          margin-left: 10px;
        }
        .essay-box {
          margin-top: 15px;
          height: 100px;
          border-bottom: 1px dashed #aaa;
          width: 100%;
        }
        .kunci-jawaban {
          margin-top: 8px;
          color: #1a365d;
          font-weight: bold;
          font-size: 10pt;
        }
        .pembahasan {
          margin-top: 4px;
          background-color: #f7fafc;
          border-left: 3px solid #718096;
          padding: 5px 10px;
          font-style: italic;
          font-size: 10pt;
        }
        @media print {
          body {
            padding: 0;
          }
          .no-print {
            display: none;
          }
        }
      </style>
    </head>
    <body>
      <div class="no-print" style="background: #ebf8ff; border: 1px solid #90cdf4; padding: 12px; margin-bottom: 20px; text-align: center; border-radius: 4px;">
        <button onclick="window.print()" style="background: #3182ce; color: #fff; border: none; padding: 8px 16px; font-weight: bold; border-radius: 4px; cursor: pointer; font-size: 11pt;">Cetak Dokumen / Simpan PDF</button>
        <span style="margin-left: 15px; font-size: 10pt; color: #2b6cb0;">Tips: Nonaktifkan header & footer browser di setelan cetak agar lebih rapi.</span>
      </div>

      <div class="kop-surat">
        <div class="kop-title">${school.name}</div>
        <div class="kop-info">NPSN: ${school.npsn} | Alamat: ${school.address}</div>
        <div class="kop-info">Email: ${school.email} | Telp: ${school.phone}</div>
      </div>

      <table class="info-table">
        <tr>
          <td style="width: 15%; font-weight: bold;">Mata Pelajaran</td>
          <td style="width: 35%;">: ${pkg.subjectName || ''}</td>
          <td style="width: 15%; font-weight: bold;">Hari / Tanggal</td>
          <td style="width: 35%;">: .......................................</td>
        </tr>
        <tr>
          <td style="font-weight: bold;">Kelas / Semester</td>
          <td>: ${pkg.className || ''} / Ganjil</td>
          <td style="font-weight: bold;">Waktu Ujian</td>
          <td>: .......................................</td>
        </tr>
        <tr>
          <td style="font-weight: bold;">Jenis Ujian</td>
          <td>: ${pkg.type}</td>
          <td style="font-weight: bold;">Tahun Ajaran</td>
          <td>: ${school.academicYear}</td>
        </tr>
      </table>

      <div class="instructions">
        <strong>PETUNJUK UMUM:</strong><br>
        1. Tuliskan nama, nomor peserta, dan kelas Anda pada lembar jawaban yang disediakan.<br>
        2. Bacalah dengan teliti setiap butir soal sebelum Anda mulai mengerjakannya.<br>
        3. Kerjakan terlebih dahulu soal yang Anda anggap paling mudah.<br>
        4. Tanyakan kepada pengawas apabila terdapat tulisan atau gambar yang kurang jelas.
      </div>

      <div class="sheet-title">SOAL UJIAN</div>

      <div class="questions-list">
        ${questionsHtml}
      </div>

      <!-- Trigger MathJax typesetting after load -->
      <script>
        window.addEventListener('load', () => {
          if (window.MathJax && window.MathJax.typeset) {
            window.MathJax.typeset();
          }
        });
      </script>
    </body>
    </html>
  `);
  
  printWindow.document.close();
}
