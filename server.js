const express = require('express');
const cors = require('cors');
const exceljs = require('exceljs');
const fs = require('fs');
const pdfkit = require('pdfkit');
const moment = require('moment');
require('moment/locale/id');

const app = express();

app.use(cors());

app.use(require('body-parser').json());
app.use(require('body-parser').urlencoded({ extended: true }));
app.use(express.json({ limit: '1mb' }));

const getCurrentTimestamp = () => {
    const now = new Date();

    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');

    const timestamp = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    return timestamp;
};

// Mengambil tanggal hari ini dan memformatnya menjadi "DD MMMM YYYY"
const today = moment().format('DD MMMM YYYY');

app.post('/search', (req, res) => {
    const nim = req.body.nim;

    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile('wisuda.xlsx')
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let student = [];
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber !== 1) {
                    const nimFromSheet = row.getCell(2).value.toString();

                    if (nimFromSheet === nim) {
                        student.push({
                            "No": row.getCell(1).value,
                            "NIM": nimFromSheet,
                            "Nama": row.getCell(3).value,
                            "Program Studi": row.getCell(4).value,
                            "Fakultas": row.getCell(5).value,
                            "Ukuran Almamater": row.getCell(10).value,
                            "Nomor Urut": row.getCell(12).value,
                            "Status Tagihan Wisuda": row.getCell(11).value
                        })
                        return false;
                    }
                }
            });

            if (student.length > 0) {
                res.json({ "found": true, "mahasiswa": student });
            } else {
                res.json({ "message": `NIM: ${nim} tidak terdaftar` });
            }
        })
        .catch(err => {
            console.error(err);
            res.status(500).send('Error reading Excel file');
        });
});

app.post('/download', (req, res) => {
    const nim = req.body.nim;

    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile('wisuda.xlsx')
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let student = null;
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber !== 1) {
                    const nimFromSheet = row.getCell(2).value.toString();

                    if (nimFromSheet === nim) {
                        student = {
                            no: row.getCell(1).value,
                            nim: row.getCell(2).value,
                            nama: row.getCell(3).value,
                            programStudi: row.getCell(4).value,
                            fakultas: row.getCell(5).value,
                            ukuranAlmamater: row.getCell(10).value,
                            nomorUrut: row.getCell(12).value,
                            statusTagihanWisuda: row.getCell(11).value
                        };
                        return false;
                    }
                }
            });

            if (student) {
                const pdfPath = `BUKTI_WISUDA_${student.nim}.pdf`;

                const doc = new pdfkit({ size: 'A4', margin: { right: 10 } });
                const buffers = [];
                doc.on('data', buffers.push.bind(buffers));
                doc.on('end', () => {
                    const pdfData = Buffer.concat(buffers);

                    // Kirim file PDF sebagai respons ke client
                    // res.set({
                    //     'Content-Type': 'application/pdf',
                    //     'Content-Disposition': `attachment; filename=${pdfPath}`,
                    //     'Content-Length': pdfData.length
                    // });
                    // res.send(pdfData);

                    // Simpan file PDF ke sistem file
                    const pdfPathOnServer = `pdf_output/${pdfPath}`;
                    fs.writeFile(pdfPathOnServer, pdfData, (err) => {
                        if (err) {
                            console.error('Error saving PDF file:', err);
                            res.status(500).send('Error saving PDF file');
                        } else {
                            const infoLog = `${getCurrentTimestamp()} - Download Success for NIM: ${student.nim}\n`;
                            fs.appendFileSync(`logDownloadSuccess.log`, infoLog);
                            // Kirim file PDF sebagai respons ke client
                            res.set({
                                'Content-Type': 'application/pdf',
                                'Content-Disposition': `attachment; filename=${pdfPath}`,
                                'Content-Length': pdfData.length
                            });
                            res.send(pdfData);
                        }
                    });
                });

                // Mendapatkan ukuran halaman PDF
                const pageWidth = doc.page.width;
                // Mendapatkan ukuran gambar
                const imageWidth = 539;

                // Menghitung koordinat untuk menempatkan gambar di tengah halaman
                const x = (pageWidth - imageWidth) / 2;

                // Menambahkan header dan footer
                doc.image('header.PNG', x, 14, { width: imageWidth });
                doc.image("footer.PNG", x, 771, { width: 539 });

                // Tambahkan konten PDF
                const text1 = "PESERTA WISUDA SARJANA (S1) DAN PASCASARJANA (S2 & S3)";
                const text2 = "UNIVERSITAS PASUNDAN GELOMBANG II TAHUN AKADEMIK 2023/2024";

                const text3 = "Sekretariat: Jl. Tamansari No. 4-8 Bandung, Call Center: 0811960193, Email: rektorat@unpas.ac.id";
                const text4 = "Email: rektorat@unpas.ac.id Website: www.unpas.ac.id";

                const text5 = "Selamat Anda telah terdaftar sebagai Peserta Wisuda Universitas Pasundan Gelombang II";
                const text6 = "Tahun Akademik 2023/2024, dengan data sebagai berikut:";

                const text7 = "DATA WISUDAWAN/WISUDAWATI";

                const text8 = "Surat Keterangan ini bisa digunakan sebagai bukti untuk pengambilan perlengkapan Peserta";
                const text9 = "Wisuda Universitas Pasundan Gelombang II Tahun Akademik 2023/2024.";

                const text10 = "Ceklis pengambilan perlengkapan wisuda:";
                const text11 = "[  ] Toga";
                const text12 = "[  ] Pin";
                const text13 = "[  ] Undangan Wisuda";

                const lokasiWisuda = "Sasana Budaya Ganesha (SABUGA)";
                const waktuGladi = "Jumat, 17 Mei 2024, 13.30 WIB s.d. Selesai";
                const waktuPelaksanaan = "Sabtu, 18 Mei 2024, 07.00 WIB s.d. Selesai";

                doc.registerFont('Arial Font', 'fonts/arial.ttf');
                doc.registerFont('Arial Bold Font', 'fonts/arial-bold.ttf');
                doc.font('Arial Bold Font')
                    .fontSize(12).text(text1, 96, 135, { align: 'center' });

                doc.fontSize(12).text(text2, 96, 152, { align: 'center' });

                doc.font('Arial Font')
                    .fontSize(8).text(text3, 96, 173, { align: 'center' });

                doc.fontSize(8).text(text4, 96, 183, { align: 'center' });

                doc.fontSize(11).text(text5, 74, 223, { align: 'left' });
                doc.fontSize(11).text(text6, 74, 240, { align: 'left' });

                doc.font('Arial Bold Font')
                    .fontSize(12).text(text7, 96, 268, { align: 'center' });

                doc.font('Arial Font')
                    .fontSize(11).text("NIM", 74, 296, { align: 'left' });

                doc.fontSize(11).text(`: ${student.nim}`, 215, 296, { align: 'left' });

                doc.fontSize(11).text("Nama", 74, 316, { align: 'left' });
                doc.fontSize(11).text(`: ${student.nama}`, 215, 316, { align: 'left' });

                doc.fontSize(11).text("Program Studi", 74, 336, { align: 'left' });
                doc.fontSize(11).text(`: ${student.programStudi}`, 215, 336, { align: 'left' });

                doc.fontSize(11).text("Fakultas", 74, 356, { align: 'left' });
                doc.fontSize(11).text(`: ${student.fakultas}`, 215, 356, { align: 'left' });

                doc.fontSize(11).text("Ukuran Toga", 74, 376, { align: 'left' });
                doc.fontSize(11).text(`: ${student.ukuranAlmamater}`, 215, 376, { align: 'left' });

                doc.fontSize(11).text("Nomor Urut/Kursi", 74, 396, { align: 'left' });
                doc.fontSize(11).text(`: ${student.nomorUrut}`, 215, 396, { align: 'left' });

                doc.fontSize(11).text("Lokasi Wisuda", 74, 416, { align: 'left' });
                doc.fontSize(11).text(`: ${lokasiWisuda}`, 215, 416, { align: 'left' });

                doc.fontSize(11).text("Waktu Gladi Resik", 74, 436, { align: 'left' });
                doc.fontSize(11).text(`: ${waktuGladi}`, 215, 436, { align: 'left' });

                doc.fontSize(11).text("Waktu Pelaksanaan", 74, 456, { align: 'left' });
                doc.fontSize(11).text(`: ${waktuPelaksanaan}`, 215, 456, { align: 'left' });

                doc.fontSize(11).text("Status Tagihan Wisuda", 74, 476, { align: 'left' });
                doc.fontSize(11).text(`: ${student.statusTagihanWisuda}`, 215, 476, { align: 'left' });

                doc.fontSize(11).text(text8, 74, 510, { align: 'left', lineBreak: false });
                doc.fontSize(11).text(text9, 74, 527, { align: 'left' });

                doc.fontSize(11).text(text10, 74, 554, { align: 'left' });
                doc.fontSize(11).text(text11, 74, 575, { align: 'left' });
                doc.fontSize(11).text(text12, 74, 595, { align: 'left' });
                doc.fontSize(11).text(text13, 74, 612, { align: 'left' });

                doc.fontSize(11).text(`Bandung, ${today}`, 426, 688, { align: 'left', lineBreak: false });
                doc.fontSize(11).text('Panitia', 447, 757, { align: 'left' });

                doc.end();
            } else {
                res.status(404).send('NIM tidak terdaftar');
            }
        })
        .catch(err => {
            console.error(err);
            res.status(500).send('Error reading Excel file');
        });
});


const PORT = process.env.PORT || 8001;
app.listen(PORT, () => console.log(`Server started on port ${PORT}`));
