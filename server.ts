import express from "express";
import path from "path";
import dotenv from "dotenv";
import { GoogleGenAI, Type } from "@google/genai";
import { createServer as createViteServer } from "vite";
import { Packer } from "docx";
import {
  generateCPDoc,
  generateATPDoc,
  generateProtaPromesDoc,
  generateKKTPDoc,
  generateRPMDoc,
  generateTopikDoc,
  generateLKPDDoc,
  generateAsesmenDoc,
} from "./src/utils/docxGenerator.js";

dotenv.config();

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;

app.use(express.json({ limit: "50mb" }));

// Support CORS for client-server split deployment
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, Authorization");
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  if (req.method === "OPTIONS") {
    res.sendStatus(200);
    return;
  }
  next();
});

// Initialize Google GenAI on server
const apiKey = process.env.GEMINI_API_KEY;
if (!apiKey) {
  console.warn("WARNING: GEMINI_API_KEY is not defined in environment variables. Gemini calls will fail.");
}

const ai = new GoogleGenAI({
  apiKey: apiKey || "MOCK_KEY_IF_MISSING_TO_PREVENT_LOAD_ERROR",
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

// Helper to check for API Key and handle errors
function ensureApiKey(res: express.Response) {
  if (!process.env.GEMINI_API_KEY) {
    res.status(500).json({
      error: "GEMINI_API_KEY is missing. Please configure it in the Secrets panel in the AI Studio UI under Settings > Secrets.",
    });
    return false;
  }
  return true;
}

// ==========================================
// API ENDPOINTS FOR DOCUMENT GENERATION
// ==========================================

// 1. GENERATE CAPAIAN PEMBELAJARAN (CP)
app.post("/api/generate/cp", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const {
    namaGuru,
    nipGuru,
    namaKepalaSekolah,
    nipKepalaSekolah,
    mataPelajaran,
    jenjangSekolah,
    fase,
    tahunAjaran,
  } = req.body;

  // Determine which regulation based on subject (religion subjects follow BKPDM No 020 2026, others follow BSKAP No 046 2025)
  const isReligious = /agama/i.test(mataPelajaran || "") || 
                      /islam/i.test(mataPelajaran || "") || 
                      /kristen/i.test(mataPelajaran || "") || 
                      /katolik/i.test(mataPelajaran || "") || 
                      /hindu/i.test(mataPelajaran || "") || 
                      /buddha/i.test(mataPelajaran || "") || 
                      /konghucu/i.test(mataPelajaran || "") || 
                      /budi pekerti/i.test(mataPelajaran || "");

  const regulasiUtama = isReligious
    ? "Keputusan Kepala BKPDM Nomor 020 Tahun 2026 tentang Capaian Pembelajaran Pendidikan Agama dan Budi Pekerti"
    : "Keputusan Kepala BSKAP Nomor 046/H/KR/2025 tentang Capaian Pembelajaran pada Pendidikan Anak Usia Dini, Jenjang Pendidikan Dasar, dan Jenjang Pendidikan Menengah pada Kurikulum Merdeka";

  const prompt = `
    Anda adalah asisten kurikulum ahli di Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi Indonesia.
    Tugas Anda adalah membuat dokumen Capaian Pembelajaran (CP) yang komprehensif, profesional, dan realistis berdasarkan data guru dan sekolah berikut:
    
    - Nama Mata Pelajaran: ${mataPelajaran}
    - Jenjang Sekolah: ${jenjangSekolah}
    - Fase: ${fase}
    - Tahun Ajaran: ${tahunAjaran}
    
    PERATURAN WAJIB YANG HARUS DIIKUTI:
    Dokumen CP harus mengikuti regulasi: "${regulasiUtama}".
    
    Buatlah dokumen CP dengan detail yang tinggi, menggunakan peristilahan resmi Kurikulum Merdeka, terdiri dari:
    1. Rasionalisasi Mata Pelajaran (mengapa mata pelajaran ini diajarkan).
    2. Tujuan Mata Pelajaran (apa saja kompetensi yang dituju).
    3. Karakteristik Mata Pelajaran (karakter materi, cara belajar, atau pengorganisasian pembelajaran).
    4. Capaian Pembelajaran Umum untuk Fase ${fase.toUpperCase()}.
    5. Pembagian Elemen Capaian Pembelajaran. Tentukan elemen-elemen yang relevan dengan mata pelajaran ${mataPelajaran} (Minimal 3 elemen, contoh: untuk PPKn/Pancasila ada Pancasila, UUD 1945, NKRI, Bhinneka Tunggal Ika. Untuk Matematika ada Bilangan, Aljabar, Pengukuran, Geometri, Analisis Data. Untuk PAI ada Al-Qur'an Hadis, Akidah, Akhlak, Fikih, Sejarah Peradaban Islam). Untuk setiap elemen, berikan Nama Elemen, Deskripsi singkat elemen, dan rumusan paragraf Capaian Pembelajaran Elemen untuk Fase ${fase.toUpperCase()} tersebut secara lengkap dan deskriptif.

    Format keluaran harus merupakan JSON murni yang sesuai dengan skema JSON berikut:
    {
      "judul": "Capaian Pembelajaran (CP) Mata Pelajaran ${mataPelajaran}",
      "regulasi": "${regulasiUtama}",
      "rasional": "Penjelasan rasionalisasi...",
      "tujuan": "Penjelasan tujuan...",
      "karakteristik": "Penjelasan karakteristik...",
      "capaianUmum": "Rumusan capaian umum Fase ${fase.toUpperCase()}...",
      "elemen": [
        {
          "nama": "Nama Elemen 1",
          "deskripsi": "Deskripsi singkat elemen 1...",
          "capaian": "Rumusan Capaian Pembelajaran Elemen 1..."
        }
      ]
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            judul: { type: Type.STRING },
            regulasi: { type: Type.STRING },
            rasional: { type: Type.STRING },
            tujuan: { type: Type.STRING },
            karakteristik: { type: Type.STRING },
            capaianUmum: { type: Type.STRING },
            elemen: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  nama: { type: Type.STRING },
                  deskripsi: { type: Type.STRING },
                  capaian: { type: Type.STRING },
                },
                required: ["nama", "deskripsi", "capaian"],
              },
            },
          },
          required: ["judul", "regulasi", "rasional", "tujuan", "karakteristik", "capaianUmum", "elemen"],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating CP:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Capaian Pembelajaran." });
  }
});

// 2. GENERATE ANALISIS CP & ATP
app.post("/api/generate/atp", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, cpData } = req.body;

  const prompt = `
    Anda adalah ahli kurikulum sekolah di Indonesia. 
    Berdasarkan Capaian Pembelajaran (CP) yang telah dianalisis sebelumnya, buatlah 'Analisis Capaian Pembelajaran (CP) dan Alur Tujuan Pembelajaran (ATP)' untuk Mata Pelajaran ${meta.mataPelajaran} Fase ${meta.fase.toUpperCase()} Tahun Ajaran ${meta.tahunAjaran}.
    
    DATA CAPAIAN PEMBELAJARAN (CP):
    Rasional: ${cpData.rasional}
    Capaian Umum: ${cpData.capaianUmum}
    Elemen CP yang digunakan:
    ${cpData.elemen.map((el: any) => `- Elemen: ${el.nama}\n  Capaian Elemen: ${el.capaian}`).join("\n")}
    
    TUGAS ANDA:
    1. Analisis setiap elemen CP tersebut dan rumuskan minimal 2-3 Tujuan Pembelajaran (TP) konkret per elemen yang mengukur kompetensi esensial.
    2. Susunlah TP tersebut menjadi satu Alur Tujuan Pembelajaran (ATP) yang logis dan berurutan secara ketat mulai dari tujuan pembelajaran yang paling penting, fundamental, atau esensial ke tujuan pembelajaran yang lebih kompleks atau lanjutan (essential-first sequence). Berikan kode unik untuk setiap TP (contoh: TP ${meta.fase}.1.1, TP ${meta.fase}.1.2, TP ${meta.fase}.2.1, dst).
    3. Untuk setiap TP, tentukan:
       a. Topik atau Materi Pokok Utama yang sesuai.
       b. Alokasi Waktu (JP) yang realistis (misalnya 4 JP, 6 JP, atau 8 JP).
       c. Dimensi Profil Lulusan yang ditumbuhkan sesuai dengan 8 Dimensi Profil Lulusan Kurikulum Merdeka (Pilih 1-2 dimensi paling relevan dari daftar berikut:
          - Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia
          - Bernalar Kritis dan Solutif
          - Kreatif dan Inovatif
          - Bergotong Royong dan Kolaboratif
          - Mandiri dan Bertanggung Jawab
          - Berkebinekaan Global dan Inklusif
          - Literat dan Berwawasan Teknologi
          - Berkarakter, Sehat Jasmani dan Rohani).
       d. Glosarium / penjelasan singkat istilah penting.

    Format keluaran harus merupakan JSON murni sesuai skema berikut:
    {
      "judul": "Analisis CP dan Alur Tujuan Pembelajaran (ATP) ${meta.mataPelajaran}",
      "fase": "${meta.fase.toUpperCase()}",
      "mataPelajaran": "${meta.mataPelajaran}",
      "atp": [
        {
          "elemen": "Nama Elemen CP",
          "capaian": "Rumusan Capaian Pembelajaran lengkap untuk elemen ini...",
          "tujuanPembelajaran": [
            {
              "kode": "TP Kode",
              "tp": "Rumusan Tujuan Pembelajaran...",
              "topik": "Nama Topik / Materi Utama...",
              "alokasiWaktu": "4 JP",
              "profilLulusan": "Dimensi Profil Lulusan...",
              "glosarium": "Penjelasan istilah..."
            }
          ]
        }
      ]
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            judul: { type: Type.STRING },
            fase: { type: Type.STRING },
            mataPelajaran: { type: Type.STRING },
            atp: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  elemen: { type: Type.STRING },
                  tujuanPembelajaran: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        kode: { type: Type.STRING },
                        tp: { type: Type.STRING },
                        topik: { type: Type.STRING },
                        alokasiWaktu: { type: Type.STRING },
                        profilLulusan: { type: Type.STRING },
                        glosarium: { type: Type.STRING },
                      },
                      required: ["kode", "tp", "topik", "alokasiWaktu", "profilLulusan", "glosarium"],
                    },
                  },
                },
                required: ["elemen", "tujuanPembelajaran"],
              },
            },
          },
          required: ["judul", "fase", "mataPelajaran", "atp"],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating ATP:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Analisis CP & ATP." });
  }
});

// 3. GENERATE PROTA & PROMES
app.post("/api/generate/prota-promes", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, atpData } = req.body;

  // Flatten the ATP data to extract all topics
  const tpList: any[] = [];
  atpData.atp.forEach((item: any) => {
    item.tujuanPembelajaran.forEach((tpObj: any) => {
      tpList.push({
        elemen: item.elemen,
        kode: tpObj.kode,
        tp: tpObj.tp,
        topik: tpObj.topik,
        alokasiWaktu: tpObj.alokasiWaktu,
      });
    });
  });

  const prompt = `
    Anda adalah perancang jadwal akademik sekolah di Indonesia sesuai Kurikulum Merdeka.
    Berdasarkan data Alur Tujuan Pembelajaran (ATP) berikut, buatlah dokumen 'Program Tahunan (PROTA)' dan 'Program Semester (PROMES)' untuk mata pelajaran ${meta.mataPelajaran} Fase ${meta.fase.toUpperCase()} Tahun Ajaran ${meta.tahunAjaran}.

    DAFTAR ATP DAN TOPIK:
    ${tpList.map((t) => `- [${t.kode}] Elemen: ${t.elemen}, Topik: ${t.topik} (${t.alokasiWaktu}) - TP: ${t.tp}`).join("\n")}

    TUGAS ANDA:
    1. Buat pemetaan Program Tahunan (PROTA):
       Petakan semua TP di atas secara berurutan. Kolom yang dihasilkan adalah No, Elemen / Lingkup Materi, Tujuan Pembelajaran, Alokasi Waktu (JP).
       Alokasi waktu total tahunan adalah 360 JP (atau total dari seluruh TP, pastikan totalnya konsisten).
    2. Buat pemetaan Program Semester (PROMES):
       Petakan seluruh Tujuan Pembelajaran secara berurutan dari nomor 1 sampai selesai (misal 10 TP).
       Tentukan alokasi JP masing-masing TP (misal 24, 36, 40, 44, dst).
       Distribusikan alokasi JP ke 12 bulan (Jul, Agt, Sep, Okt, Nov, Des, Jan, Feb, Mar, Apr, Mei, Jun) secara realistis sesuai pembagian Semester 1 (Juli-Desember) dan Semester 2 (Januari-Juni).
       Aturan Distribusi Bulanan:
       - Setiap TP hanya diajarkan di bulan tertentu yang sesuai dengan alokasi JP-nya. Misal, TP 1 diajarkan di bulan Jul sebesar 24 JP. Maka di kolom "Jul" bernilai "24", dan kolom bulan lainnya bernilai "-".
       - TP 2 diajarkan di bulan Agt sebesar 36 JP. Maka di kolom "Agt" bernilai "36", dan kolom lainnya bernilai "-".
       - Distribusikan secara berurutan sehingga Semester Ganjil (Jul s.d. Nov) terisi untuk TP awal, dan Semester Genap (Jan s.d. Mei) terisi untuk TP akhir.
       - Untuk kolom "Des" (Desember), karena merupakan masa ujian akhir semester ganjil, isi dengan "ASG" (Asesmen Sumatif Ganjil / Cadangan) untuk baris-baris TP pertengahan hingga akhir (misalnya baris 5 s.d. 10).
       - Untuk kolom "Jun" (Juni), isi dengan "ASE" (Asesmen Sumatif Akhir Tahun) di baris terakhir (TP terakhir), dan "-" untuk baris-baris lainnya.
       - Bulan-bulan lain yang tidak ada jam pembelajaran diisi dengan "-".
       - Hitung jumlah JP untuk Semester Ganjil (total JP untuk TP yang diajarkan di Juli-Desember, misal 172 JP) dan Semester Genap (total JP untuk TP yang diajarkan di Januari-Juni, misal 188 JP).

    Format keluaran harus berupa JSON murni sesuai skema berikut:
    {
      "judulProta": "PROGRAM TAHUNAN (PROTA)",
      "judulPromes": "PROGRAM SEMESTER (PROSEM)",
      "prota": [
        {
          "no": 1,
          "elemen": "Perencanaan dan Pengalamatan Jaringan",
          "tujuanPembelajaran": "Menganalisis kebutuhan teknis pengguna...",
          "alokasiWaktu": "24 JP"
        }
      ],
      "promes": [
        {
          "no": 1,
          "tujuanPembelajaran": "Menganalisis kebutuhan teknis pengguna...",
          "alokasiWaktu": 24,
          "distribusi": {
            "Jul": "24",
            "Agt": "-",
            "Sep": "-",
            "Okt": "-",
            "Nov": "-",
            "Des": "-",
            "Jan": "-",
            "Feb": "-",
            "Mar": "-",
            "Apr": "-",
            "Mei": "-",
            "Jun": "-"
          }
        }
      ],
      "jumlahJpSemester1": 172,
      "jumlahJpSemester2": 188
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            judulProta: { type: Type.STRING },
            judulPromes: { type: Type.STRING },
            prota: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  no: { type: Type.INTEGER },
                  elemen: { type: Type.STRING },
                  tujuanPembelajaran: { type: Type.STRING },
                  alokasiWaktu: { type: Type.STRING },
                },
                required: ["no", "elemen", "tujuanPembelajaran", "alokasiWaktu"],
              },
            },
            promes: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  no: { type: Type.INTEGER },
                  tujuanPembelajaran: { type: Type.STRING },
                  alokasiWaktu: { type: Type.INTEGER },
                  distribusi: {
                    type: Type.OBJECT,
                    properties: {
                      Jul: { type: Type.STRING },
                      Agt: { type: Type.STRING },
                      Sep: { type: Type.STRING },
                      Okt: { type: Type.STRING },
                      Nov: { type: Type.STRING },
                      Des: { type: Type.STRING },
                      Jan: { type: Type.STRING },
                      Feb: { type: Type.STRING },
                      Mar: { type: Type.STRING },
                      Apr: { type: Type.STRING },
                      Mei: { type: Type.STRING },
                      Jun: { type: Type.STRING },
                    },
                    required: ["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"],
                  },
                },
                required: ["no", "tujuanPembelajaran", "alokasiWaktu", "distribusi"],
              },
            },
            jumlahJpSemester1: { type: Type.INTEGER },
            jumlahJpSemester2: { type: Type.INTEGER },
          },
          required: ["judulProta", "judulPromes", "prota", "promes", "jumlahJpSemester1", "jumlahJpSemester2"],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating Prota/Promes:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate PROTA & PROMES." });
  }
});

// 4. GENERATE KKTP
app.post("/api/generate/kktp", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, atpData } = req.body;

  const tpList: any[] = [];
  atpData.atp.forEach((item: any) => {
    item.tujuanPembelajaran.forEach((tpObj: any) => {
      tpList.push({
        kode: tpObj.kode,
        tp: tpObj.tp,
        topik: tpObj.topik,
      });
    });
  });

  const prompt = `
    Anda adalah ahli asesmen pendidikan Kurikulum Merdeka di Indonesia.
    Berdasarkan data Tujuan Pembelajaran (TP) berikut, buatlah dokumen 'Kriteria Ketuntasan Tujuan Pembelajaran (KKTP)' untuk setiap TP pada mata pelajaran ${meta.mataPelajaran} Fase ${meta.fase.toUpperCase()}.

    DAFTAR TUJUAN PEMBELAJARAN (TP):
    ${tpList.map((t) => `- [${t.kode}] ${t.tp}`).join("\n")}

    TUGAS ANDA:
    1. Untuk setiap TP, susun rubrik Kriteria Ketuntasan Tujuan Pembelajaran yang terdiri dari minimal 2 aspek penilaian esensial (misalnya: "Aspek Pemahaman Konsep" dan "Aspek Penerapan Praktik/Keterampilan").
    2. Di setiap aspek, rumuskan deskripsi kriteria pencapaian untuk 4 interval tingkat kompetensi:
       - Mulai Berkembang (0-60)
       - Layak (61-70)
       - Cakap (71-80)
       - Mahir (81-100)
    3. Tentukan interval ketuntasan umum beserta keterangannya (misal: "Kurang (0-60): Belum Tuntas, Remedial", "Cukup (61-70): Tuntas (Syarat), Pendampingan", "Baik (71-80): Tuntas, Lanjut materi berikutnya", "Sangat Baik (81-100): Tuntas (Sangat Baik), Pengayaan").

    Format keluaran harus merupakan JSON murni sesuai skema berikut:
    {
      "judul": "ANALISIS KRITERIA KETERCAPAIAN (KKTP) - ${meta.mataPelajaran}",
      "kktp": [
        {
          "tpKode": "TP Kode",
          "tpTeks": "Rumusan Tujuan Pembelajaran...",
          "kriteria": [
            {
              "aspek": "Aspek Penilaian (misal: Indikator Ketercapaian / IKTP)",
              "baruBerkembang": "Deskripsi mulai berkembang...",
              "layak": "Deskripsi layak...",
              "cakap": "Deskripsi cakap...",
              "mahir": "Deskripsi mahir..."
            }
          ],
          "intervalNilai": {
            "kurang": "0-60 (Mulai Berkembang - Belum Tuntas)",
            "cukup": "61-70 (Layak - Tuntas Bersyarat)",
            "baik": "71-80 (Cakap - Tuntas)",
            "sangatBaik": "81-100 (Mahir - Tuntas Sangat Baik)"
          }
        }
      ]
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            judul: { type: Type.STRING },
            kktp: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  tpKode: { type: Type.STRING },
                  tpTeks: { type: Type.STRING },
                  kriteria: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        aspek: { type: Type.STRING },
                        baruBerkembang: { type: Type.STRING },
                        layak: { type: Type.STRING },
                        cakap: { type: Type.STRING },
                        mahir: { type: Type.STRING },
                      },
                      required: ["aspek", "baruBerkembang", "layak", "cakap", "mahir"],
                    },
                  },
                  intervalNilai: {
                    type: Type.OBJECT,
                    properties: {
                      kurang: { type: Type.STRING },
                      cukup: { type: Type.STRING },
                      baik: { type: Type.STRING },
                      sangatBaik: { type: Type.STRING },
                    },
                    required: ["kurang", "cukup", "baik", "sangatBaik"],
                  },
                },
                required: ["tpKode", "tpTeks", "kriteria", "intervalNilai"],
              },
            },
          },
          required: ["judul", "kktp"],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating KKTP:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate KKTP." });
  }
});

// 5. GENERATE DAFTAR TOPIK PEMBELAJARAN
app.post("/api/generate/topik", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, atpData } = req.body;

  // Flatten the ATP data to extract all topics
  const tpList: any[] = [];
  atpData.atp.forEach((item: any) => {
    item.tujuanPembelajaran.forEach((tpObj: any) => {
      tpList.push({
        elemen: item.elemen,
        kode: tpObj.kode,
        tp: tpObj.tp,
        topik: tpObj.topik,
        alokasiWaktu: tpObj.alokasiWaktu,
      });
    });
  });

  const prompt = `
    Anda adalah ahli kurikulum sekolah di Indonesia.
    Berdasarkan data Alur Tujuan Pembelajaran (ATP) berikut, susunlah 'Daftar Topik Pembelajaran' yang sistematis dan berurutan dari tujuan pembelajaran yang paling penting/esensial ke yang lebih kompleks/lanjutan.
    
    DATA ATP DAN TP:
    ${tpList.map((t) => `- [${t.kode}] Topik: ${t.topik} (${t.alokasiWaktu}) - TP: ${t.tp}`).join("\n")}

    TUGAS ANDA:
    1. Identifikasi topik/materi utama dari setiap Tujuan Pembelajaran (TP) tersebut.
    2. Susun secara berurutan berdasarkan tingkat kepentingan/keesensialan materi (materi prasyarat/esensial di awal).
    3. Untuk setiap topik utama, berikan deskripsi ringkas cakupan pembelajarannya, kode TP terkait, rumusan TP terkait, alokasi waktu (JP), dan tentukan perkiraan semester (Semester 1 atau 2) secara proporsional.

    Format keluaran harus merupakan JSON murni sesuai skema berikut:
    {
      "judul": "Daftar Topik Pembelajaran ${meta.mataPelajaran}",
      "daftarTopik": [
        {
          "no": 1,
          "topik": "Nama Topik / Materi Utama",
          "deskripsi": "Deskripsi ringkas cakupan materi yang diajarkan pada topik ini...",
          "tpKode": "TP Kode",
          "tpTeks": "Rumusan TP...",
          "alokasiWaktu": "4 JP",
          "semester": 1
        }
      ]
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            judul: { type: Type.STRING },
            daftarTopik: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  no: { type: Type.INTEGER },
                  topik: { type: Type.STRING },
                  deskripsi: { type: Type.STRING },
                  tpKode: { type: Type.STRING },
                  tpTeks: { type: Type.STRING },
                  alokasiWaktu: { type: Type.STRING },
                  semester: { type: Type.INTEGER },
                },
                required: ["no", "topik", "deskripsi", "tpKode", "tpTeks", "alokasiWaktu", "semester"],
              },
            },
          },
          required: ["judul", "daftarTopik"],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating Topik:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Daftar Topik." });
  }
});

// 6. GENERATE RENCANA PEMBELAJARAN MENDALAM (RPM) SINGLE
app.post("/api/generate/rpm-single", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, topikItem } = req.body;

  const prompt = `
    Anda adalah ahli pedagogi luar biasa dan penyusun rencana pelaksanaan pembelajaran inovatif (Rencana Pembelajaran Mendalam / Modul Ajar Pembelajaran Mendalam) untuk Kurikulum Merdeka di Indonesia.
    Tugas Anda adalah membuat dokumen Rencana Pembelajaran Mendalam (RPM) / Modul Ajar yang sangat mendetail, profesional, dan berpusat pada murid berdasarkan topik pembelajaran berikut:

    Mata Pelajaran: ${meta.mataPelajaran}
    Fase: ${meta.fase.toUpperCase()}
    Materi Pokok (Topik): ${topikItem.topik}
    Tujuan Pembelajaran Utama: [${topikItem.tpKode}] ${topikItem.tpTeks}
    Alokasi Waktu: ${topikItem.alokasiWaktu}
    Semester: Semester ${topikItem.semester}

    Dokumen ini harus mengikuti struktur Modul Ajar Pembelajaran Mendalam dengan komponen-komponen berikut secara lengkap dan mendalam:

    1. IDENTITAS MODUL
       - namaGuru: ${meta.namaGuru}
       - sekolah: ${meta.jenjangSekolah}
       - mataPelajaran: ${meta.mataPelajaran}
       - kelasSemester: Kelas ${meta.fase === 'f' ? 'XI/XII' : meta.fase === 'e' ? 'X' : 'VII/VIII/IX'} / Semester ${topikItem.semester === 1 ? 'Ganjil' : 'Genap'}
       - semesterTahun: ${topikItem.semester === 1 ? 'Ganjil' : 'Genap'} - ${meta.tahunAjaran}
       - topik: ${topikItem.topik}
       - modelPembelajaran: e.g. "Tatap Muka / Reguler", "Project-Based Learning (PjBL)", "Discovery Learning", dll. yang relevan dengan topik ini.

    2. A. IDENTIFIKASI
       - targetPesertaDidik: Deskripsi lengkap target siswa, misal: "Peserta didik reguler/tipikal Kelas ... yang memiliki kesiapan belajar tingkat menengah serta antusiasme tinggi..."
       - kompetensiPrasyarat: Minimal 2 kemampuan awal/dasar yang harus dimiliki siswa sebelum mempelajari materi ini.
       - profilLulusan: Minimal 3 item yang dipilih dan dirumuskan dari 8 Dimensi Profil Lulusan berikut (format "NAMA_DIMENSI: Penjelasan konkret", contoh: "Kompeten: Memiliki kecakapan teknis...", "Literat: Mampu melakukan bernalar kritis...", "Berkarakter: Mengembangkan kemandirian..."):
         * Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia
         * Bernalar Kritis dan Solutif
         * Kreatif dan Inovatif
         * Bergotong Royong dan Kolaboratif
         * Mandiri dan Bertanggung Jawab
         * Berkebinekaan Global dan Inklusif
         * Literat dan Berwawasan Teknologi
         * Berkarakter, Sehat Jasmani dan Rohani
       - saranaPrasarana:
         * fasilitasFisik: Laboratorium komputer, proyektor, dll. yang spesifik untuk topik ini.
         * pirantiLunak: Simulator, software, platform, atau tools digital spesifik untuk topik ini.
         * sumberBelajar: Lembar Kerja Peserta Didik (LKPD), modul digital, video demonstrasi, dll.

    3. B. DESAIN PEMBELAJARAN
       - capaianPembelajaran: Rumusan Capaian Pembelajaran yang relevan (misal merujuk ke SK Kepala BSKAP No. 046/H/KR/2025).
       - lintasDisiplin: Relevansi dengan disiplin ilmu lain (misal: "Matematika (Numerasi): ...", "Bahasa Inggris: ...", atau Sains/Sosial/Seni dll.).
       - tujuanPembelajaran: Minimal 2 tujuan pembelajaran konkret dengan format aktivitas active learning (misal: "Melalui diskusi mendalam...", "Melalui praktikum mandiri...").
       - topikPembelajaran: Cakupan/sub-materi yang dipelajari.
       - praktikPedagogis: Pendekatan pembelajaran mendalam (Deep Learning), model yang digunakan (misal: Project-Based Learning, Guided Inquiry), dan fokus pembelajaran.
       - kemitraanPembelajaran: Hubungan materi dengan skenario dunia nyata atau mitra industri (DUDI) atau aplikasi praktis lapangan.
       - lingkunganBelajar: Pengaturan kelas/laboratorium agar kondusif, inklusif, dan kolaboratif.
       - pemanfaatanDigital: Integrasi teknologi untuk pembelajaran kolaboratif/virtual.

    4. C. LANGKAH PEMBELAJARAN
       Harus terdiri dari 4 Fase Kegiatan yang berpusat pada murid dengan menerapkan pendekatan Deep Learning (Mindful, Meaningful, Joyful):
       - pendahuluan: Durasi (e.g. "15 Menit"), Pendekatan ("MINDFUL"), Aktivitas Guru (minimal 3 poin), Aktivitas Siswa (minimal 2 poin). Fokus pada kesadaran penuh, hening sejenak (S-T-O-P), apersepsi, dan pemantik rasa ingin tahu.
       - intiEksplorasi: Durasi (e.g. "45 Menit"), Pendekatan ("MEANINGFUL"), Aktivitas Guru (minimal 3 poin), Aktivitas Siswa (minimal 3 poin). Fokus pada eksplorasi konsep, analisis studi kasus nyata, dan diskusi kelompok secara mendalam.
       - intiRancangan: Durasi (e.g. "45 Menit"), Pendekatan ("JOYFUL"), Aktivitas Guru (minimal 3 poin), Aktivitas Siswa (minimal 2 poin). Fokus pada perancangan, praktikum, simulasi, tantangan proyek, dan eksperimen yang menyenangkan.
       - penutup: Durasi (e.g. "15 Menit"), Pendekatan ("MINDFUL & MEANINGFUL"), Aktivitas Guru (minimal 3 poin), Aktivitas Siswa (minimal 2 poin). Fokus pada refleksi pribadi, presentasi karya, kesimpulan bersama, dan penguatan konsep.

    5. D. ASESMEN
       Harus mencakup tiga jenis penilaian dengan masing-masing kolom Metode Penilaian, Instrumen Penilaian, dan Kriteria Keberhasilan:
       - diagnostik (Awal)
       - formatif (Proses)
       - sumatif (Akhir)

    6. E. REMEDIAL & PENGAYAAN
       - remedial: Sasaran peserta didik, dan minimal 2 bentuk Aktivitas Intervensi terstruktur yang disederhanakan.
       - pengayaan: Sasaran peserta didik berkemampuan tinggi, dan minimal 2 bentuk tantangan level tinggi (HOTS).

    7. F. REFLEKSI
       - siswa: Minimal 3 pertanyaan reflektif pemicu kesadaran belajar.
       - guru: Minimal 3 pertanyaan refleksi diri pendidik.

    Format keluaran harus berupa JSON murni tanpa markdown lain di luar objek JSON tersebut. Pastikan semua field terisi lengkap dan mendalam.
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            judul: { type: Type.STRING },
            identitas: {
              type: Type.OBJECT,
              properties: {
                namaGuru: { type: Type.STRING },
                sekolah: { type: Type.STRING },
                mataPelajaran: { type: Type.STRING },
                kelasSemester: { type: Type.STRING },
                semesterTahun: { type: Type.STRING },
                topik: { type: Type.STRING },
                modelPembelajaran: { type: Type.STRING },
              },
              required: ["namaGuru", "sekolah", "mataPelajaran", "kelasSemester", "semesterTahun", "topik", "modelPembelajaran"],
            },
            identifikasi: {
              type: Type.OBJECT,
              properties: {
                targetPesertaDidik: { type: Type.STRING },
                kompetensiPrasyarat: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING }
                },
                profilLulusan: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING }
                },
                saranaPrasarana: {
                  type: Type.OBJECT,
                  properties: {
                    fasilitasFisik: { type: Type.STRING },
                    pirantiLunak: { type: Type.STRING },
                    sumberBelajar: { type: Type.STRING },
                  },
                  required: ["fasilitasFisik", "pirantiLunak", "sumberBelajar"],
                }
              },
              required: ["targetPesertaDidik", "kompetensiPrasyarat", "profilLulusan", "saranaPrasarana"],
            },
            desainPembelajaran: {
              type: Type.OBJECT,
              properties: {
                capaianPembelajaran: { type: Type.STRING },
                lintasDisiplin: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING }
                },
                tujuanPembelajaran: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING }
                },
                topikPembelajaran: { type: Type.STRING },
                praktikPedagogis: { type: Type.STRING },
                kemitraanPembelajaran: { type: Type.STRING },
                lingkunganBelajar: { type: Type.STRING },
                pemanfaatanDigital: { type: Type.STRING },
              },
              required: [
                "capaianPembelajaran",
                "lintasDisiplin",
                "tujuanPembelajaran",
                "topikPembelajaran",
                "praktikPedagogis",
                "kemitraanPembelajaran",
                "lingkunganBelajar",
                "pemanfaatanDigital"
              ],
            },
            langkahPembelajaran: {
              type: Type.OBJECT,
              properties: {
                pendahuluan: {
                  type: Type.OBJECT,
                  properties: {
                    durasi: { type: Type.STRING },
                    guru: { type: Type.ARRAY, items: { type: Type.STRING } },
                    siswa: { type: Type.ARRAY, items: { type: Type.STRING } },
                    pendekatan: { type: Type.STRING },
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"],
                },
                intiEksplorasi: {
                  type: Type.OBJECT,
                  properties: {
                    durasi: { type: Type.STRING },
                    guru: { type: Type.ARRAY, items: { type: Type.STRING } },
                    siswa: { type: Type.ARRAY, items: { type: Type.STRING } },
                    pendekatan: { type: Type.STRING },
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"],
                },
                intiRancangan: {
                  type: Type.OBJECT,
                  properties: {
                    durasi: { type: Type.STRING },
                    guru: { type: Type.ARRAY, items: { type: Type.STRING } },
                    siswa: { type: Type.ARRAY, items: { type: Type.STRING } },
                    pendekatan: { type: Type.STRING },
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"],
                },
                penutup: {
                  type: Type.OBJECT,
                  properties: {
                    durasi: { type: Type.STRING },
                    guru: { type: Type.ARRAY, items: { type: Type.STRING } },
                    siswa: { type: Type.ARRAY, items: { type: Type.STRING } },
                    pendekatan: { type: Type.STRING },
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"],
                },
              },
              required: ["pendahuluan", "intiEksplorasi", "intiRancangan", "penutup"],
            },
            asesmen: {
              type: Type.OBJECT,
              properties: {
                diagnostik: {
                  type: Type.OBJECT,
                  properties: {
                    metode: { type: Type.STRING },
                    instrumen: { type: Type.STRING },
                    kriteria: { type: Type.STRING },
                  },
                  required: ["metode", "instrumen", "kriteria"],
                },
                formatif: {
                  type: Type.OBJECT,
                  properties: {
                    metode: { type: Type.STRING },
                    instrumen: { type: Type.STRING },
                    kriteria: { type: Type.STRING },
                  },
                  required: ["metode", "instrumen", "kriteria"],
                },
                sumatif: {
                  type: Type.OBJECT,
                  properties: {
                    metode: { type: Type.STRING },
                    instrumen: { type: Type.STRING },
                    kriteria: { type: Type.STRING },
                  },
                  required: ["metode", "instrumen", "kriteria"],
                },
              },
              required: ["diagnostik", "formatif", "sumatif"],
            },
            remedialPengayaan: {
              type: Type.OBJECT,
              properties: {
                remedial: {
                  type: Type.OBJECT,
                  properties: {
                    sasaran: { type: Type.STRING },
                    aktivitas: { type: Type.ARRAY, items: { type: Type.STRING } },
                  },
                  required: ["sasaran", "aktivitas"],
                },
                pengayaan: {
                  type: Type.OBJECT,
                  properties: {
                    sasaran: { type: Type.STRING },
                    aktivitas: { type: Type.ARRAY, items: { type: Type.STRING } },
                  },
                  required: ["sasaran", "aktivitas"],
                },
              },
              required: ["remedial", "pengayaan"],
            },
            refleksi: {
              type: Type.OBJECT,
              properties: {
                siswa: { type: Type.ARRAY, items: { type: Type.STRING } },
                guru: { type: Type.ARRAY, items: { type: Type.STRING } },
              },
              required: ["siswa", "guru"],
            },
          },
          required: [
            "judul",
            "identitas",
            "identifikasi",
            "desainPembelajaran",
            "langkahPembelajaran",
            "asesmen",
            "remedialPengayaan",
            "refleksi"
          ],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating RPM-single:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Rencana Pembelajaran Mendalam." });
  }
});

// 7. GENERATE LKPD SINGLE
app.post("/api/generate/lkpd-single", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, topikItem } = req.body;

  const prompt = `
    Anda adalah ahli kurikulum sekolah dan pembuat bahan ajar interaktif di Indonesia.
    Tugas Anda adalah membuat Lembar Kerja Peserta Didik (LKPD) yang menarik, menantang, dan sistematis sesuai Kurikulum Merdeka untuk topik berikut:

    Mata Pelajaran: ${meta.mataPelajaran}
    Fase: ${meta.fase.toUpperCase()}
    Topik Pembelajaran: ${topikItem.topik}
    Tujuan Pembelajaran: [${topikItem.tpKode}] ${topikItem.tpTeks}
    Alokasi Waktu: ${topikItem.alokasiWaktu}
    Semester: Semester ${topikItem.semester}

    LKPD harus memiliki format dan konten profesional berkualitas tinggi. Generate detail konseptual yang nyata, bukan placeholder.
    Berikan:
    1. Judul LKPD yang representatif dan relevan dengan topik.
    2. Identitas LKPD lengkap (Nama Sekolah, Mata Pelajaran, Kelas/Semester, Alokasi Waktu, Nama Guru, Semester & Tahun Ajaran, Topik, Model Pembelajaran).
    3. Tujuan Pembelajaran (minimal 3 butir tujuan pembelajaran konkret yang operasional menggunakan kata kerja taksonomi Bloom).
    4. Petunjuk Pengerjaan yang sistematis (minimal 4 langkah panduan pengerjaan).
    5. Ringkasan Materi Konseptual (minimal 3 sampai 4 komponen utama materi dengan penjelasan teoretis mendalam dan ringkasan konseptual terstruktur).
    6. Daftar Aktivitas / Studi Kasus: Buatlah minimal 3 sampai 5 Aktivitas pemecahan masalah (troubleshooting) atau proyek mandiri/kelompok yang sangat menantang dan realistis di industri terkait topik tersebut. Setiap aktivitas harus memiliki judul yang menarik (misal: "Aktivitas 1: ...") dan teks deskripsi studi kasus yang detail disertai pertanyaan analisis/perancangan mendalam.
    7. Rubrik Penilaian ringkas (deskripsi teks kriteria kelulusan).

    Format keluaran harus merupakan JSON murni sesuai skema berikut:
    {
      "topik": "${topikItem.topik}",
      "tpKode": "${topikItem.tpKode}",
      "judul": "LEMBAR KERJA PESERTA DIDIK (LKPD)",
      "identitas": {
        "sekolah": "${meta.jenjangSekolah}",
        "mataPelajaran": "${meta.mataPelajaran}",
        "kelasSemester": "Kelas ${meta.fase === 'f' ? 'XI' : 'X'} / Semester ${topikItem.semester}",
        "alokasiWaktu": "${topikItem.alokasiWaktu}",
        "namaGuru": "${meta.namaGuru}",
        "semesterTahun": "Semester ${topikItem.semester === 1 ? 'Ganjil' : 'Genap'} - ${meta.tahunAjaran}",
        "topikMateri": "${topikItem.topik}",
        "modelPembelajaran": "Tatap Muka / Reguler"
      },
      "tujuanPembelajaran": [
        "Peserta didik mampu menganalisis konsep dasar ${topikItem.topik} secara kritis.",
        "Peserta didik mampu menerapkan prinsip ${topikItem.topik} untuk memecahkan masalah industri.",
        "Peserta didik mampu merancang skema implementasi praktis terkait ${topikItem.topik}."
      ],
      "petunjukPengerjaan": [
        "Pelajari ringkasan materi yang disediakan di bawah ini dengan saksama.",
        "Kerjakan secara mandiri atau berkelompok sesuai instruksi guru pengampu.",
        "Gunakan kemampuan pemecahan masalah (troubleshooting) untuk menganalisis setiap studi kasus.",
        "Tuliskan rancangan, perhitungan, dan argumen Anda secara detail pada kotak lembar jawaban yang telah disediakan di bawah tiap soal."
      ],
      "ringkasanMateri": [
        {
          "komponen": "Definisi & Konsep Dasar",
          "ringkasan": "Penjelasan mendalam tentang dasar-dasar topik ini..."
        }
      ],
      "aktivitasList": [
        {
          "judul": "Aktivitas 1: Desain Implementasi Praktis",
          "studiKasus": "Studi Kasus:\\nSebuah skenario dunia nyata tentang ${topikItem.topik}. Rancanglah solusi terbaik dengan mempertimbangkan efisiensi..."
        }
      ],
      "rubrikPenilaian": "Keaktifan berdiskusi (25%), Ketepatan analisis konsep (50%), Kejelasan presentasi hasil (25%)."
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            topik: { type: Type.STRING },
            tpKode: { type: Type.STRING },
            judul: { type: Type.STRING },
            identitas: {
              type: Type.OBJECT,
              properties: {
                sekolah: { type: Type.STRING },
                mataPelajaran: { type: Type.STRING },
                kelasSemester: { type: Type.STRING },
                alokasiWaktu: { type: Type.STRING },
                namaGuru: { type: Type.STRING },
                semesterTahun: { type: Type.STRING },
                topikMateri: { type: Type.STRING },
                modelPembelajaran: { type: Type.STRING },
              },
              required: [
                "sekolah",
                "mataPelajaran",
                "kelasSemester",
                "alokasiWaktu",
                "namaGuru",
                "semesterTahun",
                "topikMateri",
                "modelPembelajaran"
              ],
            },
            tujuanPembelajaran: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
            petunjukPengerjaan: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
            ringkasanMateri: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  komponen: { type: Type.STRING },
                  ringkasan: { type: Type.STRING },
                },
                required: ["komponen", "ringkasan"],
              }
            },
            aktivitasList: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  judul: { type: Type.STRING },
                  studiKasus: { type: Type.STRING },
                },
                required: ["judul", "studiKasus"],
              }
            },
            rubrikPenilaian: { type: Type.STRING },
          },
          required: [
            "topik",
            "tpKode",
            "judul",
            "identitas",
            "tujuanPembelajaran",
            "petunjukPengerjaan",
            "ringkasanMateri",
            "aktivitasList"
          ],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating LKPD-single:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate LKPD." });
  }
});

// 8. GENERATE ASESMEN SINGLE
app.post("/api/generate/asesmen-single", async (req, res) => {
  if (!ensureApiKey(res)) return;

  const { meta, topikItem } = req.body;

  const prompt = `
    Anda adalah ahli asesmen sekolah dan pembuat bank soal profesional di Indonesia.
    Tugas Anda adalah membuat dokumen 'Kisi-kisi dan Soal Asesmen' lengkap dan terstruktur untuk topik pembelajaran berikut:

    Mata Pelajaran: ${meta.mataPelajaran}
    Fase: ${meta.fase.toUpperCase()}
    Kelas/Semester: ${meta.fase === "f" ? "XI" : "X"} / Ganjil
    Topik Pembelajaran: ${topikItem.topik}
    Tujuan Pembelajaran Utama: [${topikItem.tpKode}] ${topikItem.tpTeks}
    Alokasi Waktu: ${topikItem.alokasiWaktu}

    DOKUMEN HARUS TERSTRUKTUR SEBAGAI BERIKUT:
    1. Identitas administratif lengkap.
    2. Tabel KISI-KISI SOAL yang berisi tepat 10 butir kisi-kisi:
       - Soal 1-7 berjenis "Pilihan Ganda" dengan tingkat kognitif C1 (Mudah) atau C2 (Mudah).
       - Soal 8 berjenis "Uraian" dengan tingkat kognitif C3 (Sedang).
       - Soal 9 berjenis "Uraian" dengan tingkat kognitif C4 (Sedang) untuk analisis.
       - Soal 10 berjenis "Uraian" dengan tingkat kognitif C5/C6 (HOTS) untuk evaluasi/rancangan desain.
    3. Daftar BUTIR SOAL ASESMEN yang sesuai dengan Kisi-kisi di atas:
       - Soal 1 s/d 7 wajib memiliki pilihan A, B, C, D, E.
       - Soal 8 s/d 10 berupa soal Uraian (tidak memiliki pilihan ganda).
    4. KUNCI JAWABAN & RUBRIK PENILAIAN lengkap untuk seluruh soal:
       - Pilihan Ganda (1-7): Skor Maksimal 5 per soal (total 35). Kunci jawaban menyebutkan pilihan dan penjelasannya secara singkat, contoh: "C. Topologi Bus".
       - Uraian (8-10):
         * Soal 8: Skor Maksimal 15. Berikan "Kunci Jawaban & Langkah Penyelesaian" lengkap, disertai "Kriteria Penskoran" yang rinci.
         * Soal 9: Skor Maksimal 20. Berikan "Kunci Jawaban Analisis" yang mendalam secara sistematis, disertai "Kriteria Penskoran" yang rinci.
         * Soal 10: Skor Maksimal 30. Berikan "Rancangan Desain Solusi (HOTS)" komprehensif, disertai "Kriteria Penskoran" yang rinci.
       - Total skor maksimal keseluruhan adalah 100.

    Format keluaran harus berupa JSON murni sesuai skema berikut:
    {
      "topik": "${topikItem.topik}",
      "tpKode": "${topikItem.tpKode}",
      "judul": "Kisi-kisi dan Soal Asesmen - ${topikItem.topik}",
      "kisiKisiUmum": "Asesmen ini mengukur pemahaman siswa tentang ${topikItem.topik} dengan indikator soal meliputi pemahaman teoretis, analisis kasus, dan aplikasi praktis.",
      "identitas": {
        "namaGuru": "${meta.namaGuru || "Rudi Akbar Saragih, S.Kom."}",
        "sekolah": "${meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan"}",
        "mataPelajaran": "${meta.mataPelajaran}",
        "semesterTahun": "Ganjil - ${meta.tahunAjaran || "2026/2027"}",
        "topikMateri": "${topikItem.topik}",
        "modelPembelajaran": "Tatap Muka / Reguler"
      },
      "kisiKisi": [
        {
          "no": 1,
          "tujuanPembelajaran": "Mengidentifikasi karakteristik berbagai jenis topologi...",
          "indikatorSoal": "Peserta didik dapat mengidentifikasi...",
          "levelKognitif": "C1 (Mudah)",
          "bentukSoal": "Pilihan Ganda",
          "noSoal": 1
        }
      ],
      "daftarSoal": [
        {
          "no": 1,
          "tipe": "Pilihan Ganda",
          "levelKognitif": "C1 (Mudah)",
          "soal": "Pertanyaan soal...",
          "pilihan": {
            "A": "Jawaban A",
            "B": "Jawaban B",
            "C": "Jawaban C",
            "D": "Jawaban D",
            "E": "Jawaban E"
          },
          "kunciJawaban": "C. Topologi Bus",
          "skorMaksimal": 5
        },
        {
          "no": 8,
          "tipe": "Uraian",
          "levelKognitif": "C3 (Sedang)",
          "soal": "Pertanyaan uraian...",
          "pilihan": null,
          "kunciJawaban": "Kunci Jawaban & Langkah Penyelesaian:...\\n\\nKriteria Penskoran:...",
          "skorMaksimal": 15
        }
      ]
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            topik: { type: Type.STRING },
            tpKode: { type: Type.STRING },
            judul: { type: Type.STRING },
            kisiKisiUmum: { type: Type.STRING },
            identitas: {
              type: Type.OBJECT,
              properties: {
                namaGuru: { type: Type.STRING },
                sekolah: { type: Type.STRING },
                mataPelajaran: { type: Type.STRING },
                semesterTahun: { type: Type.STRING },
                topikMateri: { type: Type.STRING },
                modelPembelajaran: { type: Type.STRING },
              },
              required: ["namaGuru", "sekolah", "mataPelajaran", "semesterTahun", "topikMateri", "modelPembelajaran"],
            },
            kisiKisi: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  no: { type: Type.INTEGER },
                  tujuanPembelajaran: { type: Type.STRING },
                  indikatorSoal: { type: Type.STRING },
                  levelKognitif: { type: Type.STRING },
                  bentukSoal: { type: Type.STRING },
                  noSoal: { type: Type.INTEGER },
                },
                required: ["no", "tujuanPembelajaran", "indikatorSoal", "levelKognitif", "bentukSoal", "noSoal"],
              },
            },
            daftarSoal: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  no: { type: Type.INTEGER },
                  tipe: { type: Type.STRING },
                  levelKognitif: { type: Type.STRING },
                  soal: { type: Type.STRING },
                  pilihan: {
                    type: Type.OBJECT,
                    properties: {
                      A: { type: Type.STRING },
                      B: { type: Type.STRING },
                      C: { type: Type.STRING },
                      D: { type: Type.STRING },
                      E: { type: Type.STRING },
                    },
                  },
                  kunciJawaban: { type: Type.STRING },
                  skorMaksimal: { type: Type.INTEGER },
                },
                required: ["no", "tipe", "levelKognitif", "soal", "kunciJawaban", "skorMaksimal"],
              },
            },
          },
          required: ["topik", "tpKode", "judul", "kisiKisiUmum", "identitas", "kisiKisi", "daftarSoal"],
        },
      },
    });

    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error: any) {
    console.error("Error generating Asesmen-single:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Asesmen." });
  }
});

// 6. DOWNLOAD DOCX GENERATION ENDPOINT
app.post("/api/download-docx", async (req, res) => {
  const { type, metadata, data } = req.body;

  if (!type || !metadata || !data) {
    return res.status(400).json({ error: "Missing type, metadata or data parameters." });
  }

  try {
    let doc;
    let filename = "dokumen.docx";

    switch (type) {
      case "cp":
        doc = generateCPDoc(metadata, data);
        filename = `01_Capaian_Pembelajaran_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "atp":
        doc = generateATPDoc(metadata, data);
        filename = `02_Analisis_CP_dan_ATP_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "prota-promes":
        doc = generateProtaPromesDoc(metadata, data);
        filename = `03_Prota_dan_Promes_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "kktp":
        doc = generateKKTPDoc(metadata, data);
        filename = `04_KKTP_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "topik":
        doc = generateTopikDoc(metadata, data);
        filename = `05_Daftar_Topik_Pembelajaran_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "rpm":
        doc = generateRPMDoc(metadata, data);
        filename = `06_Rencana_Pembelajaran_Mendalam_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "lkpd":
        doc = generateLKPDDoc(metadata, data);
        filename = `07_LKPD_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      case "asesmen":
        doc = generateAsesmenDoc(metadata, data);
        filename = `08_Soal_Asesmen_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
        break;
      default:
        return res.status(400).json({ error: "Invalid document type." });
    }

    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (error: any) {
    console.error("Error exporting DOCX:", error);
    res.status(500).json({ error: error.message || "Gagal mengekspor dokumen DOCX." });
  }
});

// ==========================================
// VITE OR STATIC SERVING MIDDLEWARE
// ==========================================

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
