import { useState, useEffect } from "react";
import { 
  FileText, 
  CheckCircle, 
  Download, 
  RefreshCw, 
  ArrowRight, 
  ArrowLeft, 
  Edit3, 
  Eye, 
  AlertCircle, 
  BookOpen, 
  User, 
  Calendar, 
  Award, 
  Check, 
  Plus, 
  Trash2, 
  FileCheck, 
  Settings, 
  FileSpreadsheet,
  Layers,
  HelpCircle
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import JSZip from "jszip";
import { 
  TeacherMetadata, 
  CPData, 
  ATPData, 
  ProtaPromesData, 
  KKTPData, 
  RPMData,
  TopikData,
  TopikItem,
  LKPDData,
  AsesmenData
} from "./types.js";

const API_BASE = (import.meta as any).env.VITE_API_BASE_URL || "";

const DIMENSI_PROFIL_LULUSAN = [
  "Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia",
  "Bernalar Kritis dan Solutif",
  "Kreatif dan Inovatif",
  "Bergotong Royong dan Kolaboratif",
  "Mandiri dan Bertanggung Jawab",
  "Berkebinekaan Global dan Inklusif",
  "Literat dan Berwawasan Teknologi",
  "Berkarakter, Sehat Jasmani dan Rohani"
];

const DEFAULT_METADATA: TeacherMetadata = {
  namaGuru: "Budi Santoso, S.Pd.",
  nipGuru: "198503122010011005",
  namaKepalaSekolah: "Dr. H. Ahmad Fauzi, M.Pd.",
  nipKepalaSekolah: "197204151998031002",
  mataPelajaran: "Pendidikan Pancasila",
  jenjangSekolah: "SD Negeri 1 Merdeka",
  fase: "A",
  tahunAjaran: "2026/2027"
};

const INDONESIAN_QUOTES = [
  "Menganalisis regulasi kurikulum terbaru...",
  "Menyelaraskan dengan Keputusan Kepala BSKAP Nomor 046/H/KR/2025...",
  "Memproses materi Pendidikan Agama sesuai Keputusan Kepala BKPDM Nomor 020 Tahun 2026...",
  "Menyusun Capaian Pembelajaran Elemen secara komprehensif...",
  "Mengekstrak Tujuan Pembelajaran esensial...",
  "Merancang Alur Tujuan Pembelajaran secara logis...",
  "Menghitung distribusi jam pelajaran per semester...",
  "Membuat rubrik penilaian interval Kriteria Ketuntasan...",
  "Menyusun Rencana Pembelajaran Mendalam (RPM) yang aktif dan kreatif..."
];

export default function App() {
  // Wizard Steps:
  // 0: Form Input
  // 1: Capaian Pembelajaran (CP)
  // 2: Analisis CP & ATP
  // 3: Prota & Promes
  // 4: KKTP
  // 5: Daftar Topik Pembelajaran
  // 6: Rencana Pembelajaran Mendalam (RPM)
  // 7: LKPD
  // 8: Soal Asesmen & Kisi-kisi
  const [step, setStep] = useState<number>(0);
  const [metadata, setMetadata] = useState<TeacherMetadata>(DEFAULT_METADATA);
  
  // Document states
  const [cpData, setCpData] = useState<CPData | null>(null);
  const [atpData, setAtpData] = useState<ATPData | null>(null);
  const [protaPromesData, setProtaPromesData] = useState<ProtaPromesData | null>(null);
  const [kktpData, setKktpData] = useState<KKTPData | null>(null);
  const [topikData, setTopikData] = useState<TopikData | null>(null);
  const [rpmList, setRpmList] = useState<{[key: string]: RPMData}>({});
  const [lkpdList, setLkpdList] = useState<{[key: string]: LKPDData}>({});
  const [asesmenList, setAsesmenList] = useState<{[key: string]: AsesmenData}>({});

  // Active topic index for detail preview under steps 6, 7, and 8
  const [selectedRpmTopic, setSelectedRpmTopic] = useState<string>("");
  const [selectedLkpdTopic, setSelectedLkpdTopic] = useState<string>("");
  const [selectedAsesmenTopic, setSelectedAsesmenTopic] = useState<string>("");

  // App behaviors
  const [loading, setLoading] = useState<boolean>(false);
  const [loadingMsgIdx, setLoadingMsgIdx] = useState<number>(0);
  const [loadingProgress, setLoadingProgress] = useState<number>(0);
  const [error, setError] = useState<string | null>(null);
  const [editMode, setEditMode] = useState<boolean>(false); // false: Preview (Tinjau), true: Edit (Kustomisasi)

  // Rotating loading messages
  useEffect(() => {
    let interval: any;
    if (loading) {
      interval = setInterval(() => {
        setLoadingMsgIdx((prev) => (prev + 1) % INDONESIAN_QUOTES.length);
      }, 3500);
    } else {
      setLoadingMsgIdx(0);
    }
    return () => clearInterval(interval);
  }, [loading]);

  // Loading progress bar simulator
  useEffect(() => {
    let interval: any;
    if (loading) {
      setLoadingProgress(5);
      interval = setInterval(() => {
        setLoadingProgress((prev) => {
          if (prev >= 95) return 95; // hold near end until finished
          return prev + Math.floor(Math.random() * 8) + 1;
        });
      }, 500);
    } else {
      setLoadingProgress(0);
    }
    return () => clearInterval(interval);
  }, [loading]);

  // Common subjects helper
  const handleSelectSubject = (subj: string) => {
    setMetadata(prev => ({ ...prev, mataPelajaran: subj }));
  };

  // REST API calls
  const generateCP = async (isRecreate = false) => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/cp`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(metadata),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate CP.");
      setCpData(data);
      if (!isRecreate) {
        setStep(1);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateATP = async (isRecreate = false) => {
    if (!cpData) return;
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/atp`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, cpData }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate ATP.");
      setAtpData(data);
      if (!isRecreate) {
        setStep(2);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateProtaPromes = async (isRecreate = false) => {
    if (!atpData) return;
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/prota-promes`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, atpData }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate Prota & Promes.");
      setProtaPromesData(data);
      if (!isRecreate) {
        setStep(3);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateKKTP = async (isRecreate = false) => {
    if (!atpData) return;
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/kktp`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, atpData }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate KKTP.");
      setKktpData(data);
      if (!isRecreate) {
        setStep(4);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateTopik = async (isRecreate = false) => {
    if (!atpData) return;
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/topik`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, atpData }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate Daftar Topik.");
      setTopikData(data);
      if (data.daftarTopik && data.daftarTopik.length > 0) {
        setSelectedRpmTopic(data.daftarTopik[0].topik);
        setSelectedLkpdTopic(data.daftarTopik[0].topik);
        setSelectedAsesmenTopic(data.daftarTopik[0].topik);
      }
      if (!isRecreate) {
        setStep(5);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateRpmSingle = async (topikItem: any) => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/rpm-single`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, topikItem }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate RPM untuk topik ini.");
      setRpmList(prev => ({ ...prev, [topikItem.topik]: data }));
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateAllRpm = async () => {
    if (!topikData || !topikData.daftarTopik) return;
    setLoading(true);
    setError(null);
    try {
      const newList = { ...rpmList };
      for (const t of topikData.daftarTopik) {
        setLoadingProgress(Math.floor((t.no / topikData.daftarTopik.length) * 100));
        const response = await fetch(`${API_BASE}/api/generate/rpm-single`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ meta: metadata, topikItem: t }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Gagal menggenerate RPM untuk topik: ${t.topik}`);
        newList[t.topik] = data;
      }
      setRpmList(newList);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
      setLoadingProgress(0);
    }
  };

  const generateLkpdSingle = async (topikItem: any) => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/lkpd-single`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, topikItem }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate LKPD untuk topik ini.");
      setLkpdList(prev => ({ ...prev, [topikItem.topik]: data }));
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateAllLkpd = async () => {
    if (!topikData || !topikData.daftarTopik) return;
    setLoading(true);
    setError(null);
    try {
      const newList = { ...lkpdList };
      for (const t of topikData.daftarTopik) {
        setLoadingProgress(Math.floor((t.no / topikData.daftarTopik.length) * 100));
        const response = await fetch(`${API_BASE}/api/generate/lkpd-single`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ meta: metadata, topikItem: t }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Gagal menggenerate LKPD untuk topik: ${t.topik}`);
        newList[t.topik] = data;
      }
      setLkpdList(newList);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
      setLoadingProgress(0);
    }
  };

  const generateAsesmenSingle = async (topikItem: any) => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`${API_BASE}/api/generate/asesmen-single`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ meta: metadata, topikItem }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menggenerate Asesmen untuk topik ini.");
      setAsesmenList(prev => ({ ...prev, [topikItem.topik]: data }));
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const generateAllAsesmen = async () => {
    if (!topikData || !topikData.daftarTopik) return;
    setLoading(true);
    setError(null);
    try {
      const newList = { ...asesmenList };
      for (const t of topikData.daftarTopik) {
        setLoadingProgress(Math.floor((t.no / topikData.daftarTopik.length) * 100));
        const response = await fetch(`${API_BASE}/api/generate/asesmen-single`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ meta: metadata, topikItem: t }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Gagal menggenerate Asesmen untuk topik: ${t.topik}`);
        newList[t.topik] = data;
      }
      setAsesmenList(newList);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
      setLoadingProgress(0);
    }
  };

  // DOCX Single file downloader helper
  const downloadDocx = async (type: string, docData: any) => {
    try {
      let dataToSend = docData;
      if (type === "atp" && docData) {
        // Enrich atpData with cpData elements' capaian fields if missing
        const enrichedAtp = docData.atp.map((item: any) => {
          if (!item.capaian) {
            const correspondingCp = cpData?.elemen?.find(
              (e: any) => e.nama.toLowerCase() === item.elemen.toLowerCase()
            );
            return {
              ...item,
              capaian: correspondingCp?.capaian || correspondingCp?.deskripsi || "",
            };
          }
          return item;
        });
        dataToSend = { ...docData, atp: enrichedAtp };
      }

      const response = await fetch(`${API_BASE}/api/download-docx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type, metadata, data: dataToSend }),
      });
      if (!response.ok) {
        const errJson = await response.json();
        throw new Error(errJson.error || "Gagal mendownload DOCX");
      }
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      
      let filename = `dokumen_${type}.docx`;
      if (type === "cp") filename = `01_Capaian_Pembelajaran_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
      if (type === "atp") filename = `02_Analisis_CP_dan_ATP_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
      if (type === "prota-promes") filename = `03_Prota_dan_Promes_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
      if (type === "kktp") filename = `04_KKTP_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
      if (type === "topik") filename = `05_Daftar_Topik_Pembelajaran_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`;
      if (type === "rpm") filename = `06_Rencana_Pembelajaran_Mendalam_${(docData.topik || metadata.mataPelajaran).replace(/\s+/g, "_")}.docx`;
      if (type === "lkpd") filename = `07_LKPD_${(docData.topik || metadata.mataPelajaran).replace(/\s+/g, "_")}.docx`;
      if (type === "asesmen") filename = `08_Soal_Asesmen_${(docData.topik || metadata.mataPelajaran).replace(/\s+/g, "_")}.docx`;

      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      alert("Error mengunduh dokumen: " + err.message);
    }
  };

  // ZIP Downloader for all documents
  const downloadAllAsZip = async () => {
    setLoading(true);
    setError(null);
    try {
      const zip = new JSZip();

      const fetchBlob = async (type: string, data: any) => {
        let dataToSend = data;
        if (type === "atp" && data) {
          const enrichedAtp = data.atp.map((item: any) => {
            if (!item.capaian) {
              const correspondingCp = cpData?.elemen?.find(
                (e: any) => e.nama.toLowerCase() === item.elemen.toLowerCase()
              );
              return {
                ...item,
                capaian: correspondingCp?.capaian || correspondingCp?.deskripsi || "",
              };
            }
            return item;
          });
          dataToSend = { ...data, atp: enrichedAtp };
        }

        const response = await fetch(`${API_BASE}/api/download-docx`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ type, metadata, data: dataToSend }),
        });
        if (!response.ok) throw new Error(`Gagal mendownload ${type}`);
        return await response.blob();
      };

      if (cpData) {
        const blob = await fetchBlob("cp", cpData);
        zip.file(`01_Capaian_Pembelajaran_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`, blob);
      }
      if (atpData) {
        const blob = await fetchBlob("atp", atpData);
        zip.file(`02_Analisis_CP_dan_ATP_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`, blob);
      }
      if (protaPromesData) {
        const blob = await fetchBlob("prota-promes", protaPromesData);
        zip.file(`03_Prota_dan_Promes_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`, blob);
      }
      if (kktpData) {
        const blob = await fetchBlob("kktp", kktpData);
        zip.file(`04_KKTP_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`, blob);
      }
      if (topikData) {
        const blob = await fetchBlob("topik", topikData);
        zip.file(`05_Daftar_Topik_Pembelajaran_${metadata.mataPelajaran.replace(/\s+/g, "_")}.docx`, blob);
      }
      
      // RPM for all topics
      if (topikData && topikData.daftarTopik) {
        for (const t of topikData.daftarTopik) {
          const rpm = rpmList[t.topik];
          if (rpm) {
            const blob = await fetchBlob("rpm", rpm);
            zip.file(`06_Rencana_Pembelajaran_Mendalam_Topik_${t.no}_${t.topik.replace(/\s+/g, "_")}.docx`, blob);
          }
        }
      }

      // LKPD for all topics
      if (topikData && topikData.daftarTopik) {
        for (const t of topikData.daftarTopik) {
          const lkpd = lkpdList[t.topik];
          if (lkpd) {
            const blob = await fetchBlob("lkpd", lkpd);
            zip.file(`07_LKPD_Topik_${t.no}_${t.topik.replace(/\s+/g, "_")}.docx`, blob);
          }
        }
      }

      // Asesmen for all topics
      if (topikData && topikData.daftarTopik) {
        for (const t of topikData.daftarTopik) {
          const asesmen = asesmenList[t.topik];
          if (asesmen) {
            const blob = await fetchBlob("asesmen", asesmen);
            zip.file(`08_Soal_Asesmen_Topik_${t.no}_${t.topik.replace(/\s+/g, "_")}.docx`, blob);
          }
        }
      }

      const content = await zip.generateAsync({ type: "blob" });
      const url = window.URL.createObjectURL(content);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Administrasi_Guru_Merdeka_${metadata.mataPelajaran.replace(/\s+/g, "_")}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      alert("Gagal mendownload semua berkas: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  // Step names translation helper
  const stepNames = [
    "Identitas",
    "CP",
    "Analisis CP & ATP",
    "Prota & Promes",
    "KKTP",
    "Daftar Topik Pembelajaran",
    "Rencana Pembelajaran Mendalam",
    "LKPD",
    "Soal Asesmen & Kisi-kisi"
  ];

  return (
    <div className="min-h-screen bg-[#f8fafc] text-slate-800 flex flex-col font-sans">
      {/* Header Bar */}
      <header className="h-14 bg-white border-b border-slate-200 flex items-center justify-between px-6 shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center text-white font-bold">S</div>
          <h1 className="text-lg font-bold tracking-tight text-slate-900">SAGO<span className="text-blue-600">Admin</span></h1>
          <span className="ml-4 px-2 py-0.5 bg-slate-100 text-slate-500 rounded text-[10px] font-semibold uppercase tracking-widest">v4.0 PRO</span>
        </div>
        <div className="flex items-center gap-4 text-sm">
          <div className="flex items-center gap-2 text-slate-500 text-xs sm:text-sm">
            <span className="w-2 h-2 bg-green-500 rounded-full"></span>
            Server Active: Reg-2025/2026
          </div>
        </div>
      </header>

      {/* Split Navigation & Content Layout */}
      <div className="flex-1 flex flex-col lg:flex-row overflow-hidden">
        {/* Sidebar Navigation */}
        <aside className="w-full lg:w-64 bg-white border-b lg:border-b-0 lg:border-r border-slate-200 flex flex-col shrink-0">
          <div className="p-4 flex flex-col gap-4">
            <h2 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Tahapan Administrasi</h2>
            <nav className="flex flex-row lg:flex-col overflow-x-auto lg:overflow-x-visible gap-2 pb-2 lg:pb-0 scrollbar-none">
              {stepNames.map((name, idx) => {
                const isActive = step === idx;
                const isCompleted = step > idx;
                const isAccessible = idx === 0 || 
                  (idx === 1 && cpData) || 
                  (idx === 2 && atpData) || 
                  (idx === 3 && protaPromesData) || 
                  (idx === 4 && kktpData) || 
                  (idx === 5 && topikData) ||
                  (idx === 6 && topikData) ||
                  (idx === 7 && topikData) ||
                  (idx === 8 && topikData);

                if (isActive) {
                  return (
                    <button
                      key={idx}
                      onClick={() => setStep(idx)}
                      className="flex items-center gap-3 p-2.5 rounded-lg bg-blue-50 text-blue-700 border border-blue-200 text-left shrink-0 cursor-pointer w-full"
                    >
                      <span className="w-6 h-6 flex items-center justify-center bg-blue-600 text-white rounded-full text-xs font-bold shrink-0">
                        {idx + 1}
                      </span>
                      <span className="text-xs font-bold leading-none">{name}</span>
                    </button>
                  );
                } else if (isCompleted) {
                  return (
                    <button
                      key={idx}
                      onClick={() => setStep(idx)}
                      className="flex items-center gap-3 p-2.5 rounded-lg bg-green-50 text-green-700 border border-green-100 text-left shrink-0 cursor-pointer w-full"
                    >
                      <span className="w-6 h-6 flex items-center justify-center bg-green-500 text-white rounded-full text-xs font-bold shrink-0">
                        ✓
                      </span>
                      <span className="text-xs font-semibold leading-none">{name}</span>
                    </button>
                  );
                } else {
                  return (
                    <button
                      key={idx}
                      disabled={!isAccessible}
                      onClick={() => isAccessible && setStep(idx)}
                      className={`flex items-center gap-3 p-2.5 rounded-lg border border-transparent text-left shrink-0 w-full ${
                        isAccessible 
                          ? "hover:bg-slate-50 text-slate-600 cursor-pointer" 
                          : "text-slate-400 opacity-60 cursor-not-allowed"
                      }`}
                    >
                      <span className="w-6 h-6 flex items-center justify-center border-2 border-slate-200 rounded-full text-xs font-bold text-slate-400 shrink-0">
                        {idx + 1}
                      </span>
                      <span className="text-xs font-medium leading-none">{name}</span>
                    </button>
                  );
                }
              })}
            </nav>
          </div>
          
          <div className="mt-auto p-4 border-t border-slate-100 bg-slate-50 hidden lg:block">
            <button
              onClick={downloadAllAsZip}
              disabled={!(cpData || atpData || protaPromesData || kktpData || topikData)}
              className="w-full py-2 bg-blue-600 hover:bg-blue-700 text-white disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed rounded text-xs font-bold uppercase tracking-tight flex items-center justify-center gap-2 transition-colors cursor-pointer"
            >
              <span>Unduh Semua (.ZIP)</span>
            </button>
          </div>
        </aside>

        {/* Main Content Area */}
        <main className="flex-1 flex flex-col bg-slate-100 overflow-y-auto">
          {/* Metadata Summary Header */}
          <div className="bg-white border-b border-slate-200 p-4 grid grid-cols-2 md:grid-cols-4 gap-4 shrink-0 shadow-sm">
            <div>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">Nama Guru / NIP</p>
              <p className="text-sm font-bold text-slate-800 truncate">{metadata.namaGuru || "—"}</p>
              <p className="text-xs text-slate-500 truncate">{metadata.nipGuru ? `NIP. ${metadata.nipGuru}` : "—"}</p>
            </div>
            <div>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">Mata Pelajaran</p>
              <p className="text-sm font-bold text-slate-800 truncate">{metadata.mataPelajaran || "—"}</p>
              <p className="text-xs text-blue-600 truncate">Fase {metadata.fase || "—"} / Merdeka</p>
            </div>
            <div>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">Fase / Jenjang</p>
              <p className="text-sm font-bold text-slate-800 truncate">Fase {metadata.fase || "—"} / {metadata.jenjangSekolah || "—"}</p>
              <p className="text-xs text-slate-500 truncate font-mono">TA {metadata.tahunAjaran || "—"}</p>
            </div>
            <div>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">Kepala Sekolah</p>
              <p className="text-sm font-bold text-slate-800 truncate">{metadata.namaKepalaSekolah || "—"}</p>
              <p className="text-xs text-slate-500 truncate">{metadata.nipKepalaSekolah ? `NIP. ${metadata.nipKepalaSekolah}` : "—"}</p>
            </div>
          </div>

          <div className="p-4 md:p-6 flex flex-col gap-6 flex-1">
            
            {/* Error Display */}
            {error && (
              <div className="bg-red-50 border-l-4 border-red-500 p-4 rounded-md shadow-sm flex items-start gap-3">
                <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                <div className="flex-1">
                  <h3 className="font-bold text-red-800 text-sm">Terjadi Kesalahan</h3>
                  <p className="text-red-700 text-xs mt-1">{error}</p>
                </div>
                <button onClick={() => setError(null)} className="text-red-400 hover:text-red-600 text-xs font-bold px-1.5 py-0.5 rounded border border-red-200 bg-white">Tutup</button>
              </div>
            )}

            {/* Loading Spinner Overlaid screen */}
            <AnimatePresence>
              {loading && (
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4"
                >
                  <div className="bg-white rounded-xl shadow-2xl p-6 md:p-8 max-w-md w-full text-center border border-slate-100 flex flex-col items-center">
                    <div className="relative mb-6">
                      <div className="w-20 h-20 rounded-full border-4 border-slate-100 border-t-blue-600 animate-spin" />
                      <div className="absolute inset-0 flex items-center justify-center">
                        <span className="text-sm font-bold font-mono text-blue-800">{loadingProgress}%</span>
                      </div>
                    </div>
                    
                    <h3 className="text-lg font-bold text-slate-800">Menyusun Dokumen Administrasi</h3>
                    <p className="text-slate-500 text-xs mt-1 px-4">Menggunakan kecerdasan buatan Gemini untuk membuat dokumen akademik yang akurat</p>
                    
                    {/* Simulated progress bar */}
                    <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden mt-6">
                      <div className="bg-blue-600 h-full transition-all duration-300" style={{ width: `${loadingProgress}%` }} />
                    </div>

                    <div className="bg-blue-50 text-blue-800 text-xs py-3 px-4 rounded-md mt-6 border border-blue-100 w-full min-h-[4rem] flex items-center justify-center">
                      <span className="italic font-medium text-center leading-relaxed">
                        "{INDONESIAN_QUOTES[loadingMsgIdx]}"
                      </span>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Step-by-Step Renderers */}
            <div className="flex-1 flex flex-col gap-6">
          
          {/* STEP 0: THE METADATA FORM */}
          {step === 0 && (
            <motion.div 
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white rounded-xl shadow-md border border-slate-200 overflow-hidden"
            >
              <div className="bg-blue-50 text-blue-800 px-6 py-4 border-b border-slate-100 flex items-center gap-2.5">
                <Settings className="w-5 h-5 text-blue-600" />
                <h2 className="font-bold text-base sm:text-lg">Input Data Profil Guru & Sekolah</h2>
              </div>
              
              <div className="p-6 flex flex-col gap-6">
                
                {/* School Information Header Section */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                  {/* Guru Section */}
                  <div className="bg-slate-50/50 p-4 rounded-lg border border-slate-100 flex flex-col gap-4">
                    <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm border-b border-slate-100 pb-2 mb-1">
                      <User className="w-4 h-4" />
                      <h3>Profil Tenaga Pendidik</h3>
                    </div>
                    
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">NAMA LENGKAP GURU</label>
                      <input 
                        type="text" 
                        value={metadata.namaGuru}
                        onChange={(e) => setMetadata({ ...metadata, namaGuru: e.target.value })}
                        placeholder="Nama Guru beserta gelar"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">NIP GURU (Nomor Induk Pegawai)</label>
                      <input 
                        type="text" 
                        value={metadata.nipGuru}
                        onChange={(e) => setMetadata({ ...metadata, nipGuru: e.target.value })}
                        placeholder="NIP Guru (gunakan '-' jika non-PNS)"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                      />
                    </div>
                  </div>

                  {/* Kepala Sekolah Section */}
                  <div className="bg-slate-50/50 p-4 rounded-lg border border-slate-100 flex flex-col gap-4">
                    <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm border-b border-slate-100 pb-2 mb-1">
                      <Award className="w-4 h-4" />
                      <h3>Profil Kepala Sekolah</h3>
                    </div>
                    
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">NAMA KEPALA SEKOLAH</label>
                      <input 
                        type="text" 
                        value={metadata.namaKepalaSekolah}
                        onChange={(e) => setMetadata({ ...metadata, namaKepalaSekolah: e.target.value })}
                        placeholder="Nama Kepala Sekolah beserta gelar"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">NIP KEPALA SEKOLAH</label>
                      <input 
                        type="text" 
                        value={metadata.nipKepalaSekolah}
                        onChange={(e) => setMetadata({ ...metadata, nipKepalaSekolah: e.target.value })}
                        placeholder="NIP Kepala Sekolah"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                      />
                    </div>
                  </div>
                </div>

                {/* Kurikulum & Academic Section */}
                <div className="bg-slate-50/50 p-4 rounded-lg border border-slate-100 flex flex-col gap-5">
                  <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm border-b border-slate-100 pb-2">
                    <BookOpen className="w-4 h-4" />
                    <h3>Konfigurasi Akademik & Kurikulum</h3>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">NAMA MATA PELAJARAN</label>
                      <input 
                        type="text" 
                        value={metadata.mataPelajaran}
                        onChange={(e) => setMetadata({ ...metadata, mataPelajaran: e.target.value })}
                        placeholder="Contoh: Pendidikan Pancasila, Bahasa Indonesia, Matematika"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                      />
                      
                      {/* Subject Quick Picks */}
                      <div className="mt-1.5">
                        <span className="text-[10px] text-slate-400 font-bold uppercase block mb-1">Pilih Cepat Mata Pelajaran:</span>
                        <div className="flex flex-wrap gap-1">
                          {[
                            "Pendidikan Pancasila",
                            "Pendidikan Agama Islam dan Budi Pekerti",
                            "Bahasa Indonesia",
                            "Matematika",
                            "IPA",
                            "IPS",
                            "Seni Budaya"
                          ].map((subj) => (
                            <button
                              key={subj}
                              type="button"
                              onClick={() => handleSelectSubject(subj)}
                              className={`text-[10px] py-1 px-2 rounded-full border transition-all ${
                                metadata.mataPelajaran === subj 
                                  ? "bg-blue-100 text-blue-800 border-blue-300 font-bold" 
                                  : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                              }`}
                            >
                              {subj}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">NAMA INSTITUSI / SEKOLAH</label>
                      <input 
                        type="text" 
                        value={metadata.jenjangSekolah}
                        onChange={(e) => setMetadata({ ...metadata, jenjangSekolah: e.target.value })}
                        placeholder="Nama Sekolah Lengkap"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">FASE KURIKULUM MERDEKA</label>
                      <select
                        value={metadata.fase}
                        onChange={(e) => setMetadata({ ...metadata, fase: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all"
                      >
                        <option value="A">Fase A (Kelas 1-2 SD)</option>
                        <option value="B">Fase B (Kelas 3-4 SD)</option>
                        <option value="C">Fase C (Kelas 5-6 SD)</option>
                        <option value="D">Fase D (Kelas 7-9 SMP)</option>
                        <option value="E">Fase E (Kelas 10 SMA)</option>
                        <option value="F">Fase F (Kelas 11-12 SMA)</option>
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">TAHUN AJARAN</label>
                      <input 
                        type="text" 
                        value={metadata.tahunAjaran}
                        onChange={(e) => setMetadata({ ...metadata, tahunAjaran: e.target.value })}
                        placeholder="Contoh: 2026/2027"
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
                      />
                    </div>
                  </div>
                </div>

                <div className="bg-amber-50/50 p-4 rounded-lg border border-amber-200/60 flex items-start gap-3 mt-1 text-amber-900">
                  <HelpCircle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="text-xs leading-relaxed">
                    <span className="font-bold block mb-0.5">Sistem Integrasi Peraturan Perundangan:</span>
                    Aplikasi ini otomatis menyelaraskan Capaian Pembelajaran (CP) berdasarkan <strong>Keputusan Kepala BSKAP No 046/H/KR/2025</strong>. Untuk PAI, PAK, atau Pendidikan Agama Kristen/Katolik dan Budi Pekerti, dokumen akan otomatis berpedoman pada <strong>Keputusan Kepala BKPDM Nomor 020 Tahun 2026</strong>.
                  </div>
                </div>

                {/* Submit Action */}
                <div className="flex justify-end pt-2">
                  <button
                    onClick={() => generateCP()}
                    className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg shadow transition-all flex items-center gap-2 text-sm sm:text-base cursor-pointer hover:shadow-lg active:scale-[0.98]"
                  >
                    <span>Mulai Pembuatan Administrasi</span>
                    <ArrowRight className="w-5 h-5" />
                  </button>
                </div>

              </div>
            </motion.div>
          )}

          {/* SHARED ACTIONS HEADER FOR GENERATED DOCUMENTS */}
          {step > 0 && (
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 bg-white p-4 rounded-xl shadow-sm border border-slate-200">
              <div className="flex items-center gap-2.5">
                <div className="bg-blue-100 text-blue-800 p-2 rounded-lg font-bold text-xs shrink-0 uppercase tracking-wider">
                  DOKUMEN {step}/5
                </div>
                <div>
                  <h2 className="text-base sm:text-lg font-bold text-slate-800">{stepNames[step]}</h2>
                  <p className="text-slate-500 text-xs mt-0.5">Edit konten langsung pada tab Kustomisasi atau tinjau kertas A4 sebelum mengekspor.</p>
                </div>
              </div>

              {/* Toggle Preview / Edit & Single Download */}
              <div className="flex items-center flex-wrap gap-2 w-full sm:w-auto">
                <div className="flex bg-slate-100 p-1 rounded-md border border-slate-200 shrink-0">
                  <button
                    onClick={() => setEditMode(false)}
                    className={`flex items-center gap-1 text-xs font-bold py-1.5 px-3 rounded-sm transition-all ${
                      !editMode 
                        ? "bg-white text-blue-800 shadow-sm" 
                        : "text-slate-600 hover:text-slate-800"
                    }`}
                  >
                    <Eye className="w-3.5 h-3.5" />
                    <span>Tinjau</span>
                  </button>
                  <button
                    onClick={() => setEditMode(true)}
                    className={`flex items-center gap-1 text-xs font-bold py-1.5 px-3 rounded-sm transition-all ${
                      editMode 
                        ? "bg-white text-blue-800 shadow-sm" 
                        : "text-slate-600 hover:text-slate-800"
                    }`}
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                    <span>Kustomisasi</span>
                  </button>
                </div>

                <button
                  onClick={() => {
                    if (step === 1 && cpData) downloadDocx("cp", cpData);
                    if (step === 2 && atpData) downloadDocx("atp", atpData);
                    if (step === 3 && protaPromesData) downloadDocx("prota-promes", protaPromesData);
                    if (step === 4 && kktpData) downloadDocx("kktp", kktpData);
                    if (step === 5 && topikData) downloadDocx("topik", topikData);
                    if (step === 6 && rpmList[selectedRpmTopic]) downloadDocx("rpm", rpmList[selectedRpmTopic]);
                    if (step === 7 && lkpdList[selectedLkpdTopic]) downloadDocx("lkpd", lkpdList[selectedLkpdTopic]);
                    if (step === 8 && asesmenList[selectedAsesmenTopic]) downloadDocx("asesmen", asesmenList[selectedAsesmenTopic]);
                  }}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs py-2 px-3.5 rounded border border-slate-300 shadow-sm transition-all flex items-center justify-center gap-1.5 cursor-pointer flex-1 sm:flex-initial"
                >
                  <Download className="w-3.5 h-3.5 text-slate-500" />
                  <span>Unduh DOCX</span>
                </button>
              </div>
            </div>
          )}

          {/* STEP 1: CAPAIAN PEMBELAJARAN (CP) */}
          {step === 1 && cpData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              {/* Main Document Content */}
              <div className="bg-white border border-slate-200 rounded-xl shadow-md min-h-[500px] flex flex-col">
                
                {/* PREVIEW MODE (TINJAU) */}
                {!editMode ? (
                  <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-serif leading-relaxed text-sm sm:text-base text-slate-900 A4-paper">
                    {/* Header Doc */}
                    <div className="text-center border-b border-double border-slate-300 pb-4 mb-6">
                      <h1 className="font-bold text-lg sm:text-xl uppercase tracking-wide">Capaian Pembelajaran (CP)</h1>
                      <h2 className="font-bold text-base mt-1 uppercase tracking-wide">Mata Pelajaran: {metadata.mataPelajaran}</h2>
                      <p className="text-xs text-slate-500 mt-2 font-sans italic">Regulasi: {cpData.regulasi}</p>
                    </div>

                    {/* Metadata block */}
                    <div className="grid grid-cols-2 gap-x-6 gap-y-1 bg-slate-50 py-3 px-4 rounded border border-slate-200/60 text-xs font-sans text-slate-600 mb-6">
                      <div><strong className="text-slate-800">Nama Guru:</strong> {metadata.namaGuru}</div>
                      <div><strong className="text-slate-800">Jenjang/Fase:</strong> {metadata.jenjangSekolah} / Fase {metadata.fase}</div>
                      <div><strong className="text-slate-800">NIP Guru:</strong> {metadata.nipGuru}</div>
                      <div><strong className="text-slate-800">Tahun Ajaran:</strong> {metadata.tahunAjaran}</div>
                    </div>

                    {/* Content Blocks */}
                    <div className="flex flex-col gap-5 flex-1">
                      <div>
                        <h3 className="font-bold text-base border-b border-slate-200 pb-1 mb-2 font-sans text-emerald-800">1. Rasionalisasi Mata Pelajaran</h3>
                        <p className="text-justify indent-8 text-slate-700 whitespace-pre-wrap">{cpData.rasional}</p>
                      </div>

                      <div>
                        <h3 className="font-bold text-base border-b border-slate-200 pb-1 mb-2 font-sans text-emerald-800">2. Tujuan Mata Pelajaran</h3>
                        <p className="text-justify indent-8 text-slate-700 whitespace-pre-wrap">{cpData.tujuan}</p>
                      </div>

                      <div>
                        <h3 className="font-bold text-base border-b border-slate-200 pb-1 mb-2 font-sans text-emerald-800">3. Karakteristik Mata Pelajaran</h3>
                        <p className="text-justify indent-8 text-slate-700 whitespace-pre-wrap">{cpData.karakteristik}</p>
                      </div>

                      <div>
                        <h3 className="font-bold text-base border-b border-slate-200 pb-1 mb-2 font-sans text-emerald-800">4. Capaian Pembelajaran Fase {metadata.fase}</h3>
                        <p className="text-justify indent-8 text-slate-700 whitespace-pre-wrap font-bold">{cpData.capaianUmum}</p>
                      </div>

                      <div>
                        <h3 className="font-bold text-base border-b border-slate-200 pb-1 mb-3 font-sans text-emerald-800">5. Capaian Pembelajaran per Elemen</h3>
                        <div className="overflow-x-auto">
                          <table className="w-full border-collapse border border-slate-300 text-xs sm:text-sm text-left">
                            <thead>
                              <tr className="bg-slate-100">
                                <th className="border border-slate-300 p-2.5 font-bold font-sans w-1/4">Elemen</th>
                                <th className="border border-slate-300 p-2.5 font-bold font-sans w-3/4">Capaian Pembelajaran Elemen</th>
                              </tr>
                            </thead>
                            <tbody>
                              {cpData.elemen.map((el, index) => (
                                <tr key={index}>
                                  <td className="border border-slate-300 p-2.5 font-sans">
                                    <div className="font-bold text-slate-800">{el.nama}</div>
                                    <div className="text-[10px] text-slate-500 italic mt-0.5">{el.deskripsi}</div>
                                  </td>
                                  <td className="border border-slate-300 p-2.5 text-justify leading-relaxed whitespace-pre-wrap">{el.capaian}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                ) : (
                  /* EDIT MODE (KUSTOMISASI) */
                  <div className="p-6 flex flex-col gap-5">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">JUDUL DOKUMEN</label>
                      <input 
                        type="text" 
                        value={cpData.judul}
                        onChange={(e) => setCpData({ ...cpData, judul: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 font-bold"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">1. RASIONALISASI MATA PELAJARAN</label>
                      <textarea 
                        rows={6}
                        value={cpData.rasional}
                        onChange={(e) => setCpData({ ...cpData, rasional: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 leading-relaxed"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">2. TUJUAN MATA PELAJARAN</label>
                      <textarea 
                        rows={6}
                        value={cpData.tujuan}
                        onChange={(e) => setCpData({ ...cpData, tujuan: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 leading-relaxed"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">3. KARAKTERISTIK MATA PELAJARAN</label>
                      <textarea 
                        rows={6}
                        value={cpData.karakteristik}
                        onChange={(e) => setCpData({ ...cpData, karakteristik: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 leading-relaxed"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">4. CAPAIAN PEMBELAJARAN UMUM (FASE {metadata.fase})</label>
                      <textarea 
                        rows={5}
                        value={cpData.capaianUmum}
                        onChange={(e) => setCpData({ ...cpData, capaianUmum: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 font-semibold"
                      />
                    </div>

                    {/* Edit Elemen */}
                    <div className="flex flex-col gap-3">
                      <div className="flex items-center justify-between border-b border-slate-200 pb-1">
                        <label className="text-xs font-bold text-slate-500">5. CAPAIAN PEMBELAJARAN PER ELEMEN</label>
                        <button
                          onClick={() => {
                            const updated = [...cpData.elemen, { nama: "Elemen Baru", deskripsi: "Deskripsi...", capaian: "Capaian..." }];
                            setCpData({ ...cpData, elemen: updated });
                          }}
                          className="text-xs bg-emerald-50 text-emerald-800 py-1 px-2.5 rounded hover:bg-emerald-100 flex items-center gap-1 font-bold"
                        >
                          <Plus className="w-3 h-3" /> Tambah Elemen
                        </button>
                      </div>

                      {cpData.elemen.map((el, idx) => (
                        <div key={idx} className="bg-slate-50 p-4 rounded border border-slate-200/80 flex flex-col gap-3 relative">
                          <button
                            onClick={() => {
                              const updated = cpData.elemen.filter((_, i) => i !== idx);
                              setCpData({ ...cpData, elemen: updated });
                            }}
                            className="absolute top-3 right-3 text-red-500 hover:text-red-700 p-1 hover:bg-red-50 rounded"
                            title="Hapus elemen"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-400">NAMA ELEMEN</span>
                              <input 
                                type="text" 
                                value={el.nama}
                                onChange={(e) => {
                                  const updated = [...cpData.elemen];
                                  updated[idx].nama = e.target.value;
                                  setCpData({ ...cpData, elemen: updated });
                                }}
                                className="bg-white border border-slate-300 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-emerald-500 font-bold"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-400">DESKRIPSI ELEMEN</span>
                              <input 
                                type="text" 
                                value={el.deskripsi}
                                onChange={(e) => {
                                  const updated = [...cpData.elemen];
                                  updated[idx].deskripsi = e.target.value;
                                  setCpData({ ...cpData, elemen: updated });
                                }}
                                className="bg-white border border-slate-300 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-emerald-500 italic"
                              />
                            </div>
                          </div>
                          
                          <div className="flex flex-col gap-1">
                            <span className="text-[10px] font-bold text-slate-400">CAPAIAN PEMBELAJARAN ELEMEN</span>
                            <textarea 
                              rows={3}
                              value={el.capaian}
                              onChange={(e) => {
                                const updated = [...cpData.elemen];
                                updated[idx].capaian = e.target.value;
                                setCpData({ ...cpData, elemen: updated });
                              }}
                              className="bg-white border border-slate-300 rounded px-2 py-1 text-sm focus:ring-1 focus:ring-emerald-500 leading-relaxed"
                            />
                          </div>
                        </div>
                      ))}
                    </div>

                  </div>
                )}
              </div>

              {/* Step Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(0)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => generateCP(true)}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>
                  <button
                    onClick={() => {
                      if (atpData) setStep(2);
                      else generateATP();
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial hover:shadow-md"
                  >
                    <span>Buat Analisis CP & ATP</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 2: ANALISIS CP & ATP */}
          {step === 2 && atpData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              <div className="bg-white border border-slate-200 rounded-xl shadow-md min-h-[500px] flex flex-col">
                
                {/* PREVIEW MODE */}
                {!editMode ? (
                  <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-sans leading-relaxed text-slate-900 A4-paper gap-6 w-full">
                    {/* Header Banner */}
                    <div className="bg-[#007A54] text-white p-6 rounded-lg text-center shadow-sm w-full">
                      <h1 className="font-bold text-lg sm:text-2xl uppercase tracking-wider">CAPAIAN & ALUR TUJUAN PEMBELAJARAN</h1>
                      <div className="w-16 h-1 bg-white mx-auto my-3 rounded"></div>
                      <p className="text-xs sm:text-sm font-medium tracking-wide">
                        Mata Pelajaran: {metadata.mataPelajaran} | Fase/Kelas: {metadata.fase} / XI
                      </p>
                    </div>

                    {/* Identity Table */}
                    <div className="overflow-hidden border border-emerald-100 rounded-lg shadow-sm w-full">
                      <table className="w-full text-xs text-left border-collapse">
                        <tbody>
                          <tr className="border-b border-emerald-100">
                            <td className="w-1/3 bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Nama Guru / Penyusun</td>
                            <td className="p-2.5 text-slate-700">{metadata.namaGuru}</td>
                          </tr>
                          <tr className="border-b border-emerald-100">
                            <td className="bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Instansi / Sekolah</td>
                            <td className="p-2.5 text-slate-700">{metadata.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan"}</td>
                          </tr>
                          <tr className="border-b border-emerald-100">
                            <td className="bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Mata Pelajaran</td>
                            <td className="p-2.5 text-slate-700">{metadata.mataPelajaran}</td>
                          </tr>
                          <tr>
                            <td className="bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Jenjang / Fase / Kelas</td>
                            <td className="p-2.5 text-slate-700">SMK/MAK / {metadata.fase} / XI</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Capaian Pembelajaran Section */}
                    <div className="w-full">
                      <h2 className="text-sm sm:text-base font-bold text-[#007A54] border-b-2 border-[#007A54] pb-1 mb-3 uppercase tracking-wide">
                        CAPAIAN PEMBELAJARAN
                      </h2>
                      <div className="overflow-x-auto border border-slate-200 rounded-lg shadow-sm w-full">
                        <table className="w-full text-[11px] sm:text-xs text-left border-collapse">
                          <thead>
                            <tr className="bg-[#007A54] text-white">
                              <th className="p-3 font-semibold w-[30%]">Elemen</th>
                              <th className="p-3 font-semibold w-[70%]">Capaian Pembelajaran (CP)</th>
                            </tr>
                          </thead>
                          <tbody>
                            {atpData.atp.map((item, idx) => {
                              const cpContent = item.capaian || cpData?.elemen?.find(e => e.nama.toLowerCase() === item.elemen.toLowerCase())?.capaian || "";
                              return (
                                <tr key={`cp-${idx}`} className={idx % 2 === 0 ? "bg-white" : "bg-slate-50/50"}>
                                  <td className="p-3 border-b border-slate-200 font-bold text-[#007A54] align-top">{item.elemen}</td>
                                  <td className="p-3 border-b border-slate-200 text-slate-700 leading-relaxed text-justify">{cpContent}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Alur Tujuan Pembelajaran Section */}
                    <div className="w-full">
                      <h2 className="text-sm sm:text-base font-bold text-[#007A54] border-b-2 border-[#007A54] pb-1 mb-3 uppercase tracking-wide">
                        ALUR TUJUAN PEMBELAJARAN (ATP)
                      </h2>
                      <div className="overflow-x-auto border border-slate-200 rounded-lg shadow-sm w-full">
                        <table className="w-full text-[10px] sm:text-xs text-left border-collapse">
                          <thead>
                            <tr className="bg-[#007A54] text-white text-center">
                              <th className="p-2.5 border border-slate-300 font-semibold w-[15%]">Elemen</th>
                              <th className="p-2.5 border border-slate-300 font-semibold w-[25%]">Capaian Pembelajaran (CP)</th>
                              <th className="p-2.5 border border-slate-300 font-semibold w-[32%]">Tujuan Pembelajaran (TP)</th>
                              <th className="p-2.5 border border-slate-300 font-semibold w-[13%]">Kata Kunci Materi</th>
                              <th className="p-2.5 border border-slate-300 font-semibold w-[10%]">Profil Lulusan (SKL)</th>
                              <th className="p-2.5 border border-slate-300 font-semibold w-[5%]">Perkiraan JP</th>
                            </tr>
                          </thead>
                          <tbody>
                            {atpData.atp.map((item, idx) => {
                              const tps = item.tujuanPembelajaran || [];
                              const rowSpanVal = tps.length || 1;
                              const cpContent = item.capaian || cpData?.elemen?.find(e => e.nama.toLowerCase() === item.elemen.toLowerCase())?.capaian || "";

                              if (tps.length === 0) {
                                return (
                                  <tr key={`atp-empty-${idx}`} className="border-b border-slate-200">
                                    <td className="p-2.5 border border-slate-200 font-bold text-[#007A54] align-top">{item.elemen}</td>
                                    <td className="p-2.5 border border-slate-200 text-slate-600 align-top text-justify">{cpContent}</td>
                                    <td className="p-2.5 border border-slate-200 align-top">-</td>
                                    <td className="p-2.5 border border-slate-200 align-top">-</td>
                                    <td className="p-2.5 border border-slate-200 align-top">-</td>
                                    <td className="p-2.5 border border-slate-200 text-center align-top">-</td>
                                  </tr>
                                );
                              }

                              return tps.map((tpObj, tpIdx) => {
                                const formatTpTextPreview = (text: string, count: number) => {
                                  const trimmed = (text || "").trim();
                                  if (/^\d+[\.\)]/.test(trimmed)) {
                                    return trimmed;
                                  }
                                  return `${count + 1}. ${trimmed}`;
                                };

                                return (
                                  <tr key={`atp-${idx}-${tpIdx}`} className="border-b border-slate-200">
                                    {tpIdx === 0 && (
                                      <td className="p-2.5 border border-slate-200 font-bold text-[#007A54] bg-[#007A54]/5 align-top" rowSpan={rowSpanVal}>
                                        {item.elemen}
                                      </td>
                                    )}
                                    {tpIdx === 0 && (
                                      <td className="p-2.5 border border-slate-200 text-slate-600 align-top text-justify bg-[#007A54]/5" rowSpan={rowSpanVal}>
                                        {cpContent}
                                      </td>
                                    )}
                                    <td className="p-2.5 border border-slate-200 text-slate-800 align-top text-justify leading-relaxed">
                                      <span className="font-bold font-mono text-emerald-700 bg-emerald-50 px-1 rounded text-[9px] mr-1">{tpObj.kode}</span>
                                      {formatTpTextPreview(tpObj.tp, tpIdx)}
                                      {tpObj.glosarium && (
                                        <div className="text-[10px] text-slate-500 italic mt-1 font-sans">
                                          Glosarium: {tpObj.glosarium}
                                        </div>
                                      )}
                                    </td>
                                    <td className="p-2.5 border border-slate-200 text-slate-700 align-top">{tpObj.topik}</td>
                                    <td className="p-2.5 border border-slate-200 text-slate-600 text-[10px] leading-snug align-top">
                                      {tpObj.profilLulusan || tpObj.profilPancasila}
                                    </td>
                                    <td className="p-2.5 border border-slate-200 text-center font-semibold text-slate-700 align-top whitespace-nowrap">
                                      {tpObj.alokasiWaktu}
                                    </td>
                                  </tr>
                                );
                              });
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Side-by-Side Signatures */}
                    <div className="mt-8 pt-6 border-t border-slate-100 grid grid-cols-2 text-xs w-full">
                      <div className="flex flex-col gap-1 text-left">
                        <p>Mengetahui,</p>
                        <p className="font-semibold text-slate-850">Kepala Sekolah</p>
                        <div className="h-16"></div>
                        <p className="font-bold text-slate-900 underline">
                          ( {metadata.namaKepalaSekolah || "__________________________"} )
                        </p>
                        <p className="text-slate-500">NIP. {metadata.nipKepalaSekolah || "..................................."}</p>
                      </div>
                      <div className="flex flex-col gap-1 text-left pl-8 border-l border-slate-100">
                        <p>......................, 29 Juni 2026</p>
                        <p className="font-semibold text-slate-850">Guru Mata Pelajaran</p>
                        <div className="h-16"></div>
                        <p className="font-bold text-slate-900">
                          {metadata.namaGuru || "__________________________"}
                        </p>
                        <p className="text-slate-500">NIP. {metadata.nipGuru || "..................................."}</p>
                      </div>
                    </div>
                  </div>
                ) : (
                  /* EDIT MODE */
                  <div className="p-6 flex flex-col gap-6">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">JUDUL ATP</label>
                      <input 
                        type="text" 
                        value={atpData.judul}
                        onChange={(e) => setAtpData({ ...atpData, judul: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 font-bold"
                      />
                    </div>

                    <div className="flex flex-col gap-4">
                      {atpData.atp.map((item, elemIdx) => (
                        <div key={elemIdx} className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                          <div className="flex items-center justify-between border-b border-slate-200 pb-2 mb-3">
                            <div className="flex flex-col gap-1 w-2/3">
                              <span className="text-[10px] font-bold text-slate-400">ELEMEN CAPAIAN PEMBELAJARAN</span>
                              <input 
                                type="text" 
                                value={item.elemen}
                                onChange={(e) => {
                                  const updated = { ...atpData };
                                  updated.atp[elemIdx].elemen = e.target.value;
                                  setAtpData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-sm font-bold text-emerald-850"
                              />
                            </div>
                            <button
                              onClick={() => {
                                const updated = { ...atpData };
                                updated.atp[elemIdx].tujuanPembelajaran.push({
                                  kode: `TP ${metadata.fase}.${elemIdx + 1}.${item.tujuanPembelajaran.length + 1}`,
                                  tp: "Tujuan Pembelajaran Baru...",
                                  topik: "Materi Baru...",
                                  alokasiWaktu: "4 JP",
                                  profilLulusan: "Bernalar Kritis dan Solutif",
                                  profilPancasila: "Bernalar Kritis dan Solutif",
                                  glosarium: ""
                                });
                                setAtpData(updated);
                              }}
                              className="text-xs bg-emerald-100 text-emerald-800 font-bold py-1 px-2.5 rounded hover:bg-emerald-200 flex items-center gap-1"
                            >
                              <Plus className="w-3.5 h-3.5" /> Tambah TP
                            </button>
                          </div>

                          <div className="flex flex-col gap-3">
                            {item.tujuanPembelajaran.map((tpObj, tpIdx) => (
                              <div key={tpIdx} className="bg-white p-3.5 rounded border border-slate-200/80 flex flex-col gap-3 relative shadow-sm">
                                <button
                                  onClick={() => {
                                    const updated = { ...atpData };
                                    updated.atp[elemIdx].tujuanPembelajaran = item.tujuanPembelajaran.filter((_, i) => i !== tpIdx);
                                    setAtpData(updated);
                                  }}
                                  className="absolute top-2 right-2 text-red-500 hover:text-red-700 p-1 hover:bg-red-50 rounded"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>

                                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[10px] font-bold text-slate-400">KODE TP</span>
                                    <input 
                                      type="text" 
                                      value={tpObj.kode}
                                      onChange={(e) => {
                                        const updated = { ...atpData };
                                        updated.atp[elemIdx].tujuanPembelajaran[tpIdx].kode = e.target.value;
                                        setAtpData(updated);
                                      }}
                                      className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-mono font-bold text-emerald-800"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[10px] font-bold text-slate-400">TOPIK / MATERI</span>
                                    <input 
                                      type="text" 
                                      value={tpObj.topik}
                                      onChange={(e) => {
                                        const updated = { ...atpData };
                                        updated.atp[elemIdx].tujuanPembelajaran[tpIdx].topik = e.target.value;
                                        setAtpData(updated);
                                      }}
                                      className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-semibold"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[10px] font-bold text-slate-400">ALOKASI JP</span>
                                    <input 
                                      type="text" 
                                      value={tpObj.alokasiWaktu}
                                      onChange={(e) => {
                                        const updated = { ...atpData };
                                        updated.atp[elemIdx].tujuanPembelajaran[tpIdx].alokasiWaktu = e.target.value;
                                        setAtpData(updated);
                                      }}
                                      className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-bold text-slate-600"
                                    />
                                  </div>
                                </div>

                                <div className="flex flex-col gap-0.5">
                                  <span className="text-[10px] font-bold text-slate-400">RUMUSAN TUJUAN PEMBELAJARAN</span>
                                  <textarea 
                                    rows={2}
                                    value={tpObj.tp}
                                    onChange={(e) => {
                                      const updated = { ...atpData };
                                      updated.atp[elemIdx].tujuanPembelajaran[tpIdx].tp = e.target.value;
                                      setAtpData(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2 py-1 text-xs leading-relaxed"
                                  />
                                </div>

                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <span className="text-[10px] font-bold text-slate-400">DIMENSI PROFIL LULUSAN</span>
                                    <div className="border border-slate-300 rounded p-2 bg-white max-h-32 overflow-y-auto flex flex-col gap-1 shadow-inner">
                                      {DIMENSI_PROFIL_LULUSAN.map((dim, dIdx) => {
                                        const selectedList = (tpObj.profilLulusan || tpObj.profilPancasila || "")
                                          .split(",")
                                          .map(s => s.trim())
                                          .filter(Boolean);
                                        const isChecked = selectedList.includes(dim);
                                        return (
                                          <label key={dIdx} className="flex items-start gap-2 text-[10px] sm:text-[11px] text-slate-700 hover:bg-slate-50 p-1 rounded cursor-pointer select-none">
                                            <input 
                                              type="checkbox" 
                                              checked={isChecked}
                                              onChange={() => {
                                                let newList;
                                                if (isChecked) {
                                                  newList = selectedList.filter(s => s !== dim);
                                                } else {
                                                  newList = [...selectedList, dim];
                                                }
                                                const updatedValue = newList.join(", ");
                                                const updated = { ...atpData };
                                                updated.atp[elemIdx].tujuanPembelajaran[tpIdx].profilLulusan = updatedValue;
                                                updated.atp[elemIdx].tujuanPembelajaran[tpIdx].profilPancasila = updatedValue;
                                                setAtpData(updated);
                                              }}
                                              className="mt-0.5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 w-3.5 h-3.5 cursor-pointer"
                                            />
                                            <span className="leading-tight">{dim}</span>
                                          </label>
                                        );
                                      })}
                                    </div>
                                  </div>
                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[10px] font-bold text-slate-400">GLOSARIUM / KATA KUNCI</span>
                                    <input 
                                      type="text" 
                                      value={tpObj.glosarium}
                                      onChange={(e) => {
                                        const updated = { ...atpData };
                                        updated.atp[elemIdx].tujuanPembelajaran[tpIdx].glosarium = e.target.value;
                                        setAtpData(updated);
                                      }}
                                      className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs italic"
                                    />
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(1)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => generateATP(true)}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>
                  <button
                    onClick={() => {
                      if (protaPromesData) setStep(3);
                      else generateProtaPromes();
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial hover:shadow-md"
                  >
                    <span>Buat Prota dan Promes</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 3: PROTA & PROMES */}
          {step === 3 && protaPromesData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              <div className="bg-white border border-slate-200 rounded-xl shadow-md min-h-[500px] flex flex-col">
                
                {/* PREVIEW MODE */}
                {!editMode ? (
                  <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-serif leading-relaxed text-sm sm:text-base text-slate-900 A4-paper">
                    
                    {/* Header banner block */}
                    <div className="bg-[#007a54] text-white p-6 text-center rounded-lg mb-6">
                      <h1 className="font-sans font-bold text-lg sm:text-2xl uppercase tracking-wider mb-1">PROGRAM TAHUNAN & PROGRAM SEMESTER</h1>
                      <p className="font-sans text-xs sm:text-sm font-medium opacity-90">
                        Mata Pelajaran: {metadata.mataPelajaran} | Fase/Kelas: {metadata.fase} / XI
                      </p>
                    </div>

                    {/* Administration Identity Table */}
                    <div className="overflow-x-auto mb-6">
                      <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                        <tbody>
                          <tr>
                            <td className="border border-slate-300 p-2.5 font-sans font-bold bg-slate-50 w-[30%]">Nama Guru / Penyusun</td>
                            <td className="border border-slate-300 p-2.5 font-sans text-slate-800 font-medium">{metadata.namaGuru}</td>
                          </tr>
                          <tr>
                            <td className="border border-slate-300 p-2.5 font-sans font-bold bg-slate-50">Instansi / Sekolah</td>
                            <td className="border border-slate-300 p-2.5 font-sans text-slate-800 font-medium">{metadata.jenjangSekolah || "SMKS TI"}</td>
                          </tr>
                          <tr>
                            <td className="border border-slate-300 p-2.5 font-sans font-bold bg-slate-50">Mata Pelajaran</td>
                            <td className="border border-slate-300 p-2.5 font-sans text-slate-800 font-medium">{metadata.mataPelajaran}</td>
                          </tr>
                          <tr>
                            <td className="border border-slate-300 p-2.5 font-sans font-bold bg-slate-50">Jenjang / Fase / Kelas</td>
                            <td className="border border-slate-300 p-2.5 font-sans text-slate-800 font-medium">SMK/MAK / {metadata.fase} / XI</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Part A: PROTA */}
                    <h3 className="font-sans font-bold text-base text-[#007a54] uppercase tracking-wide border-b-2 border-[#007a54] pb-1.5 mb-4 mt-6">
                      PROGRAM TAHUNAN (PROTA)
                    </h3>

                    <div className="overflow-x-auto mb-10">
                      <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                        <thead>
                          <tr className="bg-[#007a54] text-white">
                            <th className="border border-slate-300 p-2.5 font-sans font-bold text-center w-[8%]">No</th>
                            <th className="border border-slate-300 p-2.5 font-sans font-bold w-[25%]">Elemen / Lingkup Materi</th>
                            <th className="border border-slate-300 p-2.5 font-sans font-bold w-[52%]">Tujuan Pembelajaran</th>
                            <th className="border border-slate-300 p-2.5 font-sans font-bold text-center w-[15%]">Alokasi Waktu (JP)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {protaPromesData.prota.map((item, index) => (
                            <tr key={index} className="hover:bg-slate-50/50">
                              <td className="border border-slate-300 p-2.5 text-center font-bold text-slate-700">{item.no || (index + 1)}</td>
                              <td className="border border-slate-300 p-2.5 font-sans font-bold text-slate-800">{item.elemen}</td>
                              <td className="border border-slate-300 p-2.5 font-sans text-slate-700 leading-relaxed text-justify">{item.tujuanPembelajaran}</td>
                              <td className="border border-slate-300 p-2.5 text-center font-bold font-sans text-slate-800">{item.alokasiWaktu}</td>
                            </tr>
                          ))}
                          <tr className="bg-slate-50 font-bold text-slate-800 border-t-2 border-slate-300">
                            <td colSpan={3} className="border border-slate-300 p-2.5 text-right uppercase tracking-wide font-sans">
                              TOTAL ALOKASI WAKTU SATU TAHUN
                            </td>
                            <td className="border border-slate-300 p-2.5 text-center font-sans text-[#007a54] font-extrabold text-sm sm:text-base">
                              {protaPromesData.prota.reduce((sum, item) => {
                                const num = parseInt(item.alokasiWaktu) || 0;
                                return sum + num;
                              }, 0)} JP
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Part B: PROMES */}
                    <h3 className="font-sans font-bold text-base text-[#007a54] uppercase tracking-wide border-b-2 border-[#007a54] pb-1.5 mb-4 mt-8">
                      PROGRAM SEMESTER (PROSEM)
                    </h3>

                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse border border-slate-300 text-[11px] text-left leading-normal">
                        <thead>
                          <tr className="bg-[#007a54] text-white">
                            <th rowSpan={2} className="border border-slate-300 p-2 text-center w-[5%] font-sans font-bold">No</th>
                            <th rowSpan={2} className="border border-slate-300 p-2 w-[40%] font-sans font-bold">Tujuan Pembelajaran</th>
                            <th rowSpan={2} className="border border-slate-300 p-2 text-center w-[8%] font-sans font-bold">Alokasi (JP)</th>
                            <th colSpan={6} className="border border-slate-300 p-2 text-center font-sans font-bold">Semester 1 (Ganjil)</th>
                            <th colSpan={6} className="border border-slate-300 p-2 text-center font-sans font-bold">Semester 2 (Genap)</th>
                          </tr>
                          <tr className="bg-[#007a54] text-white">
                            {["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => (
                              <th key={m} className="border border-slate-300 p-1 text-center font-sans font-semibold text-[10px] min-w-[28px]">{m}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {protaPromesData.promes.map((item, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/50">
                              <td className="border border-slate-300 p-2 text-center font-bold text-slate-700">{item.no || (idx + 1)}</td>
                              <td className="border border-slate-300 p-2 font-sans text-xs text-slate-800 leading-normal">{item.tujuanPembelajaran}</td>
                              <td className="border border-slate-300 p-2 text-center font-semibold font-sans text-xs">{item.alokasiWaktu}</td>
                              {["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => {
                                const val = item.distribusi[m] || "-";
                                const isNum = !isNaN(parseInt(val));
                                const bgClass = isNum 
                                  ? "bg-sky-50 text-sky-800 font-bold" 
                                  : val === "ASG" || val === "ASE" 
                                    ? "bg-slate-100 text-slate-600 font-bold" 
                                    : "text-slate-400";
                                return (
                                  <td key={m} className={`border border-slate-300 p-1 text-center text-[10px] font-sans ${bgClass}`}>
                                    {val}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                          <tr className="bg-slate-100 font-bold text-slate-800 border-t border-slate-300">
                            <td colSpan={2} className="border border-slate-300 p-2 text-right uppercase tracking-wide font-sans text-xs">
                              Jumlah JP per Semester
                            </td>
                            <td className="border border-slate-300 p-2 text-center font-sans text-xs">
                              {protaPromesData.promes.reduce((sum, item) => sum + (Number(item.alokasiWaktu) || 0), 0)}
                            </td>
                            <td colSpan={6} className="border border-slate-300 p-2 text-center font-sans text-xs text-emerald-800 bg-emerald-50/40">
                              Semester Ganjil: {protaPromesData.jumlahJpSemester1 || 172} JP
                            </td>
                            <td colSpan={6} className="border border-slate-300 p-2 text-center font-sans text-xs text-emerald-800 bg-emerald-50/40">
                              Semester Genap: {protaPromesData.jumlahJpSemester2 || 188} JP
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <p className="text-[10px] text-slate-500 mt-3 italic font-sans leading-relaxed">
                      *Keterangan: ASG = Asesmen Sumatif Ganjil / Cadangan; ASE = Asesmen Sumatif Akhir Tahun / Kelulusan Fase. Angka menunjukkan distribusi alokasi waktu JP per bulan efektif.
                    </p>

                    {/* Side-by-side signature block */}
                    <div className="mt-12 flex justify-between text-xs font-sans text-slate-800 leading-normal" style={{ pageBreakInside: "avoid" }}>
                      <div>
                        <p>Mengetahui,</p>
                        <p className="font-bold">Kepala Sekolah</p>
                        <div className="h-20" />
                        <p className="font-bold underline">({metadata.namaKepalaSekolah || "__________________________"})</p>
                        <p>NIP. {metadata.nipKepalaSekolah || "__________________________"}</p>
                      </div>
                      <div className="text-right sm:text-left">
                        <p>......................, 29 Juni 2026</p>
                        <p className="font-bold">Guru Mata Pelajaran</p>
                        <div className="h-20" />
                        <p className="font-bold underline">{metadata.namaGuru || "__________________________"}</p>
                        <p>NIP. {metadata.nipGuru || "__________________________"}</p>
                      </div>
                    </div>

                  </div>
                ) : (
                  /* EDIT MODE */
                  <div className="p-6 flex flex-col gap-6">
                    {/* PROTA Header */}
                    <div className="border-b border-slate-200 pb-2 mb-2 flex items-center justify-between">
                      <h3 className="font-bold text-slate-700 text-sm">EDIT PROGRAM TAHUNAN (PROTA)</h3>
                      <button
                        onClick={() => {
                          const updated = { ...protaPromesData };
                          updated.prota.push({
                            no: updated.prota.length + 1,
                            elemen: "Elemen Baru...",
                            tujuanPembelajaran: "Tujuan Pembelajaran...",
                            alokasiWaktu: "24 JP"
                          });
                          setProtaPromesData(updated);
                        }}
                        className="text-xs bg-emerald-100 text-emerald-800 py-1 px-2.5 rounded font-bold hover:bg-emerald-200 flex items-center gap-1"
                      >
                        <Plus className="w-3.5 h-3.5" /> Tambah Baris Prota
                      </button>
                    </div>

                    <div className="flex flex-col gap-3.5">
                      {protaPromesData.prota.map((item, idx) => (
                        <div key={idx} className="bg-slate-50 p-4 rounded-lg border border-slate-200 flex flex-col gap-3 relative shadow-sm">
                          <button
                            onClick={() => {
                              const updated = { ...protaPromesData };
                              updated.prota = protaPromesData.prota.filter((_, i) => i !== idx);
                              // re-number
                              updated.prota.forEach((p, i) => { p.no = i + 1; });
                              setProtaPromesData(updated);
                            }}
                            className="absolute top-2.5 right-2.5 text-red-500 hover:text-red-700 p-1 hover:bg-red-50 rounded"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>

                          <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                            <div className="flex flex-col gap-1 w-16">
                              <span className="text-[10px] font-bold text-slate-400">NO</span>
                              <input 
                                type="number" 
                                value={item.no || (idx + 1)}
                                onChange={(e) => {
                                  const updated = { ...protaPromesData };
                                  updated.prota[idx].no = parseInt(e.target.value) || (idx + 1);
                                  setProtaPromesData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-sm font-semibold text-center"
                              />
                            </div>

                            <div className="flex flex-col gap-1 sm:col-span-2">
                              <span className="text-[10px] font-bold text-slate-400">ELEMEN / LINGKUP MATERI</span>
                              <input 
                                type="text" 
                                value={item.elemen}
                                onChange={(e) => {
                                  const updated = { ...protaPromesData };
                                  updated.prota[idx].elemen = e.target.value;
                                  setProtaPromesData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-sm font-semibold"
                              />
                            </div>

                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-400">ALOKASI WAKTU</span>
                              <input 
                                type="text" 
                                value={item.alokasiWaktu}
                                onChange={(e) => {
                                  const updated = { ...protaPromesData };
                                  updated.prota[idx].alokasiWaktu = e.target.value;
                                  setProtaPromesData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-sm font-bold text-slate-600"
                              />
                            </div>
                          </div>

                          <div className="flex flex-col gap-1">
                            <span className="text-[10px] font-bold text-slate-400">TUJUAN PEMBELAJARAN (TP)</span>
                            <textarea 
                              rows={2}
                              value={item.tujuanPembelajaran}
                              onChange={(e) => {
                                const updated = { ...protaPromesData };
                                updated.prota[idx].tujuanPembelajaran = e.target.value;
                                setProtaPromesData(updated);
                              }}
                              className="bg-white border border-slate-300 rounded px-2 py-1 text-xs leading-relaxed"
                            />
                          </div>
                        </div>
                      ))}
                    </div>

                    {/* PROMES Header */}
                    <div className="border-b border-slate-200 pb-2 mt-6 mb-2 flex items-center justify-between">
                      <h3 className="font-bold text-slate-700 text-sm">EDIT PROGRAM SEMESTER (PROSEM)</h3>
                      <button
                        onClick={() => {
                          const updated = { ...protaPromesData };
                          updated.promes.push({
                            no: updated.promes.length + 1,
                            tujuanPembelajaran: "Tujuan Pembelajaran...",
                            alokasiWaktu: 24,
                            distribusi: {
                              Jul: "-", Agt: "-", Sep: "-", Okt: "-", Nov: "-", Des: "-",
                              Jan: "-", Feb: "-", Mar: "-", Apr: "-", Mei: "-", Jun: "-"
                            }
                          });
                          setProtaPromesData(updated);
                        }}
                        className="text-xs bg-emerald-100 text-emerald-800 py-1 px-2.5 rounded font-bold hover:bg-emerald-200 flex items-center gap-1"
                      >
                        <Plus className="w-3.5 h-3.5" /> Tambah Baris Promes
                      </button>
                    </div>

                    <div className="flex flex-col gap-5">
                      {protaPromesData.promes.map((item, idx) => (
                        <div key={idx} className="bg-white p-4 rounded-lg border border-slate-200 flex flex-col gap-3 relative shadow-sm">
                          <button
                            onClick={() => {
                              const updated = { ...protaPromesData };
                              updated.promes = protaPromesData.promes.filter((_, i) => i !== idx);
                              updated.promes.forEach((p, i) => { p.no = i + 1; });
                              setProtaPromesData(updated);
                            }}
                            className="absolute top-2.5 right-2.5 text-red-500 hover:text-red-700 p-1 hover:bg-red-50 rounded"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>

                          <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                            <div className="flex flex-col gap-1 w-16">
                              <span className="text-[10px] font-bold text-slate-400">NO</span>
                              <input 
                                type="number" 
                                value={item.no || (idx + 1)}
                                onChange={(e) => {
                                  const updated = { ...protaPromesData };
                                  updated.promes[idx].no = parseInt(e.target.value) || (idx + 1);
                                  setProtaPromesData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-sm font-semibold text-center"
                              />
                            </div>

                            <div className="flex flex-col gap-1 sm:col-span-2">
                              <span className="text-[10px] font-bold text-slate-400">TUJUAN PEMBELAJARAN (TP)</span>
                              <textarea 
                                rows={2}
                                value={item.tujuanPembelajaran}
                                onChange={(e) => {
                                  const updated = { ...protaPromesData };
                                  updated.promes[idx].tujuanPembelajaran = e.target.value;
                                  setProtaPromesData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2 py-1 text-xs leading-normal"
                              />
                            </div>

                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-400">ALOKASI WAKTU (JP)</span>
                              <input 
                                type="number" 
                                value={item.alokasiWaktu}
                                onChange={(e) => {
                                  const updated = { ...protaPromesData };
                                  updated.promes[idx].alokasiWaktu = parseInt(e.target.value) || 0;
                                  setProtaPromesData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-sm font-bold text-slate-600"
                              />
                            </div>
                          </div>

                          <div className="border-t border-slate-100 pt-2">
                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2 block">Distribusi Pembelajaran Bulanan</span>
                            <div className="grid grid-cols-4 sm:grid-cols-6 lg:grid-cols-12 gap-2">
                              {["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => (
                                <div key={m} className="flex flex-col gap-1">
                                  <label className="text-[9px] font-bold text-slate-500 text-center uppercase bg-slate-50 py-0.5 rounded border border-slate-100">{m}</label>
                                  <input
                                    type="text"
                                    value={item.distribusi[m] || "-"}
                                    onChange={(e) => {
                                      const updated = { ...protaPromesData };
                                      updated.promes[idx].distribusi[m] = e.target.value;
                                      setProtaPromesData(updated);
                                    }}
                                    className="w-full text-center bg-white border border-slate-200 rounded py-1 px-0.5 text-xs font-semibold focus:ring-1 focus:ring-emerald-500"
                                  />
                                </div>
                              ))}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>

                    <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4">
                      <div className="flex flex-col gap-1">
                        <span className="text-[10px] font-bold text-slate-400">TOTAL JP SEMESTER GANJIL</span>
                        <input 
                          type="number" 
                          value={protaPromesData.jumlahJpSemester1}
                          onChange={(e) => {
                            const updated = { ...protaPromesData };
                            updated.jumlahJpSemester1 = parseInt(e.target.value) || 0;
                            setProtaPromesData(updated);
                          }}
                          className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-sm font-bold text-slate-700"
                        />
                      </div>
                      <div className="flex flex-col gap-1">
                        <span className="text-[10px] font-bold text-slate-400">TOTAL JP SEMESTER GENAP</span>
                        <input 
                          type="number" 
                          value={protaPromesData.jumlahJpSemester2}
                          onChange={(e) => {
                            const updated = { ...protaPromesData };
                            updated.jumlahJpSemester2 = parseInt(e.target.value) || 0;
                            setProtaPromesData(updated);
                          }}
                          className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-sm font-bold text-slate-700"
                        />
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(2)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => generateProtaPromes(true)}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>
                  <button
                    onClick={() => {
                      if (kktpData) setStep(4);
                      else generateKKTP();
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial hover:shadow-md"
                  >
                    <span>Buat KKTP</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 4: KKTP */}
          {step === 4 && kktpData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              <div className="bg-white border border-slate-200 rounded-xl shadow-md min-h-[500px] flex flex-col">
                
                {/* PREVIEW MODE */}
                {!editMode ? (
                  <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-sans leading-relaxed text-slate-900 A4-paper gap-6 w-full">
                    {/* Header Banner */}
                    <div className="bg-[#007A54] text-white p-6 rounded-lg text-center shadow-sm w-full">
                      <h1 className="font-bold text-lg sm:text-2xl uppercase tracking-wider">ANALISIS KRITERIA KETERCAPAIAN (KKTP)</h1>
                      <div className="w-16 h-1 bg-white mx-auto my-3 rounded"></div>
                      <p className="text-xs sm:text-sm font-medium tracking-wide">
                        Mata Pelajaran: {metadata.mataPelajaran} | Fase/Kelas: {metadata.fase} / XI
                      </p>
                    </div>

                    {/* Identity Table */}
                    <div className="overflow-hidden border border-emerald-100 rounded-lg shadow-sm w-full">
                      <table className="w-full text-xs text-left border-collapse">
                        <tbody>
                          <tr className="border-b border-emerald-100">
                            <td className="w-1/3 bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Nama Guru / Penyusun</td>
                            <td className="p-2.5 text-slate-700">{metadata.namaGuru}</td>
                          </tr>
                          <tr className="border-b border-emerald-100">
                            <td className="bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Instansi / Sekolah</td>
                            <td className="p-2.5 text-slate-700">{metadata.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan"}</td>
                          </tr>
                          <tr className="border-b border-emerald-100">
                            <td className="bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Mata Pelajaran</td>
                            <td className="p-2.5 text-slate-700">{metadata.mataPelajaran}</td>
                          </tr>
                          <tr>
                            <td className="bg-emerald-50/50 p-2.5 font-bold text-emerald-800">Jenjang / Fase / Kelas</td>
                            <td className="p-2.5 text-slate-700">SMK/MAK / {metadata.fase} / XI</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Pendekatan KKTP Section */}
                    <div className="w-full">
                      <h2 className="text-sm sm:text-base font-bold text-[#007A54] border-b-2 border-[#007A54] pb-1 mb-3 uppercase tracking-wide">
                        PENDEKATAN KKTP
                      </h2>
                      <div className="overflow-x-auto border border-slate-200 rounded-lg shadow-sm w-full">
                        <table className="w-full text-xs text-left border-collapse">
                          <thead>
                            <tr className="bg-[#007A54] text-white">
                              <th className="p-3 font-semibold w-[30%] border-r border-emerald-650">Metode Pendekatan</th>
                              <th className="p-3 font-semibold w-[70%]">Deskripsi dan Alasan Penggunaan</th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr className="border-b border-slate-200">
                              <td className="p-3 font-bold text-slate-800 align-top border-r border-slate-200 bg-slate-50/50">Pendekatan Interval Nilai</td>
                              <td className="p-3 text-slate-700 leading-relaxed text-justify">
                                Menggunakan interval kriteria untuk menentukan ketuntasan belajar peserta didik. Pendekatan ini dipilih karena memberikan gambaran objektif, terukur, dan mempermudah pemetaan tingkat kompetensi peserta didik (Mulai Berkembang, Layak, Cakap, Mahir) pada setiap indikator mata pelajaran {metadata.mataPelajaran} Fase {metadata.fase} (Kelas XI).
                              </td>
                            </tr>
                            <tr>
                              <td className="p-3 font-bold text-slate-800 align-top border-r border-slate-200 bg-slate-50/50">Ketentuan Ketuntasan</td>
                              <td className="p-3 text-slate-700 leading-relaxed text-justify">
                                Peserta didik dinyatakan mencapai ketuntasan minimum jika mencapai kriteria minimal <span className="font-semibold text-[#007A54]">Layak (61-70)</span> pada indikator-indikator kunci kompetensi esensial.
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Analisis KKTP Section */}
                    <div className="w-full">
                      <h2 className="text-sm sm:text-base font-bold text-[#007A54] border-b-2 border-[#007A54] pb-1 mb-3 uppercase tracking-wide">
                        ANALISIS KKTP
                      </h2>
                      <div className="overflow-x-auto border border-slate-200 rounded-lg shadow-sm w-full">
                        <table className="w-full text-[10px] sm:text-xs text-left border-collapse">
                          <thead>
                            <tr className="bg-[#007A54] text-white text-center">
                              <th className="p-2 border border-slate-300 font-semibold w-[4%]" rowSpan={2}>No</th>
                              <th className="p-2 border border-slate-300 font-semibold w-[18%]" rowSpan={2}>Tujuan Pembelajaran (TP)</th>
                              <th className="p-2 border border-slate-300 font-semibold w-[22%]" rowSpan={2}>Indikator Ketercapaian (IKTP)</th>
                              <th className="p-2 border border-slate-300 font-semibold w-[56%]" colSpan={4}>Kriteria Interval Nilai (%)</th>
                            </tr>
                            <tr className="bg-[#007A54] text-white text-center text-[9px] sm:text-[10px]">
                              <th className="p-2 border border-slate-300 font-semibold w-[14%]">Mulai Berkembang (0-60)</th>
                              <th className="p-2 border border-slate-300 font-semibold w-[14%]">Layak (61-70)</th>
                              <th className="p-2 border border-slate-300 font-semibold w-[14%]">Cakap (71-80)</th>
                              <th className="p-2 border border-slate-300 font-semibold w-[14%]">Mahir (81-100)</th>
                            </tr>
                          </thead>
                          <tbody>
                            {kktpData.kktp.map((item, idx) => {
                              const criteriaList = item.kriteria || [];
                              const rowSpanVal = criteriaList.length || 1;

                              if (criteriaList.length === 0) {
                                return (
                                  <tr key={`kktp-empty-${idx}`} className="border-b border-slate-200">
                                    <td className="p-2 border border-slate-200 text-center font-bold align-top">{idx + 1}</td>
                                    <td className="p-2 border border-slate-200 font-bold align-top">
                                      <span className="font-mono text-emerald-700 mr-1 text-[9px]">{item.tpKode}</span>
                                      {item.tpTeks}
                                    </td>
                                    <td className="p-2 border border-slate-200 align-top">-</td>
                                    <td className="p-2 border border-slate-200 align-top">-</td>
                                    <td className="p-2 border border-slate-200 align-top">-</td>
                                    <td className="p-2 border border-slate-200 align-top">-</td>
                                    <td className="p-2 border border-slate-200 align-top">-</td>
                                  </tr>
                                );
                              }

                              return criteriaList.map((crit, cIdx) => (
                                <tr key={`kktp-${idx}-${cIdx}`} className="border-b border-slate-200 hover:bg-slate-50/30">
                                  {cIdx === 0 && (
                                    <td className="p-2 border border-slate-200 text-center font-bold align-top" rowSpan={rowSpanVal}>
                                      {idx + 1}
                                    </td>
                                  )}
                                  {cIdx === 0 && (
                                    <td className="p-2 border border-slate-200 font-medium align-top leading-relaxed text-justify" rowSpan={rowSpanVal}>
                                      <div className="font-bold text-slate-800 text-[11px] sm:text-xs">
                                        <span className="font-bold font-mono text-emerald-700 bg-emerald-50 px-1 rounded text-[9px] mr-1">{item.tpKode}</span>
                                        {item.tpTeks}
                                      </div>
                                    </td>
                                  )}
                                  <td className="p-2 border border-slate-200 font-semibold text-[#007A54] align-top text-justify bg-emerald-50/5">
                                    {crit.aspek}
                                  </td>
                                  <td className="p-2 border border-slate-200 text-slate-600 align-top text-justify text-[10px] sm:text-[11px]">
                                    {crit.baruBerkembang || crit.mulaiBerkembang}
                                  </td>
                                  <td className="p-2 border border-slate-200 text-slate-600 align-top text-justify text-[10px] sm:text-[11px] bg-amber-50/10">
                                    {crit.layak}
                                  </td>
                                  <td className="p-2 border border-slate-200 text-slate-600 align-top text-justify text-[10px] sm:text-[11px] bg-emerald-50/10">
                                    {crit.cakap}
                                  </td>
                                  <td className="p-2 border border-slate-200 text-slate-800 font-medium align-top text-justify text-[10px] sm:text-[11px] bg-blue-50/10">
                                    {crit.mahir}
                                  </td>
                                </tr>
                              ));
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Tindak Lanjut Section */}
                    <div className="w-full">
                      <h2 className="text-sm sm:text-base font-bold text-[#007A54] border-b-2 border-[#007A54] pb-1 mb-3 uppercase tracking-wide">
                        TINDAK LANJUT
                      </h2>
                      <div className="overflow-x-auto border border-slate-200 rounded-lg shadow-sm w-full">
                        <table className="w-full text-[10px] sm:text-xs text-left border-collapse">
                          <thead>
                            <tr className="bg-[#007A54] text-white">
                              <th className="p-2.5 border border-emerald-600 font-semibold w-[25%]">Kategori Interval</th>
                              <th className="p-2.5 border border-emerald-600 font-semibold w-[20%]">Status Kelulusan</th>
                              <th className="p-2.5 border border-emerald-600 font-semibold w-[55%]">Tindak Lanjut / Intervensi</th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr className="border-b border-slate-200">
                              <td className="p-2.5 border border-slate-200 font-bold bg-slate-50/50">
                                0 - 60 %<br/><span className="font-normal text-slate-500 text-[10px]">(Mulai Berkembang)</span>
                              </td>
                              <td className="p-2.5 border border-slate-200 font-bold text-red-600">Belum Tuntas</td>
                              <td className="p-2.5 border border-slate-200 text-slate-700 leading-relaxed text-justify">
                                Peserta didik wajib mengikuti program bimbingan intensif/remedial individu atau kelompok kecil pada materi yang belum dikuasai (misalnya, simulasi routing/perhitungan subnetting ulang menggunakan alat peraga interaktif sebelum dilakukan asesmen ulang).
                              </td>
                            </tr>
                            <tr className="border-b border-slate-200">
                              <td className="p-2.5 border border-slate-200 font-bold bg-slate-50/50">
                                61 - 70 %<br/><span className="font-normal text-slate-500 text-[10px]">(Layak)</span>
                              </td>
                              <td className="p-2.5 border border-slate-200 font-bold text-amber-600">Tuntas (Syarat)</td>
                              <td className="p-2.5 border border-slate-200 text-slate-700 leading-relaxed text-justify">
                                Peserta didik diberikan pendampingan berupa latihan terbimbing tambahan (peer tutoring) atau tugas mandiri terfokus untuk memperkuat pemahaman operasional dasar dari kompetensi yang belum sepenuhnya lancar.
                              </td>
                            </tr>
                            <tr className="border-b border-slate-200">
                              <td className="p-2.5 border border-slate-200 font-bold bg-slate-50/50">
                                71 - 80 %<br/><span className="font-normal text-slate-500 text-[10px]">(Cakap)</span>
                              </td>
                              <td className="p-2.5 border border-slate-200 font-bold text-emerald-600">Tuntas</td>
                              <td className="p-2.5 border border-slate-200 text-slate-700 leading-relaxed text-justify">
                                Peserta didik langsung melanjutkan ke materi pembelajaran berikutnya dan diarahkan untuk mengeksplorasi studi kasus penyelesaian masalah jaringan riil dengan skala yang lebih luas.
                              </td>
                            </tr>
                            <tr>
                              <td className="p-2.5 border border-slate-200 font-bold bg-slate-50/50">
                                81 - 100 %<br/><span className="font-normal text-slate-500 text-[10px]">(Mahir)</span>
                              </td>
                              <td className="p-2.5 border border-slate-200 font-bold text-blue-600">Tuntas (Sangat Baik)</td>
                              <td className="p-2.5 border border-slate-200 text-slate-700 leading-relaxed text-justify">
                                Peserta didik diberikan program pengayaan berupa tantangan proyek desain infrastruktur jaringan industri skala enterprise, diajarkan teknologi advanced (seperti implementasi jaringan nirkabel cerdas/SDN), atau ditunjuk sebagai tutor sebaya (peer tutor) untuk membantu rekan sekelas yang berada di kategori mulai berkembang.
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Side-by-Side Signatures */}
                    <div className="mt-8 pt-6 border-t border-slate-100 grid grid-cols-2 text-xs w-full">
                      <div className="flex flex-col gap-1 text-left">
                        <p>Mengetahui,</p>
                        <p className="font-semibold text-slate-850">Kepala Sekolah</p>
                        <div className="h-16"></div>
                        <p className="font-bold text-slate-900 underline">
                          ( {metadata.namaKepalaSekolah || "__________________________"} )
                        </p>
                        <p className="text-slate-500">NIP. {metadata.nipKepalaSekolah || "..................................."}</p>
                      </div>
                      <div className="flex flex-col gap-1 text-left pl-8 border-l border-slate-100">
                        <p>......................, 29 Juni 2026</p>
                        <p className="font-semibold text-slate-850">Guru Mata Pelajaran</p>
                        <div className="h-16"></div>
                        <p className="font-bold text-slate-900">
                          {metadata.namaGuru || "__________________________"}
                        </p>
                        <p className="text-slate-500">NIP. {metadata.nipGuru || "..................................."}</p>
                      </div>
                    </div>
                  </div>
                ) : (
                  /* EDIT MODE */
                  <div className="p-6 flex flex-col gap-6">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">JUDUL KKTP</label>
                      <input 
                        type="text" 
                        value={kktpData.judul}
                        onChange={(e) => setKktpData({ ...kktpData, judul: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 font-bold"
                      />
                    </div>

                    <div className="flex flex-col gap-5">
                      {kktpData.kktp.map((item, idx) => (
                        <div key={idx} className="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col gap-3">
                          <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 border-b border-slate-200 pb-2.5">
                            <div className="flex flex-col gap-0.5 sm:col-span-1">
                              <span className="text-[10px] font-bold text-slate-400">KODE TP</span>
                              <input 
                                type="text" 
                                value={item.tpKode}
                                onChange={(e) => {
                                  const updated = { ...kktpData };
                                  updated.kktp[idx].tpKode = e.target.value;
                                  setKktpData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs font-mono font-bold text-emerald-800"
                              />
                            </div>
                            <div className="flex flex-col gap-0.5 sm:col-span-3">
                              <span className="text-[10px] font-bold text-slate-400">DESKRIPSI TUJUAN PEMBELAJARAN</span>
                              <input 
                                type="text" 
                                value={item.tpTeks}
                                onChange={(e) => {
                                  const updated = { ...kktpData };
                                  updated.kktp[idx].tpTeks = e.target.value;
                                  setKktpData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs font-semibold"
                              />
                            </div>
                          </div>

                          {/* Edit Kriteria/Aspek */}
                          <div className="flex flex-col gap-4">
                            <span className="text-[10px] font-bold text-slate-400">RUBRIK KRITERIA KOMPETENSI:</span>
                            {item.kriteria.map((crit, cIdx) => (
                              <div key={cIdx} className="bg-white p-3 rounded border border-slate-200 flex flex-col gap-2 shadow-sm">
                                <div className="flex flex-col gap-0.5">
                                  <span className="text-[9px] font-bold text-slate-400">NAMA ASPEK</span>
                                  <input 
                                    type="text" 
                                    value={crit.aspek}
                                    onChange={(e) => {
                                      const updated = { ...kktpData };
                                      updated.kktp[idx].kriteria[cIdx].aspek = e.target.value;
                                      setKktpData(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-bold text-slate-700"
                                  />
                                </div>

                                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-2">
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] font-bold text-slate-400">MULAI BERKEMBANG (0-60)</span>
                                      <textarea 
                                        rows={2}
                                        value={crit.baruBerkembang}
                                        onChange={(e) => {
                                          const updated = { ...kktpData };
                                          updated.kktp[idx].kriteria[cIdx].baruBerkembang = e.target.value;
                                          setKktpData(updated);
                                        }}
                                        className="bg-white border border-slate-300 rounded px-2 py-0.5 text-[11px]"
                                      />
                                    </div>
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] font-bold text-slate-400">LAYAK (61-70)</span>
                                      <textarea 
                                        rows={2}
                                        value={crit.layak}
                                        onChange={(e) => {
                                          const updated = { ...kktpData };
                                          updated.kktp[idx].kriteria[cIdx].layak = e.target.value;
                                          setKktpData(updated);
                                        }}
                                        className="bg-white border border-slate-300 rounded px-2 py-0.5 text-[11px]"
                                      />
                                    </div>
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] font-bold text-slate-400">CAKAP (71-80)</span>
                                      <textarea 
                                        rows={2}
                                        value={crit.cakap}
                                        onChange={(e) => {
                                          const updated = { ...kktpData };
                                          updated.kktp[idx].kriteria[cIdx].cakap = e.target.value;
                                          setKktpData(updated);
                                        }}
                                        className="bg-white border border-slate-300 rounded px-2 py-0.5 text-[11px]"
                                      />
                                    </div>
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] font-bold text-slate-400">MAHIR (81-100)</span>
                                      <textarea 
                                        rows={2}
                                        value={crit.mahir}
                                        onChange={(e) => {
                                          const updated = { ...kktpData };
                                          updated.kktp[idx].kriteria[cIdx].mahir = e.target.value;
                                          setKktpData(updated);
                                        }}
                                        className="bg-white border border-slate-300 rounded px-2 py-0.5 text-[11px]"
                                      />
                                    </div>
                                  </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(3)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => generateKKTP(true)}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>
                  <button
                    onClick={() => {
                      if (topikData) setStep(5);
                      else generateTopik();
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 sm:flex-initial hover:shadow-md"
                  >
                    <span>Buat Daftar Topik</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 5: DAFTAR TOPIK PEMBELAJARAN */}
          {step === 5 && topikData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              <div className="bg-white border border-slate-200 rounded-xl shadow-md min-h-[500px] flex flex-col">
                
                {/* PREVIEW MODE */}
                {!editMode ? (
                  <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-serif leading-relaxed text-sm sm:text-base text-slate-900 A4-paper w-full">
                    <div className="text-center border-b border-double border-slate-300 pb-4 mb-6">
                      <h1 className="font-bold text-lg sm:text-xl uppercase tracking-wide">Daftar Topik / Materi Utama Pembelajaran</h1>
                      <h2 className="font-bold text-base mt-1 uppercase tracking-wide">Mata Pelajaran: {metadata.mataPelajaran} | Fase {metadata.fase}</h2>
                      <p className="text-xs text-slate-500 mt-2 font-sans italic">{topikData.judul}</p>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                        <thead>
                          <tr className="bg-slate-50">
                            <th className="border border-slate-300 p-2.5 font-sans font-bold w-[10%] text-center">No</th>
                            <th className="border border-slate-300 p-2.5 font-sans font-bold w-[30%]">Materi Pokok / Topik</th>
                            <th className="border border-slate-300 p-2.5 font-sans font-bold w-[60%]">Tujuan Pembelajaran Terkait</th>
                          </tr>
                        </thead>
                        <tbody>
                          {topikData.daftarTopik.map((t, idx) => (
                            <tr key={idx}>
                              <td className="border border-slate-300 p-2.5 text-center font-bold text-slate-700 bg-slate-50/50">{t.no}</td>
                              <td className="border border-slate-300 p-2.5 font-sans font-semibold">{t.topik}</td>
                              <td className="border border-slate-300 p-2.5 text-justify leading-relaxed">{t.tpTeks}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : (
                  /* EDIT MODE */
                  <div className="p-6 flex flex-col gap-5">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs font-bold text-slate-500">JUDUL DOKUMEN</label>
                      <input 
                        type="text" 
                        value={topikData.judul}
                        onChange={(e) => setTopikData({ ...topikData, judul: e.target.value })}
                        className="bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 font-bold"
                      />
                    </div>

                    <div className="flex flex-col gap-4">
                      {topikData.daftarTopik.map((t, idx) => (
                        <div key={idx} className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-400">MATERI POKOK / TOPIK</span>
                              <input 
                                type="text" 
                                value={t.topik}
                                onChange={(e) => {
                                  const updated = { ...topikData };
                                  updated.daftarTopik[idx].topik = e.target.value;
                                  setTopikData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs font-semibold"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-400">TUJUAN PEMBELAJARAN</span>
                              <textarea 
                                rows={2}
                                value={t.tpTeks}
                                onChange={(e) => {
                                  const updated = { ...topikData };
                                  updated.daftarTopik[idx].tpTeks = e.target.value;
                                  setTopikData(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                              />
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(4)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center flex-wrap gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => generateTopik(true)}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>

                  <button
                    onClick={() => setStep(6)}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 hover:shadow-md"
                  >
                    <span>Buat Rencana Pembelajaran Mendalam</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 6: RENCANA PEMBELAJARAN MENDALAM (RPM) */}
          {step === 6 && topikData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              {/* Batch Action Bar */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                <div className="text-xs sm:text-sm">
                  <span className="font-bold text-slate-800">Rencana Pembelajaran Mendalam (RPM)</span>
                  <p className="text-slate-500 text-[10px] sm:text-xs mt-0.5">Pembuatan modul ajar komprehensif untuk setiap materi pokok.</p>
                </div>
                <button
                  onClick={generateAllRpm}
                  className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded text-xs transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Generate Semua RPM Otomatis</span>
                </button>
              </div>

              <div className="flex flex-col lg:flex-row gap-6 items-stretch">
                {/* Topic selector panel */}
                <div className="w-full lg:w-72 bg-white rounded-xl border border-slate-200 p-4 flex flex-col gap-3 shrink-0">
                  <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Materi Pokok</h3>
                  <div className="flex flex-col gap-2 max-h-[300px] lg:max-h-none overflow-y-auto">
                    {topikData.daftarTopik.map((t, idx) => {
                      const isSelected = selectedRpmTopic === t.topik;
                      const hasData = !!rpmList[t.topik];
                      return (
                        <button
                          key={idx}
                          onClick={() => setSelectedRpmTopic(t.topik)}
                          className={`w-full text-left p-3 rounded-lg border transition-all text-xs flex flex-col gap-1.5 cursor-pointer ${
                            isSelected 
                              ? "bg-slate-900 text-white border-slate-950 shadow" 
                              : "bg-white hover:bg-slate-50 text-slate-700 border-slate-200"
                          }`}
                        >
                          <div className="flex items-center justify-between gap-1 w-full">
                            <span className="font-mono font-bold">Topik {t.no}</span>
                            <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                              hasData 
                                ? "bg-green-100 text-green-800 border border-green-200" 
                                : "bg-slate-100 text-slate-500 border border-slate-200"
                            }`}>
                              {hasData ? "Sudah Dibuat" : "Belum Dibuat"}
                            </span>
                          </div>
                          <p className={`font-semibold line-clamp-2 leading-tight ${isSelected ? "text-slate-100" : "text-slate-800"}`}>
                            {t.topik}
                          </p>
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* Main Preview and Editor Panel */}
                <div className="flex-1 bg-white border border-slate-200 rounded-xl shadow-md flex flex-col min-h-[450px]">
                  {rpmList[selectedRpmTopic] ? (
                    <div>
                      {!editMode ? (
                        /* Preview Paper */
                        <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-sans leading-relaxed text-sm sm:text-base text-slate-900 A4-paper w-full">
                          {/* Top Header Banner matching the PDF screenshots exactly */}
                          <div className="bg-[#007C58] text-white p-6 rounded-md mb-6 text-center shadow-sm">
                            <h1 className="font-bold text-lg sm:text-xl uppercase tracking-wide">MODUL AJAR PEMBELAJARAN MENDALAM</h1>
                            <p className="text-xs sm:text-sm mt-1 opacity-90 font-medium">
                              Mata Pelajaran: {rpmList[selectedRpmTopic].identitas?.mataPelajaran || metadata.mataPelajaran} | 
                              Fase: {metadata.fase.toUpperCase()} | 
                              Kelas/Semester: {rpmList[selectedRpmTopic].identitas?.kelasSemester || "-"}
                            </p>
                          </div>

                          <div className="flex flex-col gap-6 flex-1 text-slate-800 text-xs sm:text-sm">
                            {/* I. INFORMASI ADMINISTRATIF */}
                            <div>
                              <h3 className="font-bold text-sm sm:text-base border-b border-slate-200 pb-1 mb-2 text-[#007C58] uppercase tracking-tight">I. Informasi Administratif</h3>
                              <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                <table className="w-full border-collapse text-left">
                                  <tbody>
                                    <tr className="border-b border-slate-100">
                                      <td className="p-2.5 font-bold w-[30%] bg-slate-50 border-r border-slate-150">Nama Penyusun / Guru</td>
                                      <td className="p-2.5">{rpmList[selectedRpmTopic].identitas?.namaGuru || metadata.namaGuru}</td>
                                    </tr>
                                    <tr className="border-b border-slate-100">
                                      <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Institusi / Sekolah</td>
                                      <td className="p-2.5">{rpmList[selectedRpmTopic].identitas?.sekolah || metadata.jenjangSekolah}</td>
                                    </tr>
                                    <tr className="border-b border-slate-100">
                                      <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Mata Pelajaran</td>
                                      <td className="p-2.5">{rpmList[selectedRpmTopic].identitas?.mataPelajaran || metadata.mataPelajaran}</td>
                                    </tr>
                                    <tr className="border-b border-slate-100">
                                      <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Fase / Kelas</td>
                                      <td className="p-2.5">Fase {metadata.fase.toUpperCase()} / {rpmList[selectedRpmTopic].identitas?.kelasSemester || "-"}</td>
                                    </tr>
                                    <tr className="border-b border-slate-100">
                                      <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Semester & Tahun Ajaran</td>
                                      <td className="p-2.5">{rpmList[selectedRpmTopic].identitas?.semesterTahun || `${rpmList[selectedRpmTopic].identitas?.kelasSemester?.includes("Semester") ? "" : "Semester Ganjil"} - ${metadata.tahunAjaran}`}</td>
                                    </tr>
                                    <tr className="border-b border-slate-100">
                                      <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Topik / Materi Pokok</td>
                                      <td className="p-2.5">{rpmList[selectedRpmTopic].identitas?.topik || selectedRpmTopic}</td>
                                    </tr>
                                    <tr>
                                      <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Model Pembelajaran</td>
                                      <td className="p-2.5">{rpmList[selectedRpmTopic].identitas?.modelPembelajaran || "Tatap Muka / Reguler"}</td>
                                    </tr>
                                  </tbody>
                                </table>
                              </div>
                            </div>

                            {/* II. DESAIN PENGEMBANGAN DAN IMPLEMENTASI */}
                            <div>
                              <h3 className="font-bold text-sm sm:text-base border-b border-slate-200 pb-1 mb-2 text-[#007C58] uppercase tracking-tight">II. Desain Pengembangan & Implementasi</h3>
                              
                              {/* A. IDENTIFIKASI TABLE */}
                              <div className="mb-5">
                                <h4 className="font-bold text-xs sm:text-sm text-slate-700 mb-1.5 uppercase">A. IDENTIFIKASI</h4>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full border-collapse text-left">
                                    <thead>
                                      <tr className="bg-[#007C58] text-white">
                                        <th className="p-2.5 font-bold w-[30%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Aspek Identifikasi</th>
                                        <th className="p-2.5 font-bold w-[70%] text-[11px] sm:text-xs uppercase">Deskripsi Analisis Pengembangan</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Target Peserta Didik</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].identifikasi?.targetPesertaDidik || "-"}</td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Kompetensi Prasyarat</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5 flex flex-col gap-0.5">
                                            {(rpmList[selectedRpmTopic].identifikasi?.kompetensiPrasyarat || []).map((p, idx) => (
                                              <li key={idx}>{p}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Profil Lulusan (8 Dimensi)</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5 flex flex-col gap-0.5">
                                            {(rpmList[selectedRpmTopic].identifikasi?.profilLulusan || []).map((p, idx) => (
                                              <li key={idx}>{p}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                      <tr>
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Sarana dan Prasarana</td>
                                        <td className="p-2.5 flex flex-col gap-1">
                                          <div><strong>• Fasilitas Fisik:</strong> {rpmList[selectedRpmTopic].identifikasi?.saranaPrasarana?.fasilitasFisik || "-"}</div>
                                          <div><strong>• Piranti Lunak:</strong> {rpmList[selectedRpmTopic].identifikasi?.saranaPrasarana?.pirantiLunak || "-"}</div>
                                          <div><strong>• Sumber Belajar:</strong> {rpmList[selectedRpmTopic].identifikasi?.saranaPrasarana?.sumberBelajar || "-"}</div>
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              {/* B. DESAIN PEMBELAJARAN TABLE */}
                              <div className="mb-5">
                                <h4 className="font-bold text-xs sm:text-sm text-slate-700 mb-1.5 uppercase">B. DESAIN PEMBELAJARAN</h4>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full border-collapse text-left">
                                    <thead>
                                      <tr className="bg-[#007C58] text-white">
                                        <th className="p-2.5 font-bold w-[30%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Aspek Desain</th>
                                        <th className="p-2.5 font-bold w-[70%] text-[11px] sm:text-xs uppercase">Rincian Implementasi Kurikulum Merdeka</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Capaian Pembelajaran</td>
                                        <td className="p-2.5 text-justify">{rpmList[selectedRpmTopic].desainPembelajaran?.capaianPembelajaran || "-"}</td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Lintas Disiplin Ilmu</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5 flex flex-col gap-0.5">
                                            {(rpmList[selectedRpmTopic].desainPembelajaran?.lintasDisiplin || []).map((p, idx) => (
                                              <li key={idx}>{p}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Tujuan Pembelajaran</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5 flex flex-col gap-0.5">
                                            {(rpmList[selectedRpmTopic].desainPembelajaran?.tujuanPembelajaran || []).map((p, idx) => (
                                              <li key={idx}>{p}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Topik Pembelajaran</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].desainPembelajaran?.topikPembelajaran || "-"}</td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Praktik Pedagogis (Deep Learning)</td>
                                        <td className="p-2.5 text-justify">{rpmList[selectedRpmTopic].desainPembelajaran?.praktikPedagogis || "-"}</td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Kemitraan Pembelajaran</td>
                                        <td className="p-2.5 text-justify">{rpmList[selectedRpmTopic].desainPembelajaran?.kemitraanPembelajaran || "-"}</td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Lingkungan Belajar</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].desainPembelajaran?.lingkunganBelajar || "-"}</td>
                                      </tr>
                                      <tr>
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Pemanfaatan Digital</td>
                                        <td className="p-2.5 text-justify">{rpmList[selectedRpmTopic].desainPembelajaran?.pemanfaatanDigital || "-"}</td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              {/* C. LANGKAH PEMBELAJARAN TABLE */}
                              <div className="mb-5">
                                <h4 className="font-bold text-xs sm:text-sm text-slate-700 mb-1.5 uppercase">C. LANGKAH PEMBELAJARAN</h4>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full border-collapse text-left">
                                    <thead>
                                      <tr className="bg-[#007C58] text-white">
                                        <th className="p-2.5 font-bold w-[20%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Fase Kegiatan</th>
                                        <th className="p-2.5 font-bold w-[65%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Aktivitas Pendidik & Peserta Didik</th>
                                        <th className="p-2.5 font-bold w-[15%] text-[11px] sm:text-xs uppercase">Pendekatan</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {/* Pendahuluan */}
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">
                                          Pendahuluan
                                          <div className="text-[10px] text-slate-500 font-normal">({rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.durasi || "15 Menit"})</div>
                                        </td>
                                        <td className="p-2.5 flex flex-col gap-2">
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Guru (Pendidik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.guru || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Murid (Peserta Didik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.siswa || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                        </td>
                                        <td className="p-2.5 font-bold text-[#007C58] text-center border-l border-slate-100">
                                          {rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.pendekatan || "MINDFUL"}
                                        </td>
                                      </tr>

                                      {/* Eksplorasi Konsep */}
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">
                                          Eksplorasi Konsep & Analisis
                                          <div className="text-[10px] text-slate-500 font-normal">({rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.durasi || "45 Menit"})</div>
                                        </td>
                                        <td className="p-2.5 flex flex-col gap-2">
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Guru (Pendidik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.guru || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Murid (Peserta Didik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.siswa || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                        </td>
                                        <td className="p-2.5 font-bold text-[#007C58] text-center border-l border-slate-100">
                                          {rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.pendekatan || "MEANINGFUL"}
                                        </td>
                                      </tr>

                                      {/* Perancangan & Simulasi */}
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">
                                          Perancangan & Simulasi
                                          <div className="text-[10px] text-slate-500 font-normal">({rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.durasi || "45 Menit"})</div>
                                        </td>
                                        <td className="p-2.5 flex flex-col gap-2">
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Guru (Pendidik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.guru || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Murid (Peserta Didik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.siswa || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                        </td>
                                        <td className="p-2.5 font-bold text-[#007C58] text-center border-l border-slate-100">
                                          {rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.pendekatan || "JOYFUL"}
                                        </td>
                                      </tr>

                                      {/* Penutup */}
                                      <tr>
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">
                                          Penutup & Evaluasi
                                          <div className="text-[10px] text-slate-500 font-normal">({rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.durasi || "15 Menit"})</div>
                                        </td>
                                        <td className="p-2.5 flex flex-col gap-2">
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Guru (Pendidik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.guru || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                          <div>
                                            <strong className="text-[10px] text-emerald-700 uppercase tracking-wide block mb-0.5">Aktivitas Murid (Peserta Didik)</strong>
                                            <ul className="list-disc pl-4 text-xs">
                                              {(rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.siswa || []).map((p, idx) => (
                                                <li key={idx}>{p}</li>
                                              ))}
                                            </ul>
                                          </div>
                                        </td>
                                        <td className="p-2.5 font-bold text-[#007C58] text-center border-l border-slate-100">
                                          {rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.pendekatan || "MINDFUL & MEANINGFUL"}
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              {/* D. ASESMEN & PENILAIAN TABLE */}
                              <div className="mb-5">
                                <h4 className="font-bold text-xs sm:text-sm text-slate-700 mb-1.5 uppercase">D. ASESMEN & PENILAIAN</h4>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full border-collapse text-left">
                                    <thead>
                                      <tr className="bg-[#007C58] text-white">
                                        <th className="p-2.5 font-bold w-[20%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Kategori Asesmen</th>
                                        <th className="p-2.5 font-bold w-[25%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Metode Penilaian</th>
                                        <th className="p-2.5 font-bold w-[25%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Instrumen Penilaian</th>
                                        <th className="p-2.5 font-bold w-[30%] text-[11px] sm:text-xs uppercase">Kriteria Keberhasilan</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Asesmen Diagnostik (Awal)</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.diagnostik?.metode || "-"}</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.diagnostik?.instrumen || "-"}</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.diagnostik?.kriteria || "-"}</td>
                                      </tr>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Asesmen Formatif (Proses)</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.formatif?.metode || "-"}</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.formatif?.instrumen || "-"}</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.formatif?.kriteria || "-"}</td>
                                      </tr>
                                      <tr>
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Asesmen Sumatif (Akhir)</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.sumatif?.metode || "-"}</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.sumatif?.instrumen || "-"}</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].asesmen?.sumatif?.kriteria || "-"}</td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              {/* E. REMEDIAL & PENGAYAAN TABLE */}
                              <div className="mb-5">
                                <h4 className="font-bold text-xs sm:text-sm text-slate-700 mb-1.5 uppercase">E. REMEDIAL & PENGAYAAN</h4>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full border-collapse text-left">
                                    <thead>
                                      <tr className="bg-[#007C58] text-white">
                                        <th className="p-2.5 font-bold w-[25%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Program Kegiatan</th>
                                        <th className="p-2.5 font-bold w-[35%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Sasaran Peserta Didik</th>
                                        <th className="p-2.5 font-bold w-[40%] text-[11px] sm:text-xs uppercase">Bentuk Aktivitas Intervensi</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Program Remedial</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].remedialPengayaan?.remedial?.sasaran || "-"}</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5">
                                            {(rpmList[selectedRpmTopic].remedialPengayaan?.remedial?.aktivitas || []).map((a, idx) => (
                                              <li key={idx}>{a}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                      <tr>
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Program Pengayaan</td>
                                        <td className="p-2.5">{rpmList[selectedRpmTopic].remedialPengayaan?.pengayaan?.sasaran || "-"}</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5">
                                            {(rpmList[selectedRpmTopic].remedialPengayaan?.pengayaan?.aktivitas || []).map((a, idx) => (
                                              <li key={idx}>{a}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              {/* F. REFLEKSI TABLE */}
                              <div className="mb-5">
                                <h4 className="font-bold text-xs sm:text-sm text-slate-700 mb-1.5 uppercase">F. REFLEKSI</h4>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full border-collapse text-left">
                                    <thead>
                                      <tr className="bg-[#007C58] text-white">
                                        <th className="p-2.5 font-bold w-[25%] border-r border-emerald-800 text-[11px] sm:text-xs uppercase">Subjek Refleksi</th>
                                        <th className="p-2.5 font-bold w-[75%] text-[11px] sm:text-xs uppercase">Pertanyaan Reflektif Pemicu Kesadaran Belajar</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      <tr className="border-b border-slate-100">
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Refleksi Peserta Didik (Siswa)</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5">
                                            {(rpmList[selectedRpmTopic].refleksi?.siswa || []).map((r, idx) => (
                                              <li key={idx}>{r}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                      <tr>
                                        <td className="p-2.5 font-bold bg-slate-50 border-r border-slate-150">Refleksi Pendidik (Guru)</td>
                                        <td className="p-2.5">
                                          <ul className="list-disc pl-5">
                                            {(rpmList[selectedRpmTopic].refleksi?.guru || []).map((r, idx) => (
                                              <li key={idx}>{r}</li>
                                            ))}
                                          </ul>
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                            </div>
                          </div>
                        </div>
                      ) : (
                        /* Edit Panel with multi-groups accordion or scroll list */
                        <div className="p-6 flex flex-col gap-6 max-h-[80vh] overflow-y-auto">
                          <div className="border-b border-slate-200 pb-3 mb-2">
                            <h3 className="font-bold text-slate-800 text-sm">Form Edit Rencana Pembelajaran Mendalam (RPM)</h3>
                            <p className="text-xs text-slate-500">Sesuaikan rincian modul ajar untuk topik ini. Nilai dalam list diisikan satu baris untuk satu poin.</p>
                          </div>

                          {/* GROUP 1: IDENTITAS ADMINISTRATIF */}
                          <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-200 flex flex-col gap-4">
                            <h4 className="font-bold text-slate-700 text-xs uppercase tracking-wide border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-600 rounded-full"></span>
                              1. Identitas Administratif
                            </h4>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Nama Guru / Penyusun</label>
                                <input 
                                  type="text" 
                                  value={rpmList[selectedRpmTopic].identitas?.namaGuru || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identitas) updated[selectedRpmTopic].identitas = {};
                                    updated[selectedRpmTopic].identitas.namaGuru = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-semibold"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Institusi / Sekolah</label>
                                <input 
                                  type="text" 
                                  value={rpmList[selectedRpmTopic].identitas?.sekolah || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identitas) updated[selectedRpmTopic].identitas = {};
                                    updated[selectedRpmTopic].identitas.sekolah = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-semibold"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Mata Pelajaran</label>
                                <input 
                                  type="text" 
                                  value={rpmList[selectedRpmTopic].identitas?.mataPelajaran || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identitas) updated[selectedRpmTopic].identitas = {};
                                    updated[selectedRpmTopic].identitas.mataPelajaran = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-semibold"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Kelas / Semester</label>
                                <input 
                                  type="text" 
                                  value={rpmList[selectedRpmTopic].identitas?.kelasSemester || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identitas) updated[selectedRpmTopic].identitas = {};
                                    updated[selectedRpmTopic].identitas.kelasSemester = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-semibold"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Semester & Tahun Ajaran</label>
                                <input 
                                  type="text" 
                                  value={rpmList[selectedRpmTopic].identitas?.semesterTahun || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identitas) updated[selectedRpmTopic].identitas = {};
                                    updated[selectedRpmTopic].identitas.semesterTahun = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-semibold"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Model Pembelajaran</label>
                                <input 
                                  type="text" 
                                  value={rpmList[selectedRpmTopic].identitas?.modelPembelajaran || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identitas) updated[selectedRpmTopic].identitas = {};
                                    updated[selectedRpmTopic].identitas.modelPembelajaran = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-semibold"
                                />
                              </div>
                            </div>
                          </div>

                          {/* GROUP 2: A. IDENTIFIKASI */}
                          <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-200 flex flex-col gap-4">
                            <h4 className="font-bold text-slate-700 text-xs uppercase tracking-wide border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-600 rounded-full"></span>
                              2. Identifikasi Aspek & Sarpras
                            </h4>
                            <div className="flex flex-col gap-3.5">
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Target Peserta Didik</label>
                                <textarea 
                                  rows={2}
                                  value={rpmList[selectedRpmTopic].identifikasi?.targetPesertaDidik || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identifikasi) updated[selectedRpmTopic].identifikasi = {};
                                    updated[selectedRpmTopic].identifikasi.targetPesertaDidik = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Kompetensi Prasyarat (Satu Poin Per Baris)</label>
                                <textarea 
                                  rows={3}
                                  value={(rpmList[selectedRpmTopic].identifikasi?.kompetensiPrasyarat || []).join("\n")}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identifikasi) updated[selectedRpmTopic].identifikasi = {};
                                    updated[selectedRpmTopic].identifikasi.kompetensiPrasyarat = e.target.value.split("\n").filter(Boolean);
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-mono"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Profil Lulusan / Dimensi SKL (Satu Poin Per Baris)</label>
                                <textarea 
                                  rows={3}
                                  value={(rpmList[selectedRpmTopic].identifikasi?.profilLulusan || []).join("\n")}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].identifikasi) updated[selectedRpmTopic].identifikasi = {};
                                    updated[selectedRpmTopic].identifikasi.profilLulusan = e.target.value.split("\n").filter(Boolean);
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-mono"
                                />
                              </div>
                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-white p-3 rounded-lg border border-slate-200">
                                <div className="flex flex-col gap-1">
                                  <label className="text-[9px] font-bold text-slate-400">FASILITAS FISIK</label>
                                  <input 
                                    type="text" 
                                    value={rpmList[selectedRpmTopic].identifikasi?.saranaPrasarana?.fasilitasFisik || ""}
                                    onChange={(e) => {
                                      const updated = { ...rpmList };
                                      if (!updated[selectedRpmTopic].identifikasi) updated[selectedRpmTopic].identifikasi = {};
                                      if (!updated[selectedRpmTopic].identifikasi.saranaPrasarana) updated[selectedRpmTopic].identifikasi.saranaPrasarana = {};
                                      updated[selectedRpmTopic].identifikasi.saranaPrasarana.fasilitasFisik = e.target.value;
                                      setRpmList(updated);
                                    }}
                                    className="border border-slate-300 rounded px-2 py-1 text-xs"
                                  />
                                </div>
                                <div className="flex flex-col gap-1">
                                  <label className="text-[9px] font-bold text-slate-400">PIRANTI LUNAK</label>
                                  <input 
                                    type="text" 
                                    value={rpmList[selectedRpmTopic].identifikasi?.saranaPrasarana?.pirantiLunak || ""}
                                    onChange={(e) => {
                                      const updated = { ...rpmList };
                                      if (!updated[selectedRpmTopic].identifikasi) updated[selectedRpmTopic].identifikasi = {};
                                      if (!updated[selectedRpmTopic].identifikasi.saranaPrasarana) updated[selectedRpmTopic].identifikasi.saranaPrasarana = {};
                                      updated[selectedRpmTopic].identifikasi.saranaPrasarana.pirantiLunak = e.target.value;
                                      setRpmList(updated);
                                    }}
                                    className="border border-slate-300 rounded px-2 py-1 text-xs"
                                  />
                                </div>
                                <div className="flex flex-col gap-1">
                                  <label className="text-[9px] font-bold text-slate-400">SUMBER BELAJAR</label>
                                  <input 
                                    type="text" 
                                    value={rpmList[selectedRpmTopic].identifikasi?.saranaPrasarana?.sumberBelajar || ""}
                                    onChange={(e) => {
                                      const updated = { ...rpmList };
                                      if (!updated[selectedRpmTopic].identifikasi) updated[selectedRpmTopic].identifikasi = {};
                                      if (!updated[selectedRpmTopic].identifikasi.saranaPrasarana) updated[selectedRpmTopic].identifikasi.saranaPrasarana = {};
                                      updated[selectedRpmTopic].identifikasi.saranaPrasarana.sumberBelajar = e.target.value;
                                      setRpmList(updated);
                                    }}
                                    className="border border-slate-300 rounded px-2 py-1 text-xs"
                                  />
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* GROUP 3: DESAIN PEMBELAJARAN */}
                          <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-200 flex flex-col gap-4">
                            <h4 className="font-bold text-slate-700 text-xs uppercase tracking-wide border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-600 rounded-full"></span>
                              3. Desain Pembelajaran (Kurikulum Merdeka)
                            </h4>
                            <div className="flex flex-col gap-3.5">
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Capaian Pembelajaran (CP)</label>
                                <textarea 
                                  rows={3}
                                  value={rpmList[selectedRpmTopic].desainPembelajaran?.capaianPembelajaran || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                    updated[selectedRpmTopic].desainPembelajaran.capaianPembelajaran = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Lintas Disiplin Ilmu (Satu Poin Per Baris)</label>
                                <textarea 
                                  rows={2}
                                  value={(rpmList[selectedRpmTopic].desainPembelajaran?.lintasDisiplin || []).join("\n")}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                    updated[selectedRpmTopic].desainPembelajaran.lintasDisiplin = e.target.value.split("\n").filter(Boolean);
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-mono"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Tujuan Pembelajaran / TP (Satu Poin Per Baris)</label>
                                <textarea 
                                  rows={3}
                                  value={(rpmList[selectedRpmTopic].desainPembelajaran?.tujuanPembelajaran || []).join("\n")}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                    updated[selectedRpmTopic].desainPembelajaran.tujuanPembelajaran = e.target.value.split("\n").filter(Boolean);
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-mono"
                                />
                              </div>
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div className="flex flex-col gap-1">
                                  <label className="text-[10px] font-bold text-slate-500 uppercase">Topik Pembelajaran</label>
                                  <input 
                                    type="text" 
                                    value={rpmList[selectedRpmTopic].desainPembelajaran?.topikPembelajaran || ""}
                                    onChange={(e) => {
                                      const updated = { ...rpmList };
                                      if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                      updated[selectedRpmTopic].desainPembelajaran.topikPembelajaran = e.target.value;
                                      setRpmList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                  />
                                </div>
                                <div className="flex flex-col gap-1">
                                  <label className="text-[10px] font-bold text-slate-500 uppercase">Lingkungan Belajar</label>
                                  <input 
                                    type="text" 
                                    value={rpmList[selectedRpmTopic].desainPembelajaran?.lingkunganBelajar || ""}
                                    onChange={(e) => {
                                      const updated = { ...rpmList };
                                      if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                      updated[selectedRpmTopic].desainPembelajaran.lingkunganBelajar = e.target.value;
                                      setRpmList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                  />
                                </div>
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Praktik Pedagogis</label>
                                <textarea 
                                  rows={2}
                                  value={rpmList[selectedRpmTopic].desainPembelajaran?.praktikPedagogis || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                    updated[selectedRpmTopic].desainPembelajaran.praktikPedagogis = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Kemitraan Pembelajaran / Kontekstual Dunia Nyata</label>
                                <textarea 
                                  rows={2}
                                  value={rpmList[selectedRpmTopic].desainPembelajaran?.kemitraanPembelajaran || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                    updated[selectedRpmTopic].desainPembelajaran.kemitraanPembelajaran = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                />
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-bold text-slate-500 uppercase">Pemanfaatan Teknologi Digital</label>
                                <textarea 
                                  rows={2}
                                  value={rpmList[selectedRpmTopic].desainPembelajaran?.pemanfaatanDigital || ""}
                                  onChange={(e) => {
                                    const updated = { ...rpmList };
                                    if (!updated[selectedRpmTopic].desainPembelajaran) updated[selectedRpmTopic].desainPembelajaran = {};
                                    updated[selectedRpmTopic].desainPembelajaran.pemanfaatanDigital = e.target.value;
                                    setRpmList(updated);
                                  }}
                                  className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs"
                                />
                              </div>
                            </div>
                          </div>

                          {/* GROUP 4: LANGKAH PEMBELAJARAN */}
                          <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-200 flex flex-col gap-4">
                            <h4 className="font-bold text-slate-700 text-xs uppercase tracking-wide border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-600 rounded-full"></span>
                              4. Skenario Langkah Pembelajaran
                            </h4>
                            <div className="flex flex-col gap-4">
                              {/* Sub-block Pendahuluan */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <div className="flex justify-between items-center border-b border-slate-100 pb-1">
                                  <strong className="text-xs text-slate-800 uppercase">A. Kegiatan Pendahuluan</strong>
                                  <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded font-mono font-bold text-slate-500">MINDFUL</span>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">DURASI</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.durasi || "15 Menit"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.pendahuluan) updated[selectedRpmTopic].langkahPembelajaran.pendahuluan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.pendahuluan.durasi = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">PENDEKATAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.pendekatan || "MINDFUL"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.pendahuluan) updated[selectedRpmTopic].langkahPembelajaran.pendahuluan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.pendahuluan.pendekatan = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS GURU (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.guru || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.pendahuluan) updated[selectedRpmTopic].langkahPembelajaran.pendahuluan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.pendahuluan.guru = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS SISWA (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.pendahuluan?.siswa || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.pendahuluan) updated[selectedRpmTopic].langkahPembelajaran.pendahuluan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.pendahuluan.siswa = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>

                              {/* Sub-block Eksplorasi */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <div className="flex justify-between items-center border-b border-slate-100 pb-1">
                                  <strong className="text-xs text-slate-800 uppercase">B. Kegiatan Inti: Eksplorasi Konsep</strong>
                                  <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded font-mono font-bold text-slate-500">MEANINGFUL</span>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">DURASI</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.durasi || "45 Menit"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi) updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi.durasi = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">PENDEKATAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.pendekatan || "MEANINGFUL"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi) updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi.pendekatan = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS GURU (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.guru || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi) updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi.guru = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS SISWA (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.intiEksplorasi?.siswa || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi) updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiEksplorasi.siswa = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>

                              {/* Sub-block Rancangan */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <div className="flex justify-between items-center border-b border-slate-100 pb-1">
                                  <strong className="text-xs text-slate-800 uppercase">C. Kegiatan Inti: Perancangan & Simulasi</strong>
                                  <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded font-mono font-bold text-slate-500">JOYFUL</span>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">DURASI</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.durasi || "45 Menit"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiRancangan) updated[selectedRpmTopic].langkahPembelajaran.intiRancangan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiRancangan.durasi = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">PENDEKATAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.pendekatan || "JOYFUL"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiRancangan) updated[selectedRpmTopic].langkahPembelajaran.intiRancangan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiRancangan.pendekatan = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS GURU (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.guru || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiRancangan) updated[selectedRpmTopic].langkahPembelajaran.intiRancangan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiRancangan.guru = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS SISWA (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.intiRancangan?.siswa || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.intiRancangan) updated[selectedRpmTopic].langkahPembelajaran.intiRancangan = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.intiRancangan.siswa = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>

                              {/* Sub-block Penutup */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <div className="flex justify-between items-center border-b border-slate-100 pb-1">
                                  <strong className="text-xs text-slate-800 uppercase">D. Kegiatan Penutup & Evaluasi</strong>
                                  <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded font-mono font-bold text-slate-500">MINDFUL & MEANINGFUL</span>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">DURASI</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.durasi || "15 Menit"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.penutup) updated[selectedRpmTopic].langkahPembelajaran.penutup = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.penutup.durasi = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">PENDEKATAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.pendekatan || "MINDFUL & MEANINGFUL"}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.penutup) updated[selectedRpmTopic].langkahPembelajaran.penutup = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.penutup.pendekatan = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-semibold"
                                    />
                                  </div>
                                </div>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS GURU (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.guru || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.penutup) updated[selectedRpmTopic].langkahPembelajaran.penutup = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.penutup.guru = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">AKTIVITAS SISWA (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].langkahPembelajaran?.penutup?.siswa || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].langkahPembelajaran) updated[selectedRpmTopic].langkahPembelajaran = {};
                                        if (!updated[selectedRpmTopic].langkahPembelajaran.penutup) updated[selectedRpmTopic].langkahPembelajaran.penutup = {};
                                        updated[selectedRpmTopic].langkahPembelajaran.penutup.siswa = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* GROUP 5: ASESMEN */}
                          <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-200 flex flex-col gap-4">
                            <h4 className="font-bold text-slate-700 text-xs uppercase tracking-wide border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-600 rounded-full"></span>
                              5. Desain Asesmen & Penilaian
                            </h4>
                            <div className="flex flex-col gap-4">
                              {/* Diagnostik */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-2.5">
                                <span className="text-xs font-bold text-slate-700 uppercase">A. Asesmen Diagnostik (Awal)</span>
                                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">METODE PENILAIAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.diagnostik?.metode || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.diagnostik) updated[selectedRpmTopic].asesmen.diagnostik = {};
                                        updated[selectedRpmTopic].asesmen.diagnostik.metode = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">INSTRUMEN PENILAIAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.diagnostik?.instrumen || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.diagnostik) updated[selectedRpmTopic].asesmen.diagnostik = {};
                                        updated[selectedRpmTopic].asesmen.diagnostik.instrumen = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">KRITERIA KEBERHASILAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.diagnostik?.kriteria || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.diagnostik) updated[selectedRpmTopic].asesmen.diagnostik = {};
                                        updated[selectedRpmTopic].asesmen.diagnostik.kriteria = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                </div>
                              </div>

                              {/* Formatif */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-2.5">
                                <span className="text-xs font-bold text-slate-700 uppercase">B. Asesmen Formatif (Proses)</span>
                                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">METODE PENILAIAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.formatif?.metode || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.formatif) updated[selectedRpmTopic].asesmen.formatif = {};
                                        updated[selectedRpmTopic].asesmen.formatif.metode = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">INSTRUMEN PENILAIAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.formatif?.instrumen || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.formatif) updated[selectedRpmTopic].asesmen.formatif = {};
                                        updated[selectedRpmTopic].asesmen.formatif.instrumen = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">KRITERIA KEBERHASILAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.formatif?.kriteria || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.formatif) updated[selectedRpmTopic].asesmen.formatif = {};
                                        updated[selectedRpmTopic].asesmen.formatif.kriteria = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                </div>
                              </div>

                              {/* Sumatif */}
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-2.5">
                                <span className="text-xs font-bold text-slate-700 uppercase">C. Asesmen Sumatif (Akhir)</span>
                                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">METODE PENILAIAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.sumatif?.metode || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.sumatif) updated[selectedRpmTopic].asesmen.sumatif = {};
                                        updated[selectedRpmTopic].asesmen.sumatif.metode = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">INSTRUMEN PENILAIAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.sumatif?.instrumen || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.sumatif) updated[selectedRpmTopic].asesmen.sumatif = {};
                                        updated[selectedRpmTopic].asesmen.sumatif.instrumen = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">KRITERIA KEBERHASILAN</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].asesmen?.sumatif?.kriteria || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].asesmen) updated[selectedRpmTopic].asesmen = {};
                                        if (!updated[selectedRpmTopic].asesmen.sumatif) updated[selectedRpmTopic].asesmen.sumatif = {};
                                        updated[selectedRpmTopic].asesmen.sumatif.kriteria = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* GROUP 6: REMEDIAL, PENGAYAAN & REFLEKSI */}
                          <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-200 flex flex-col gap-4">
                            <h4 className="font-bold text-slate-700 text-xs uppercase tracking-wide border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                              <span className="w-2 h-2 bg-emerald-600 rounded-full"></span>
                              6. Remedial, Pengayaan & Refleksi
                            </h4>
                            <div className="flex flex-col gap-4">
                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <span className="text-xs font-bold text-slate-700 uppercase">A. Program Remedial</span>
                                <div className="flex flex-col gap-2">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">SASARAN PESERTA DIDIK</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].remedialPengayaan?.remedial?.sasaran || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].remedialPengayaan) updated[selectedRpmTopic].remedialPengayaan = {};
                                        if (!updated[selectedRpmTopic].remedialPengayaan.remedial) updated[selectedRpmTopic].remedialPengayaan.remedial = {};
                                        updated[selectedRpmTopic].remedialPengayaan.remedial.sasaran = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">BENTUK AKTIVITAS INTERVENSI (Poin Per Baris)</label>
                                    <textarea 
                                      rows={2}
                                      value={(rpmList[selectedRpmTopic].remedialPengayaan?.remedial?.aktivitas || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].remedialPengayaan) updated[selectedRpmTopic].remedialPengayaan = {};
                                        if (!updated[selectedRpmTopic].remedialPengayaan.remedial) updated[selectedRpmTopic].remedialPengayaan.remedial = {};
                                        updated[selectedRpmTopic].remedialPengayaan.remedial.aktivitas = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>

                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <span className="text-xs font-bold text-slate-700 uppercase">B. Program Pengayaan</span>
                                <div className="flex flex-col gap-2">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">SASARAN PESERTA DIDIK</label>
                                    <input 
                                      type="text" 
                                      value={rpmList[selectedRpmTopic].remedialPengayaan?.pengayaan?.sasaran || ""}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].remedialPengayaan) updated[selectedRpmTopic].remedialPengayaan = {};
                                        if (!updated[selectedRpmTopic].remedialPengayaan.pengayaan) updated[selectedRpmTopic].remedialPengayaan.pengayaan = {};
                                        updated[selectedRpmTopic].remedialPengayaan.pengayaan.sasaran = e.target.value;
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">BENTUK AKTIVITAS INTERVENSI (Poin Per Baris)</label>
                                    <textarea 
                                      rows={2}
                                      value={(rpmList[selectedRpmTopic].remedialPengayaan?.pengayaan?.aktivitas || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].remedialPengayaan) updated[selectedRpmTopic].remedialPengayaan = {};
                                        if (!updated[selectedRpmTopic].remedialPengayaan.pengayaan) updated[selectedRpmTopic].remedialPengayaan.pengayaan = {};
                                        updated[selectedRpmTopic].remedialPengayaan.pengayaan.aktivitas = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>

                              <div className="bg-white p-3 rounded-lg border border-slate-200 flex flex-col gap-3">
                                <span className="text-xs font-bold text-slate-700 uppercase">C. Pertanyaan Refleksi</span>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">REFLEKSI SISWA (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].refleksi?.siswa || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].refleksi) updated[selectedRpmTopic].refleksi = {};
                                        updated[selectedRpmTopic].refleksi.siswa = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                  <div className="flex flex-col gap-1">
                                    <label className="text-[9px] font-bold text-slate-400">REFLEKSI GURU (Poin Per Baris)</label>
                                    <textarea 
                                      rows={3}
                                      value={(rpmList[selectedRpmTopic].refleksi?.guru || []).join("\n")}
                                      onChange={(e) => {
                                        const updated = { ...rpmList };
                                        if (!updated[selectedRpmTopic].refleksi) updated[selectedRpmTopic].refleksi = {};
                                        updated[selectedRpmTopic].refleksi.guru = e.target.value.split("\n").filter(Boolean);
                                        setRpmList(updated);
                                      }}
                                      className="border border-slate-300 rounded px-2 py-1 text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                        </div>
                      )}
                    </div>
                  ) : (
                    /* Empty State Block */
                    <div className="flex-1 flex flex-col items-center justify-center p-8 text-center gap-4">
                      <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center text-slate-400 border border-slate-200">
                        <FileText className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-bold text-slate-800 text-sm">RPM Belum Dibuat</h4>
                        <p className="text-slate-500 text-xs mt-1 max-w-sm">Rencana Pembelajaran Mendalam untuk topik ini belum digenerate. Klik tombol di bawah untuk membuat secara instan.</p>
                      </div>
                      <button
                        onClick={() => {
                          const t = topikData.daftarTopik.find(x => x.topik === selectedRpmTopic);
                          if (t) generateRpmSingle(t);
                        }}
                        className="bg-slate-900 hover:bg-slate-800 text-white font-bold py-2 px-4 rounded text-xs transition-colors cursor-pointer"
                      >
                        Generate RPM Sekarang
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(5)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center flex-wrap gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => {
                      const t = topikData.daftarTopik.find(x => x.topik === selectedRpmTopic);
                      if (t) generateRpmSingle(t);
                    }}
                    disabled={!rpmList[selectedRpmTopic]}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>

                  <button
                    onClick={() => setStep(7)}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 hover:shadow-md"
                  >
                    <span>Buat LKPD</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 7: LKPD */}
          {step === 7 && topikData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              {/* Batch Action Bar */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                <div className="text-xs sm:text-sm">
                  <span className="font-bold text-slate-800">Lembar Kerja Peserta Didik (LKPD)</span>
                  <p className="text-slate-500 text-[10px] sm:text-xs mt-0.5">Pembuatan instrumen penugasan mandiri / kelompok bermakna untuk setiap materi pokok.</p>
                </div>
                <button
                  onClick={generateAllLkpd}
                  className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded text-xs transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Generate Semua LKPD Otomatis</span>
                </button>
              </div>

              <div className="flex flex-col lg:flex-row gap-6 items-stretch">
                {/* Topic selector panel */}
                <div className="w-full lg:w-72 bg-white rounded-xl border border-slate-200 p-4 flex flex-col gap-3 shrink-0">
                  <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Materi Pokok</h3>
                  <div className="flex flex-col gap-2 max-h-[300px] lg:max-h-none overflow-y-auto">
                    {topikData.daftarTopik.map((t, idx) => {
                      const isSelected = selectedLkpdTopic === t.topik;
                      const hasData = !!lkpdList[t.topik];
                      return (
                        <button
                          key={idx}
                          onClick={() => setSelectedLkpdTopic(t.topik)}
                          className={`w-full text-left p-3 rounded-lg border transition-all text-xs flex flex-col gap-1.5 cursor-pointer ${
                            isSelected 
                              ? "bg-slate-900 text-white border-slate-950 shadow" 
                              : "bg-white hover:bg-slate-50 text-slate-700 border-slate-200"
                          }`}
                        >
                          <div className="flex items-center justify-between gap-1 w-full">
                            <span className="font-mono font-bold">Topik {t.no}</span>
                            <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                              hasData 
                                ? "bg-green-100 text-green-800 border border-green-200" 
                                : "bg-slate-100 text-slate-500 border border-slate-200"
                            }`}>
                              {hasData ? "Sudah Dibuat" : "Belum Dibuat"}
                            </span>
                          </div>
                          <p className={`font-semibold line-clamp-2 leading-tight ${isSelected ? "text-slate-100" : "text-slate-800"}`}>
                            {t.topik}
                          </p>
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* Main Preview and Editor Panel */}
                <div className="flex-1 bg-white border border-slate-200 rounded-xl shadow-md flex flex-col min-h-[450px]">
                  {lkpdList[selectedLkpdTopic] ? (
                    <div>
                      {!editMode ? (
                        /* Preview Paper */
                        <div className="p-6 md:p-10 max-w-4xl mx-auto flex-1 flex flex-col font-sans leading-relaxed text-xs sm:text-sm text-slate-900 w-full bg-white">
                          {/* Banner Header */}
                          <div className="bg-[#007C58] text-white p-6 rounded-t-lg text-center shadow-sm">
                            <h1 className="font-bold text-lg sm:text-xl uppercase tracking-wider">
                              {lkpdList[selectedLkpdTopic].judul || "LEMBAR KERJA PESERTA DIDIK (LKPD)"}
                            </h1>
                            <p className="text-[11px] sm:text-xs mt-1.5 opacity-90 font-semibold uppercase tracking-wide">
                              Mata Pelajaran: {metadata.mataPelajaran} | Fase/Kelas: {metadata.fase.toUpperCase()} / {metadata.fase === 'f' ? 'XI' : 'X'}
                            </p>
                          </div>

                          {/* Metadata / Identitas Table */}
                          <div className="overflow-x-auto mt-4">
                            <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                              <tbody>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800 w-[30%]">Nama Guru / Penyusun</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700 w-[70%]">{lkpdList[selectedLkpdTopic].identitas?.namaGuru || metadata.namaGuru || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800">Instansi / Sekolah</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700">{lkpdList[selectedLkpdTopic].identitas?.sekolah || metadata.jenjangSekolah || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800">Mata Pelajaran</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700">{lkpdList[selectedLkpdTopic].identitas?.mataPelajaran || metadata.mataPelajaran || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800">Jenjang / Fase / Kelas</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700">SMK/MAK / {metadata.fase.toUpperCase()} / {metadata.fase === 'f' ? 'XI' : 'X'}</td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800">Semester & Tahun Ajaran</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700">{lkpdList[selectedLkpdTopic].identitas?.semesterTahun || `Semester 1 - ${metadata.tahunAjaran}`}</td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800">Topik / Materi Pokok</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700 font-semibold">{lkpdList[selectedLkpdTopic].identitas?.topikMateri || selectedLkpdTopic}</td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800">Model Pembelajaran</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700">{lkpdList[selectedLkpdTopic].identitas?.modelPembelajaran || "Tatap Muka / Reguler"}</td>
                                </tr>
                              </tbody>
                            </table>
                          </div>

                          {/* I. TUJUAN & PETUNJUK */}
                          <div className="mt-8">
                            <div className="border-l-4 border-emerald-600 pl-3 py-1 mb-3 bg-slate-50/50">
                              <h3 className="font-bold text-sm text-emerald-700 uppercase tracking-wide">Tujuan & Petunjuk</h3>
                            </div>
                            <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                              <thead>
                                <tr className="bg-[#007C58] text-white">
                                  <th className="border border-slate-300 p-2.5 font-bold w-[30%]">Aspek Kerja</th>
                                  <th className="border border-slate-300 p-2.5 font-bold w-[70%]">Deskripsi Panduan</th>
                                </tr>
                              </thead>
                              <tbody>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800 align-top">Tujuan Pembelajaran</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700 align-top">
                                    <ul className="list-disc pl-4 space-y-1">
                                      {Array.isArray(lkpdList[selectedLkpdTopic].tujuanPembelajaran) ? (
                                        lkpdList[selectedLkpdTopic].tujuanPembelajaran.map((tp, i) => (
                                          <li key={i}>{tp}</li>
                                        ))
                                      ) : (
                                        <li>[ {lkpdList[selectedLkpdTopic].tpKode} ] {lkpdList[selectedLkpdTopic].judul}</li>
                                      )}
                                    </ul>
                                  </td>
                                </tr>
                                <tr>
                                  <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800 align-top">Petunjuk Pengerjaan</td>
                                  <td className="border border-slate-300 p-2.5 text-slate-700 align-top">
                                    <ol className="list-decimal pl-4 space-y-1">
                                      {(Array.isArray(lkpdList[selectedLkpdTopic].petunjukPengerjaan) ? lkpdList[selectedLkpdTopic].petunjukPengerjaan : (Array.isArray(lkpdList[selectedLkpdTopic].petunjukBelajar) ? lkpdList[selectedLkpdTopic].petunjukBelajar : [])).map((step, i) => (
                                        <li key={i}>{step}</li>
                                      ))}
                                    </ol>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </div>

                          {/* II. RINGKASAN MATERI */}
                          {lkpdList[selectedLkpdTopic].ringkasanMateri && (
                            <div className="mt-8">
                              <div className="border-l-4 border-emerald-600 pl-3 py-1 mb-3 bg-slate-50/50">
                                <h3 className="font-bold text-sm text-emerald-700 uppercase tracking-wide">Ringkasan Materi</h3>
                              </div>
                              <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                                <thead>
                                  <tr className="bg-[#007C58] text-white">
                                    <th className="border border-slate-300 p-2.5 font-bold w-[30%]">Komponen Utama</th>
                                    <th className="border border-slate-300 p-2.5 font-bold w-[70%]">Ringkasan Konseptual {selectedLkpdTopic}</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {lkpdList[selectedLkpdTopic].ringkasanMateri.map((materi, idx) => (
                                    <tr key={idx}>
                                      <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-slate-800 align-top">{materi.komponen}</td>
                                      <td className="border border-slate-300 p-2.5 text-slate-700 align-top whitespace-pre-wrap leading-relaxed">
                                        {materi.ringkasan}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}

                          {/* III. AKTIVITAS / SOAL */}
                          <div className="mt-8">
                            <div className="border-l-4 border-emerald-600 pl-3 py-1 mb-3 bg-slate-50/50">
                              <h3 className="font-bold text-sm text-emerald-700 uppercase tracking-wide">Aktivitas / Soal</h3>
                            </div>
                            <div className="flex flex-col gap-5">
                              {Array.isArray(lkpdList[selectedLkpdTopic].aktivitasList) ? (
                                lkpdList[selectedLkpdTopic].aktivitasList.map((item, idx) => (
                                  <div key={idx} className="border border-slate-300 rounded-lg overflow-hidden shadow-sm">
                                    <div className="bg-[#007C58] text-white px-4 py-2 font-bold text-xs">
                                      {item.judul || `Aktivitas ${idx + 1}`}
                                    </div>
                                    <div className="bg-[#F0FDF4]/30 p-4 text-xs sm:text-sm text-slate-800 leading-relaxed whitespace-pre-wrap">
                                      {item.studiKasus}
                                    </div>
                                  </div>
                                ))
                              ) : (
                                /* Fallback Legacy Data */
                                <div className="border border-slate-300 rounded-lg overflow-hidden shadow-sm">
                                  <div className="bg-[#007C58] text-white px-4 py-2 font-bold text-xs">
                                    Aktivitas Penugasan Kelompok
                                  </div>
                                  <div className="bg-[#F0FDF4]/30 p-4 text-xs sm:text-sm text-slate-800 leading-relaxed whitespace-pre-wrap">
                                    <p className="font-bold text-[#007C58] mb-1.5">Instruksi Kerja:</p>
                                    <p className="mb-4">{lkpdList[selectedLkpdTopic].tugasAktivitas?.instruksi || lkpdList[selectedLkpdTopic].tugasAktivitas}</p>
                                    
                                    {lkpdList[selectedLkpdTopic].tugasAktivitas?.pertanyaan && (
                                      <div>
                                        <p className="font-bold text-[#007C58] mb-1.5">Pertanyaan Diskusi:</p>
                                        <ul className="list-disc pl-4 space-y-1.5">
                                          {lkpdList[selectedLkpdTopic].tugasAktivitas.pertanyaan.map((q, i) => (
                                            <li key={i}>{q}</li>
                                          ))}
                                        </ul>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      ) : (
                        /* Edit Panel */
                        <div className="p-6 flex flex-col gap-6">
                          <h3 className="font-bold text-slate-800 text-sm border-b border-slate-200 pb-2 uppercase tracking-wide">Edit Lembar Kerja Peserta Didik (LKPD)</h3>
                          
                          {/* Identitas Section */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 bg-slate-50 p-4 rounded-lg border border-slate-200">
                            <h4 className="col-span-full font-bold text-slate-700 text-xs uppercase tracking-wider">Identitas LKPD</h4>
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-500">Nama Guru / Penyusun</span>
                              <input 
                                type="text"
                                value={lkpdList[selectedLkpdTopic].identitas?.namaGuru || ""}
                                onChange={(e) => {
                                  const updated = { ...lkpdList };
                                  if (!updated[selectedLkpdTopic].identitas) updated[selectedLkpdTopic].identitas = {} as any;
                                  updated[selectedLkpdTopic].identitas.namaGuru = e.target.value;
                                  setLkpdList(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-500">Instansi / Sekolah</span>
                              <input 
                                type="text"
                                value={lkpdList[selectedLkpdTopic].identitas?.sekolah || ""}
                                onChange={(e) => {
                                  const updated = { ...lkpdList };
                                  if (!updated[selectedLkpdTopic].identitas) updated[selectedLkpdTopic].identitas = {} as any;
                                  updated[selectedLkpdTopic].identitas.sekolah = e.target.value;
                                  setLkpdList(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-500">Semester & Tahun Ajaran</span>
                              <input 
                                type="text"
                                value={lkpdList[selectedLkpdTopic].identitas?.semesterTahun || ""}
                                onChange={(e) => {
                                  const updated = { ...lkpdList };
                                  if (!updated[selectedLkpdTopic].identitas) updated[selectedLkpdTopic].identitas = {} as any;
                                  updated[selectedLkpdTopic].identitas.semesterTahun = e.target.value;
                                  setLkpdList(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <span className="text-[10px] font-bold text-slate-500">Model Pembelajaran</span>
                              <input 
                                type="text"
                                value={lkpdList[selectedLkpdTopic].identitas?.modelPembelajaran || ""}
                                onChange={(e) => {
                                  const updated = { ...lkpdList };
                                  if (!updated[selectedLkpdTopic].identitas) updated[selectedLkpdTopic].identitas = {} as any;
                                  updated[selectedLkpdTopic].identitas.modelPembelajaran = e.target.value;
                                  setLkpdList(updated);
                                }}
                                className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                              />
                            </div>
                          </div>

                          {/* Tujuan & Petunjuk Section */}
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div className="flex flex-col gap-1.5">
                              <label className="text-xs font-bold text-slate-600">Tujuan Pembelajaran (Satu per baris)</label>
                              <textarea 
                                rows={4}
                                value={Array.isArray(lkpdList[selectedLkpdTopic].tujuanPembelajaran) ? lkpdList[selectedLkpdTopic].tujuanPembelajaran.join("\n") : ""}
                                onChange={(e) => {
                                  const updated = { ...lkpdList };
                                  updated[selectedLkpdTopic].tujuanPembelajaran = e.target.value.split("\n");
                                  setLkpdList(updated);
                                }}
                                placeholder="Masukkan tujuan pembelajaran..."
                                className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs leading-relaxed font-sans"
                              />
                            </div>
                            <div className="flex flex-col gap-1.5">
                              <label className="text-xs font-bold text-slate-600">Petunjuk Pengerjaan (Satu per baris)</label>
                              <textarea 
                                rows={4}
                                value={(Array.isArray(lkpdList[selectedLkpdTopic].petunjukPengerjaan) ? lkpdList[selectedLkpdTopic].petunjukPengerjaan : (Array.isArray(lkpdList[selectedLkpdTopic].petunjukBelajar) ? lkpdList[selectedLkpdTopic].petunjukBelajar : [])).join("\n")}
                                onChange={(e) => {
                                  const updated = { ...lkpdList };
                                  updated[selectedLkpdTopic].petunjukPengerjaan = e.target.value.split("\n");
                                  updated[selectedLkpdTopic].petunjukBelajar = e.target.value.split("\n");
                                  setLkpdList(updated);
                                }}
                                placeholder="Masukkan petunjuk pengerjaan..."
                                className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs leading-relaxed font-sans"
                              />
                            </div>
                          </div>

                          {/* Ringkasan Materi */}
                          <div className="flex flex-col gap-3">
                            <div className="flex justify-between items-center border-b border-slate-200 pb-1">
                              <span className="text-xs font-bold text-slate-600 uppercase">Ringkasan Materi</span>
                              <button
                                onClick={() => {
                                  const updated = { ...lkpdList };
                                  if (!updated[selectedLkpdTopic].ringkasanMateri) updated[selectedLkpdTopic].ringkasanMateri = [];
                                  updated[selectedLkpdTopic].ringkasanMateri.push({ komponen: "Komponen Baru", ringkasan: "Penjelasan..." });
                                  setLkpdList(updated);
                                }}
                                className="bg-slate-800 hover:bg-slate-900 text-white font-semibold py-1 px-2.5 rounded text-[10px] cursor-pointer"
                              >
                                + Tambah Komponen
                              </button>
                            </div>
                            
                            {Array.isArray(lkpdList[selectedLkpdTopic].ringkasanMateri) && lkpdList[selectedLkpdTopic].ringkasanMateri.map((materi, idx) => (
                              <div key={idx} className="bg-slate-50 p-3 rounded border border-slate-200 flex flex-col gap-2.5 relative">
                                <button
                                  onClick={() => {
                                    const updated = { ...lkpdList };
                                    updated[selectedLkpdTopic].ringkasanMateri.splice(idx, 1);
                                    setLkpdList(updated);
                                  }}
                                  className="absolute top-2 right-2 text-red-500 hover:text-red-700 font-bold text-xs cursor-pointer"
                                >
                                  Hapus
                                </button>
                                <div className="flex flex-col gap-1 w-[80%]">
                                  <span className="text-[9px] font-bold text-slate-400">KOMPONEN UTAMA</span>
                                  <input 
                                    type="text" 
                                    value={materi.komponen || ""}
                                    onChange={(e) => {
                                      const updated = { ...lkpdList };
                                      updated[selectedLkpdTopic].ringkasanMateri[idx].komponen = e.target.value;
                                      setLkpdList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs font-semibold"
                                  />
                                </div>
                                <div className="flex flex-col gap-1">
                                  <span className="text-[9px] font-bold text-slate-400">RINGKASAN KONSEPTUAL</span>
                                  <textarea 
                                    rows={4}
                                    value={materi.ringkasan || ""}
                                    onChange={(e) => {
                                      const updated = { ...lkpdList };
                                      updated[selectedLkpdTopic].ringkasanMateri[idx].ringkasan = e.target.value;
                                      setLkpdList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-sans leading-relaxed"
                                  />
                                </div>
                              </div>
                            ))}
                          </div>

                          {/* Aktivitas List */}
                          <div className="flex flex-col gap-3">
                            <div className="flex justify-between items-center border-b border-slate-200 pb-1">
                              <span className="text-xs font-bold text-slate-600 uppercase">Aktivitas / Soal</span>
                              <button
                                onClick={() => {
                                  const updated = { ...lkpdList };
                                  if (!updated[selectedLkpdTopic].aktivitasList) updated[selectedLkpdTopic].aktivitasList = [];
                                  const count = updated[selectedLkpdTopic].aktivitasList.length;
                                  updated[selectedLkpdTopic].aktivitasList.push({ judul: `Aktivitas ${count + 1}: Judul Aktivitas`, studiKasus: "Studi Kasus:\nDeskripsi..." });
                                  setLkpdList(updated);
                                }}
                                className="bg-slate-800 hover:bg-slate-900 text-white font-semibold py-1 px-2.5 rounded text-[10px] cursor-pointer"
                              >
                                + Tambah Aktivitas
                              </button>
                            </div>

                            {Array.isArray(lkpdList[selectedLkpdTopic].aktivitasList) && lkpdList[selectedLkpdTopic].aktivitasList.map((item, idx) => (
                              <div key={idx} className="bg-slate-50 p-3 rounded border border-slate-200 flex flex-col gap-2.5 relative">
                                <button
                                  onClick={() => {
                                    const updated = { ...lkpdList };
                                    updated[selectedLkpdTopic].aktivitasList.splice(idx, 1);
                                    setLkpdList(updated);
                                  }}
                                  className="absolute top-2 right-2 text-red-500 hover:text-red-700 font-bold text-xs cursor-pointer"
                                >
                                  Hapus
                                </button>
                                <div className="flex flex-col gap-1 w-[80%]">
                                  <span className="text-[9px] font-bold text-slate-400">JUDUL AKTIVITAS</span>
                                  <input 
                                    type="text" 
                                    value={item.judul || ""}
                                    onChange={(e) => {
                                      const updated = { ...lkpdList };
                                      updated[selectedLkpdTopic].aktivitasList[idx].judul = e.target.value;
                                      setLkpdList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs font-semibold"
                                  />
                                </div>
                                <div className="flex flex-col gap-1">
                                  <span className="text-[9px] font-bold text-slate-400">STUDI KASUS / DESKRIPSI & PERTANYAAN</span>
                                  <textarea 
                                    rows={5}
                                    value={item.studiKasus || ""}
                                    onChange={(e) => {
                                      const updated = { ...lkpdList };
                                      updated[selectedLkpdTopic].aktivitasList[idx].studiKasus = e.target.value;
                                      setLkpdList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs font-sans leading-relaxed"
                                  />
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  ) : (
                    /* Empty State Block */
                    <div className="flex-1 flex flex-col items-center justify-center p-8 text-center gap-4">
                      <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center text-slate-400 border border-slate-200">
                        <FileSpreadsheet className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-bold text-slate-800 text-sm">LKPD Belum Dibuat</h4>
                        <p className="text-slate-500 text-xs mt-1 max-w-sm">Lembar Kerja Peserta Didik untuk topik ini belum digenerate. Klik tombol di bawah untuk membuat secara instan.</p>
                      </div>
                      <button
                        onClick={() => {
                          const t = topikData.daftarTopik.find(x => x.topik === selectedLkpdTopic);
                          if (t) generateLkpdSingle(t);
                        }}
                        className="bg-slate-900 hover:bg-slate-800 text-white font-bold py-2 px-4 rounded text-xs transition-colors cursor-pointer"
                      >
                        Generate LKPD Sekarang
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(6)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center flex-wrap gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => {
                      const t = topikData.daftarTopik.find(x => x.topik === selectedLkpdTopic);
                      if (t) generateLkpdSingle(t);
                    }}
                    disabled={!lkpdList[selectedLkpdTopic]}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>

                  <button
                    onClick={() => setStep(8)}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 hover:shadow-md"
                  >
                    <span>Buat Soal Asesmen</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 8: SOAL ASESMEN & KISI-KISI */}
          {step === 8 && topikData && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex-1 flex flex-col gap-6"
            >
              {/* Batch Action Bar */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                <div className="text-xs sm:text-sm">
                  <span className="font-bold text-slate-800">Soal Asesmen Pilihan Ganda & Kisi-kisi</span>
                  <p className="text-slate-500 text-[10px] sm:text-xs mt-0.5">Pembuatan bank soal pilihan ganda (5 opsi) lengkap dengan kunci jawaban dan kisi-kisi per topik.</p>
                </div>
                <button
                  onClick={generateAllAsesmen}
                  className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded text-xs transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Generate Semua Asesmen Otomatis</span>
                </button>
              </div>

              <div className="flex flex-col lg:flex-row gap-6 items-stretch">
                {/* Topic selector panel */}
                <div className="w-full lg:w-72 bg-white rounded-xl border border-slate-200 p-4 flex flex-col gap-3 shrink-0">
                  <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Materi Pokok</h3>
                  <div className="flex flex-col gap-2 max-h-[300px] lg:max-h-none overflow-y-auto">
                    {topikData.daftarTopik.map((t, idx) => {
                      const isSelected = selectedAsesmenTopic === t.topik;
                      const hasData = !!asesmenList[t.topik];
                      return (
                        <button
                          key={idx}
                          onClick={() => setSelectedAsesmenTopic(t.topik)}
                          className={`w-full text-left p-3 rounded-lg border transition-all text-xs flex flex-col gap-1.5 cursor-pointer ${
                            isSelected 
                              ? "bg-slate-900 text-white border-slate-950 shadow" 
                              : "bg-white hover:bg-slate-50 text-slate-700 border-slate-200"
                          }`}
                        >
                          <div className="flex items-center justify-between gap-1 w-full">
                            <span className="font-mono font-bold">Topik {t.no}</span>
                            <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                              hasData 
                                ? "bg-green-100 text-green-800 border border-green-200" 
                                : "bg-slate-100 text-slate-500 border border-slate-200"
                            }`}>
                              {hasData ? "Sudah Dibuat" : "Belum Dibuat"}
                            </span>
                          </div>
                          <p className={`font-semibold line-clamp-2 leading-tight ${isSelected ? "text-slate-100" : "text-slate-800"}`}>
                            {t.topik}
                          </p>
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* Main Preview and Editor Panel */}
                <div className="flex-1 bg-white border border-slate-200 rounded-xl shadow-md flex flex-col min-h-[450px]">
                  {asesmenList[selectedAsesmenTopic] ? (
                    <div>
                      {!editMode ? (
                        /* Preview Paper */
                        <div className="p-8 md:p-12 max-w-4xl mx-auto flex-1 flex flex-col font-sans leading-relaxed text-sm sm:text-base text-slate-900 A4-paper w-full bg-white">
                          
                          {/* Banner Header */}
                          <div className="bg-[#007A54] text-white p-6 rounded-lg text-center shadow-sm mb-6">
                            <h1 className="font-sans font-bold text-lg sm:text-2xl uppercase tracking-wider mb-1">
                              KISI-KISI & SOAL ASESMEN EVALUASI
                            </h1>
                            <p className="font-sans text-xs sm:text-sm font-medium opacity-90">
                              Mata Pelajaran: {metadata.mataPelajaran} | Topik: {selectedAsesmenTopic}
                            </p>
                          </div>

                          <div className="flex flex-col gap-8 flex-1 text-slate-800 font-sans">
                            {/* Administration Identity Table */}
                            <div className="overflow-x-auto mb-2">
                              <table className="w-full border-collapse border border-slate-300 text-xs text-left">
                                <tbody>
                                  <tr>
                                    <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-[#007A54] w-[30%]">Nama Guru / Penyusun</td>
                                    <td className="border border-slate-300 p-2.5 text-slate-700 w-[70%] font-semibold">{asesmenList[selectedAsesmenTopic].identitas?.namaGuru || metadata.namaGuru || "Rudi Akbar Saragih, S.Kom."}</td>
                                  </tr>
                                  <tr>
                                    <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-[#007A54]">Satuan Pendidikan / Sekolah</td>
                                    <td className="border border-slate-300 p-2.5 text-slate-700">{asesmenList[selectedAsesmenTopic].identitas?.sekolah || metadata.jenjangSekolah || "SMK/MAK"}</td>
                                  </tr>
                                  <tr>
                                    <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-[#007A54]">Mata Pelajaran</td>
                                    <td className="border border-slate-300 p-2.5 text-slate-700">{asesmenList[selectedAsesmenTopic].identitas?.mataPelajaran || metadata.mataPelajaran}</td>
                                  </tr>
                                  <tr>
                                    <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-[#007A54]">Semester / Tahun Ajaran</td>
                                    <td className="border border-slate-300 p-2.5 text-slate-700">{asesmenList[selectedAsesmenTopic].identitas?.semesterTahun || `Ganjil - ${metadata.tahunAjaran || "2026/2027"}`}</td>
                                  </tr>
                                  <tr>
                                    <td className="border border-slate-300 p-2.5 font-bold bg-[#F0FDF4] text-[#007A54]">Topik / Bahasan Utama</td>
                                    <td className="border border-slate-300 p-2.5 text-slate-700 font-semibold">{asesmenList[selectedAsesmenTopic].identitas?.topikMateri || selectedAsesmenTopic}</td>
                                  </tr>
                                </tbody>
                              </table>
                            </div>

                            {/* 2. Tabel Kisi-Kisi */}
                            {asesmenList[selectedAsesmenTopic].kisiKisi && asesmenList[selectedAsesmenTopic].kisiKisi.length > 0 && (
                              <div className="flex flex-col gap-2">
                                <h3 className="font-bold text-sm text-emerald-800 flex items-center gap-1.5 uppercase tracking-wide">
                                  <span className="w-1.5 h-4 bg-emerald-600 rounded"></span>
                                  I. Kisi-Kisi Asesmen Pembelajaran
                                </h3>
                                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                  <table className="w-full text-xs text-left border-collapse">
                                    <thead>
                                      <tr className="bg-emerald-800 text-white font-bold">
                                        <th className="p-2 border-r border-emerald-700 text-center w-10">No</th>
                                        <th className="p-2 border-r border-emerald-700">Tujuan Pembelajaran</th>
                                        <th className="p-2 border-r border-emerald-700">Indikator Soal</th>
                                        <th className="p-2 border-r border-emerald-700 text-center">Level</th>
                                        <th className="p-2 border-r border-emerald-700 text-center">Bentuk</th>
                                        <th className="p-2 text-center w-12">No Soal</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-200">
                                      {asesmenList[selectedAsesmenTopic].kisiKisi.map((k, idx) => (
                                        <tr key={idx} className={idx % 2 === 0 ? "bg-white" : "bg-slate-50/50"}>
                                          <td className="p-2 border-r border-slate-200 text-center">{k.no}</td>
                                          <td className="p-2 border-r border-slate-200 leading-normal">{k.tujuanPembelajaran}</td>
                                          <td className="p-2 border-r border-slate-200 leading-normal">{k.indikatorSoal}</td>
                                          <td className="p-2 border-r border-slate-200 text-center text-[10px] font-medium text-slate-600">{k.levelKognitif}</td>
                                          <td className="p-2 border-r border-slate-200 text-center text-[10px] font-medium text-slate-600">{k.bentukSoal}</td>
                                          <td className="p-2 text-center font-bold text-slate-800">{k.noSoal}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            )}

                            {/* 3. Butir Soal Evaluasi */}
                            <div className="flex flex-col gap-4">
                              <h3 className="font-bold text-sm text-emerald-800 flex items-center gap-1.5 uppercase tracking-wide">
                                <span className="w-1.5 h-4 bg-emerald-600 rounded"></span>
                                II. Butir Soal Asesmen
                              </h3>
                              
                              <div className="flex flex-col gap-6">
                                {(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).map((s, idx) => {
                                  const isMCQ = s.tipe === "Pilihan Ganda" || (!s.tipe && (s.pilihan || s.a));
                                  return (
                                    <div key={idx} className="flex flex-col gap-3 bg-slate-50/40 p-5 rounded-xl border border-slate-200/60 shadow-sm font-sans">
                                      <div className="font-bold text-slate-900 leading-relaxed text-justify flex items-start gap-2 text-xs sm:text-sm">
                                        <span className="bg-slate-200/80 text-slate-800 px-2 py-0.5 rounded text-xs leading-none mt-0.5">{s.no || (idx + 1)}</span>
                                        <div className="flex-1">
                                          {s.soal || s.pertanyaan}
                                          <span className="text-[10px] bg-emerald-50 text-emerald-800 border border-emerald-100 font-semibold px-1.5 py-0.5 rounded ml-2 inline-block">
                                            {s.tipe || "Pilihan Ganda"} • {s.levelKognitif || "C1"}
                                          </span>
                                        </div>
                                      </div>

                                      {isMCQ && (
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-x-4 gap-y-2 pl-7 text-xs sm:text-sm text-slate-700">
                                          {s.pilihan ? (
                                            Object.keys(s.pilihan).map((opt) => (
                                              <div key={opt} className="flex items-start gap-1.5">
                                                <strong className="font-semibold shrink-0 text-emerald-700">{opt}.</strong>
                                                <span>{s.pilihan[opt]}</span>
                                              </div>
                                            ))
                                          ) : (
                                            ["A", "B", "C", "D", "E"].map((opt) => {
                                              const val = s[opt.toLowerCase()];
                                              if (!val) return null;
                                              return (
                                                <div key={opt} className="flex items-start gap-1.5">
                                                  <strong className="font-semibold shrink-0 text-emerald-700">{opt}.</strong>
                                                  <span>{val}</span>
                                                </div>
                                              );
                                            })
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            </div>

                            {/* 4. Tabel Kunci Jawaban & Rubrik Penilaian */}
                            <div className="flex flex-col gap-2">
                              <h3 className="font-bold text-sm text-emerald-800 flex items-center gap-1.5 uppercase tracking-wide">
                                <span className="w-1.5 h-4 bg-emerald-600 rounded"></span>
                                III. Kunci Jawaban & Rubrik Penilaian
                              </h3>
                              <div className="overflow-x-auto border border-slate-200 rounded-lg">
                                <table className="w-full text-xs text-left border-collapse">
                                  <thead>
                                    <tr className="bg-emerald-800 text-white font-bold">
                                      <th className="p-2 border-r border-emerald-700 text-center w-16">No Soal</th>
                                      <th className="p-2 border-r border-emerald-700 text-center w-28">Bentuk Soal</th>
                                      <th className="p-2 border-r border-emerald-700">Kunci Jawaban / Rubrik Penilaian</th>
                                      <th className="p-2 text-center w-20">Skor Max</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-200">
                                    {/* Bagian Pilihan Ganda */}
                                    {(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).some(s => s.tipe === "Pilihan Ganda" || (!s.tipe && (s.pilihan || s.a))) && (
                                      <>
                                        <tr className="bg-emerald-50 text-emerald-800 font-bold">
                                          <td colSpan={4} className="p-2">BAGIAN I: PILIHAN GANDA</td>
                                        </tr>
                                        {(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).filter(s => s.tipe === "Pilihan Ganda" || (!s.tipe && (s.pilihan || s.a))).map((s, idx) => (
                                          <tr key={idx} className="bg-white">
                                            <td className="p-2 border-r border-slate-200 text-center font-semibold">{s.no || (idx + 1)}</td>
                                            <td className="p-2 border-r border-slate-200 text-center text-slate-500">Pilihan Ganda</td>
                                            <td className="p-2 border-r border-slate-200 font-medium text-emerald-800 bg-emerald-50/10">{s.kunciJawaban}</td>
                                            <td className="p-2 text-center font-bold text-slate-600">{s.skorMaksimal || 5}</td>
                                          </tr>
                                        ))}
                                      </>
                                    )}

                                    {/* Bagian Uraian */}
                                    {(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).some(s => s.tipe === "Uraian" || (!s.tipe && !s.pilihan && !s.a)) && (
                                      <>
                                        <tr className="bg-emerald-50 text-emerald-800 font-bold border-t border-slate-200">
                                          <td colSpan={4} className="p-2">BAGIAN II: URAIAN</td>
                                        </tr>
                                        {(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).filter(s => s.tipe === "Uraian" || (!s.tipe && !s.pilihan && !s.a)).map((s, idx) => (
                                          <tr key={idx} className="bg-white">
                                            <td className="p-2 border-r border-slate-200 text-center font-semibold">{s.no || (idx + 1)}</td>
                                            <td className="p-2 border-r border-slate-200 text-center text-slate-500">Uraian</td>
                                            <td className="p-2 border-r border-slate-200 whitespace-pre-wrap leading-normal font-sans py-3">{s.kunciJawaban || s.pembahasan}</td>
                                            <td className="p-2 text-center font-bold text-slate-600">{s.skorMaksimal || 15}</td>
                                          </tr>
                                        ))}
                                      </>
                                    )}

                                    {/* Total Row */}
                                    <tr className="bg-slate-100 font-bold text-slate-800 border-t border-slate-300">
                                      <td colSpan={3} className="p-2 text-right uppercase tracking-wide">Total Skor Maksimal Keseluruhan</td>
                                      <td className="p-2 text-center text-emerald-700 font-extrabold text-sm">{(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).reduce((sum, s) => sum + (s.skorMaksimal || 5), 0)}</td>
                                    </tr>
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          </div>
                        </div>
                      ) : (
                        /* Edit Panel */
                        <div className="p-6 flex flex-col gap-6 font-sans">
                          {(asesmenList[selectedAsesmenTopic].daftarSoal || asesmenList[selectedAsesmenTopic].soal || []).map((s, idx) => {
                            const isMCQ = s.tipe === "Pilihan Ganda" || (!s.tipe && (s.pilihan || s.a));
                            return (
                              <div key={idx} className="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col gap-3">
                                <div className="flex justify-between items-center border-b border-slate-200 pb-1">
                                  <span className="text-[10px] font-bold text-slate-400">SOAL NOMOR {idx + 1} ({s.tipe || "Pilihan Ganda"})</span>
                                  <span className="text-[10px] font-bold text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded">Skor Max: {s.skorMaksimal || (isMCQ ? 5 : 15)}</span>
                                </div>
                                <div className="flex flex-col gap-1">
                                  <span className="text-[9px] font-bold text-slate-400">PERTANYAAN</span>
                                  <textarea 
                                    rows={2}
                                    value={s.soal || s.pertanyaan}
                                    onChange={(e) => {
                                      const updated = { ...asesmenList };
                                      const target = updated[selectedAsesmenTopic];
                                      if (target.daftarSoal) target.daftarSoal[idx].soal = e.target.value;
                                      if (target.soal) target.soal[idx].pertanyaan = e.target.value;
                                      setAsesmenList(updated);
                                    }}
                                    className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                                  />
                                </div>
                                
                                {isMCQ && (
                                  <div className="grid grid-cols-1 sm:grid-cols-5 gap-3">
                                    {["A", "B", "C", "D", "E"].map((opt) => {
                                      const val = s.pilihan ? s.pilihan[opt] : s[opt.toLowerCase()];
                                      return (
                                        <div key={opt} className="flex flex-col gap-0.5">
                                          <span className="text-[9px] font-bold text-slate-400">{opt}</span>
                                          <input 
                                            type="text" 
                                            value={val || ""}
                                            onChange={(e) => {
                                              const updated = { ...asesmenList };
                                              const target = updated[selectedAsesmenTopic];
                                              if (target.daftarSoal) {
                                                if (!target.daftarSoal[idx].pilihan) target.daftarSoal[idx].pilihan = { A: "", B: "", C: "", D: "", E: "" };
                                                target.daftarSoal[idx].pilihan[opt] = e.target.value;
                                              }
                                              if (target.soal) {
                                                target.soal[idx][opt.toLowerCase()] = e.target.value;
                                              }
                                              setAsesmenList(updated);
                                            }}
                                            className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs"
                                          />
                                        </div>
                                      );
                                    })}
                                  </div>
                                )}

                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[9px] font-bold text-slate-400">{isMCQ ? "KUNCI JAWABAN (A-E)" : "KUNCI JAWABAN & RUBRIK PENILAIAN"}</span>
                                    {isMCQ ? (
                                      <input 
                                        type="text" 
                                        value={s.kunciJawaban}
                                        onChange={(e) => {
                                          const updated = { ...asesmenList };
                                          const target = updated[selectedAsesmenTopic];
                                          if (target.daftarSoal) target.daftarSoal[idx].kunciJawaban = e.target.value;
                                          if (target.soal) target.soal[idx].kunciJawaban = e.target.value;
                                          setAsesmenList(updated);
                                        }}
                                        className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-bold"
                                      />
                                    ) : (
                                      <textarea 
                                        rows={3}
                                        value={s.kunciJawaban || s.pembahasan}
                                        onChange={(e) => {
                                          const updated = { ...asesmenList };
                                          const target = updated[selectedAsesmenTopic];
                                          if (target.daftarSoal) target.daftarSoal[idx].kunciJawaban = e.target.value;
                                          if (target.soal) target.soal[idx].pembahasan = e.target.value;
                                          setAsesmenList(updated);
                                        }}
                                        className="bg-white border border-slate-300 rounded px-2.5 py-1 text-xs"
                                      />
                                    )}
                                  </div>
                                  <div className="flex flex-col gap-0.5 justify-end">
                                    <span className="text-[9px] font-bold text-slate-400">LEVEL KOGNITIF / INDIKATOR</span>
                                    <input 
                                      type="text" 
                                      value={s.levelKognitif || s.kisiKisi || ""}
                                      onChange={(e) => {
                                        const updated = { ...asesmenList };
                                        const target = updated[selectedAsesmenTopic];
                                        if (target.daftarSoal) target.daftarSoal[idx].levelKognitif = e.target.value;
                                        if (target.soal) target.soal[idx].kisiKisi = e.target.value;
                                        setAsesmenList(updated);
                                      }}
                                      className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs italic"
                                    />
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  ) : (
                    /* Empty State Block */
                    <div className="flex-1 flex flex-col items-center justify-center p-8 text-center gap-4">
                      <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center text-slate-400 border border-slate-200">
                        <Award className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-bold text-slate-800 text-sm">Soal Asesmen Belum Dibuat</h4>
                        <p className="text-slate-500 text-[10px] sm:text-xs mt-1 max-w-sm">Daftar evaluasi pilihan ganda 5 opsi untuk materi ini belum digenerate. Klik tombol di bawah untuk membuat secara instan.</p>
                      </div>
                      <button
                        onClick={() => {
                          const t = topikData.daftarTopik.find(x => x.topik === selectedAsesmenTopic);
                          if (t) generateAsesmenSingle(t);
                        }}
                        className="bg-slate-900 hover:bg-slate-800 text-white font-bold py-2 px-4 rounded text-xs transition-colors cursor-pointer"
                      >
                        Generate Asesmen Sekarang
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Navigation Controls */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-100 p-4 rounded-xl border border-slate-200">
                <button
                  onClick={() => setStep(7)}
                  className="bg-white hover:bg-slate-50 text-slate-700 font-bold py-2.5 px-5 rounded border border-slate-300 shadow-sm transition-all flex items-center gap-1.5 text-sm cursor-pointer w-full sm:w-auto justify-center"
                >
                  <ArrowLeft className="w-4 h-4" />
                  <span>Kembali</span>
                </button>

                <div className="flex items-center flex-wrap gap-2.5 w-full sm:w-auto">
                  <button
                    onClick={() => {
                      const t = topikData.daftarTopik.find(x => x.topik === selectedAsesmenTopic);
                      if (t) generateAsesmenSingle(t);
                    }}
                    disabled={!asesmenList[selectedAsesmenTopic]}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-2.5 px-4 rounded transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <RefreshCw className="w-4 h-4 text-slate-600" />
                    <span>Buat Ulang</span>
                  </button>

                  <button
                    onClick={downloadAllAsZip}
                    className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold py-2.5 px-5 rounded shadow transition-all flex items-center justify-center gap-1.5 text-sm cursor-pointer flex-1 hover:shadow-md active:scale-[0.98]"
                  >
                    <Download className="w-4 h-4" />
                    <span>Unduh Semua (.ZIP)</span>
                  </button>
                </div>
              </div>
            </motion.div>
          )}

        </div>

      </div>

        {/* Informative Help Sidebar or Footer block */}
        <div className="mx-4 md:mx-6 mb-4 md:mb-6 bg-white rounded-xl p-4 border border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4 shrink-0 shadow-sm">
          <div className="flex items-center gap-3">
            <Layers className="w-6 h-6 text-blue-600 shrink-0" />
            <div>
              <h4 className="font-bold text-xs sm:text-sm text-slate-800">Alur Penyusunan Kurikulum Merdeka Terstruktur</h4>
              <p className="text-slate-500 text-[10px] sm:text-xs mt-0.5">Pembuatan dokumen dilakukan berurutan: CP → Analisis CP & ATP → Prota & Promes → KKTP → Rencana Pembelajaran (RPM).</p>
            </div>
          </div>
          <div className="text-[10px] sm:text-xs text-slate-400 font-medium">
            Dikembangkan sesuai regulasi resmi Kemendikbudristek RI.
          </div>
        </div>

      </main>
    </div>

    {/* Footer credit */}
    <footer className="bg-white border-t border-slate-200 text-slate-400 text-center py-3 text-xs shrink-0">
      <p>© 2026 SAGO (Sistem Administrasi Guru Otomatis) • Berpedoman pada BSKAP 046/2025 & BKPDM 020/2026.</p>
    </footer>
  </div>
  );
}
