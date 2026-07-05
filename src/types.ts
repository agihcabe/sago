export interface TeacherMetadata {
  namaGuru: string;
  nipGuru: string;
  namaKepalaSekolah: string;
  nipKepalaSekolah: string;
  mataPelajaran: string;
  jenjangSekolah: string;
  fase: string;
  tahunAjaran: string;
}

export interface CPElemen {
  nama: string;
  deskripsi: string;
  capaian: string;
}

export interface CPData {
  judul: string;
  regulasi: string;
  rasional: string;
  tujuan: string;
  karakteristik: string;
  capaianUmum: string;
  elemen: CPElemen[];
}

export interface TPEntry {
  kode: string;
  tp: string;
  topik: string;
  alokasiWaktu: string;
  profilLulusan: string;
  glosarium: string;
}

export interface ATPElemen {
  elemen: string;
  capaian?: string;
  tujuanPembelajaran: TPEntry[];
}

export interface ATPData {
  judul: string;
  fase: string;
  mataPelajaran: string;
  atp: ATPElemen[];
}

export interface ProtaItem {
  no: number;
  elemen: string;
  tujuanPembelajaran: string;
  alokasiWaktu: string;
}

export interface PromesItem {
  no: number;
  tujuanPembelajaran: string;
  alokasiWaktu: number;
  distribusi: {
    Jul: string;
    Agt: string;
    Sep: string;
    Okt: string;
    Nov: string;
    Des: string;
    Jan: string;
    Feb: string;
    Mar: string;
    Apr: string;
    Mei: string;
    Jun: string;
    [key: string]: string;
  };
}

export interface ProtaPromesData {
  judulProta: string;
  judulPromes: string;
  prota: ProtaItem[];
  promes: PromesItem[];
  jumlahJpSemester1: number;
  jumlahJpSemester2: number;
}

export interface KKTPCriteria {
  aspek: string;
  baruBerkembang: string;
  layak: string;
  cakap: string;
  mahir: string;
}

export interface KKTPItem {
  tpKode: string;
  tpTeks: string;
  kriteria: KKTPCriteria[];
  intervalNilai: {
    kurang: string;
    cukup: string;
    baik: string;
    sangatBaik: string;
  };
}

export interface KKTPData {
  judul: string;
  kktp: KKTPItem[];
}

export interface RPMAsesmenItem {
  metode: string;
  instrumen: string;
  kriteria: string;
}

export interface RPMKegiatanFase {
  durasi: string;
  guru: string[];
  siswa: string[];
  pendekatan: string;
}

export interface RPMData {
  judul: string;
  identitas: {
    namaGuru: string;
    sekolah: string;
    mataPelajaran: string;
    kelasSemester: string;
    semesterTahun: string;
    topik: string;
    modelPembelajaran: string;
  };
  identifikasi: {
    targetPesertaDidik: string;
    kompetensiPrasyarat: string[];
    profilLulusan: string[];
    saranaPrasarana: {
      fasilitasFisik: string;
      pirantiLunak: string;
      sumberBelajar: string;
    };
  };
  desainPembelajaran: {
    capaianPembelajaran: string;
    lintasDisiplin: string[];
    tujuanPembelajaran: string[];
    topikPembelajaran: string;
    praktikPedagogis: string;
    kemitraanPembelajaran: string;
    lingkunganBelajar: string;
    pemanfaatanDigital: string;
  };
  langkahPembelajaran: {
    pendahuluan: RPMKegiatanFase;
    intiEksplorasi: RPMKegiatanFase;
    intiRancangan: RPMKegiatanFase;
    penutup: RPMKegiatanFase;
  };
  asesmen: {
    diagnostik: RPMAsesmenItem;
    formatif: RPMAsesmenItem;
    sumatif: RPMAsesmenItem;
  };
  remedialPengayaan: {
    remedial: {
      sasaran: string;
      aktivitas: string[];
    };
    pengayaan: {
      sasaran: string;
      aktivitas: string[];
    };
  };
  refleksi: {
    siswa: string[];
    guru: string[];
  };
}

export interface TopikItem {
  no: number;
  topik: string;
  deskripsi: string;
  tpKode: string;
  tpTeks: string;
  alokasiWaktu: string;
  semester: number;
}

export interface TopikData {
  judul: string;
  daftarTopik: TopikItem[];
}

export interface LKPDRingkasan {
  komponen: string;
  ringkasan: string;
}

export interface LKPDActivity {
  judul: string;
  studiKasus: string;
}

export interface LKPDData {
  topik: string;
  tpKode: string;
  judul: string;
  identitas: {
    sekolah: string;
    mataPelajaran: string;
    kelasSemester: string;
    alokasiWaktu: string;
    namaGuru: string;
    semesterTahun: string;
    topikMateri: string;
    modelPembelajaran: string;
  };
  tujuanPembelajaran: string[];
  petunjukPengerjaan: string[];
  ringkasanMateri: LKPDRingkasan[];
  aktivitasList: LKPDActivity[];
  rubrikPenilaian?: string;
}

export interface AsesmenKisiKisi {
  no: number;
  tujuanPembelajaran: string;
  indikatorSoal: string;
  levelKognitif: string;
  bentukSoal: string;
  noSoal: number;
}

export interface AsesmenItem {
  no: number;
  tipe?: string; // "Pilihan Ganda" atau "Uraian"
  levelKognitif?: string;
  soal: string;
  pilihan?: {
    A: string;
    B: string;
    C: string;
    D: string;
    E: string;
  };
  kunciJawaban: string;
  pembahasan?: string;
  skorMaksimal?: number;
  
  // Legacy / Edit Mode support
  pertanyaan?: string;
  kisiKisi?: string;
  a?: string;
  b?: string;
  c?: string;
  d?: string;
  e?: string;
}

export interface AsesmenData {
  topik: string;
  tpKode: string;
  judul: string;
  kisiKisiUmum: string;
  identitas?: {
    namaGuru: string;
    sekolah: string;
    mataPelajaran: string;
    semesterTahun: string;
    topikMateri: string;
    modelPembelajaran: string;
  };
  kisiKisi?: AsesmenKisiKisi[];
  daftarSoal: AsesmenItem[];
  soal?: any[]; // Legacy backward compatibility
}
