import { 
  Document, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, HeadingLevel, AlignmentType, Packer, BorderStyle,
  HeightRule, VerticalMergeType
} from "docx";

interface Metadata {
  namaGuru: string;
  nipGuru: string;
  namaKepalaSekolah: string;
  nipKepalaSekolah: string;
  mataPelajaran: string;
  jenjangSekolah: string;
  fase: string;
  tahunAjaran: string;
}

// Common border style helper for professional looking tables
const commonBorder = {
  top: { style: BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
  bottom: { style: BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
  left: { style: BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
  right: { style: BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
};

// Common header styling
function createDocHeader(title: string, meta: Metadata): Paragraph[] {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: title.toUpperCase(),
          bold: true,
          size: 28, // 14pt
          font: "Arial",
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [
        new TextRun({
          text: `MATA PELAJARAN: ${meta.mataPelajaran.toUpperCase()} | FASE ${meta.fase.toUpperCase()}`,
          bold: true,
          size: 22, // 11pt
          font: "Arial",
          color: "555555",
        }),
      ],
    }),
    // Metadata Table / Block
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: "IDENTITAS ADMINISTRATIF", bold: true, size: 18, font: "Arial", color: "444444" })
      ]
    }),
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: `Nama Guru : `, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: meta.namaGuru || "-", size: 20, font: "Arial" }),
        new TextRun({ text: `   |   NIP : `, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: meta.nipGuru || "-", size: 20, font: "Arial" }),
      ]
    }),
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: `Sekolah / Jenjang : `, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: `${meta.jenjangSekolah || "-"} (Fase ${meta.fase || "-"})`, size: 20, font: "Arial" }),
        new TextRun({ text: `   |   Tahun Ajaran : `, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: meta.tahunAjaran || "-", size: 20, font: "Arial" }),
      ]
    }),
    new Paragraph({
      spacing: { after: 240 },
      children: [
        new TextRun({ text: `Kepala Sekolah : `, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: meta.namaKepalaSekolah || "-", size: 20, font: "Arial" }),
        new TextRun({ text: `   |   NIP Kepsek : `, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: meta.nipKepalaSekolah || "-", size: 20, font: "Arial" }),
      ]
    }),
    new Paragraph({
      spacing: { after: 240 },
      children: [
        new TextRun({
          text: "_________________________________________________________________________________",
          color: "CCCCCC",
        })
      ]
    })
  ];
}

// Common signature helper for Indonesian administrative documents
function createSignatures(meta: Metadata): any[] {
  return [
    new Paragraph({ spacing: { before: 400, after: 120 } }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: "Mengetahui,", font: "Arial", size: 20 })],
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: "Kepala Sekolah", font: "Arial", size: 20 })],
                  spacing: { after: 1200 }, // Space for signature
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: `( ${meta.namaKepalaSekolah || "__________________________"} )`, bold: true, font: "Arial", size: 20 })],
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: `NIP. ${meta.nipKepalaSekolah || "..................................."}`, font: "Arial", size: 20 })],
                }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: "......................, 29 Juni 2026", font: "Arial", size: 20 })],
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: "Guru Mata Pelajaran", font: "Arial", size: 20 })],
                  spacing: { after: 1200 }, // Space for signature
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: meta.namaGuru || "__________________________", bold: true, font: "Arial", size: 20 })],
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: `NIP. ${meta.nipGuru || "..................................."}`, font: "Arial", size: 20 })],
                }),
              ],
            }),
          ],
        }),
      ],
    })
  ];
}

// 1. CAPAIAN PEMBELAJARAN (CP)
export function generateCPDoc(meta: Metadata, data: any): Document {
  const children: any[] = [
    ...createDocHeader("Capaian Pembelajaran (CP)", meta),
    
    // Regulasi rujukan
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: "Regulasi Rujukan: ",
          bold: true,
          size: 22,
          font: "Arial",
        }),
        new TextRun({
          text: data.regulasi || "Keputusan Kepala BSKAP Nomor 046/H/KR/2025",
          italics: true,
          size: 22,
          font: "Arial",
        })
      ]
    }),

    // Rasional
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "1. Rasional Mata Pelajaran", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new Paragraph({
      spacing: { after: 180 },
      children: [new TextRun({ text: data.rasional || "", size: 22, font: "Arial" })]
    }),

    // Tujuan
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "2. Tujuan Mata Pelajaran", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new Paragraph({
      spacing: { after: 180 },
      children: [new TextRun({ text: data.tujuan || "", size: 22, font: "Arial" })]
    }),

    // Karakteristik
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "3. Karakteristik Mata Pelajaran", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new Paragraph({
      spacing: { after: 180 },
      children: [new TextRun({ text: data.karakteristik || "", size: 22, font: "Arial" })]
    }),

    // Capaian Umum
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: `4. Capaian Umum Fase ${meta.fase.toUpperCase()}`, bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new Paragraph({
      spacing: { after: 180 },
      children: [new TextRun({ text: data.capaianUmum || "", size: 22, font: "Arial" })]
    }),

    // Capaian per Elemen (Table)
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "5. Capaian Pembelajaran per Elemen", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
  ];

  // Table of elements
  const tableRows = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 30, type: WidthType.PERCENTAGE },
          shading: { fill: "F2F2F2" },
          borders: commonBorder,
          children: [new Paragraph({ children: [new TextRun({ text: "Elemen", bold: true, font: "Arial", size: 20 })] })],
        }),
        new TableCell({
          width: { size: 70, type: WidthType.PERCENTAGE },
          shading: { fill: "F2F2F2" },
          borders: commonBorder,
          children: [new Paragraph({ children: [new TextRun({ text: "Capaian Pembelajaran Elemen", bold: true, font: "Arial", size: 20 })] })],
        }),
      ],
    })
  ];

  if (Array.isArray(data.elemen)) {
    data.elemen.forEach((el: any) => {
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({
              borders: commonBorder,
              children: [
                new Paragraph({
                  children: [new TextRun({ text: el.nama || "", bold: true, font: "Arial", size: 20 })],
                  spacing: { after: 60 }
                }),
                new Paragraph({
                  children: [new TextRun({ text: el.deskripsi || "", italics: true, font: "Arial", size: 18, color: "555555" })]
                })
              ],
            }),
            new TableCell({
              borders: commonBorder,
              children: [
                new Paragraph({
                  children: [new TextRun({ text: el.capaian || "", font: "Arial", size: 20 })]
                })
              ],
            }),
          ],
        })
      );
    });
  }

  children.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
  children.push(...createSignatures(meta));

  return new Document({
    sections: [{ properties: {}, children }],
  });
}

// 2. ANALISIS CP & ATP
export function generateATPDoc(meta: Metadata, data: any): Document {
  const children: any[] = [];

  // Helper to format TP text with numbers cleanly
  const formatTpText = (tpText: string, idx: number) => {
    const trimmed = (tpText || "").trim();
    if (/^\d+[\.\)]/.test(trimmed)) {
      return trimmed;
    }
    return `${idx + 1}. ${trimmed}`;
  };

  // 1. Beautiful Header Banner Block
  const bannerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "007A54" }, // Deep green from the PDF image
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new TextRun({
                    text: "CAPAIAN & ALUR TUJUAN PEMBELAJARAN",
                    bold: true,
                    color: "FFFFFF",
                    size: 28, // 14pt
                    font: "Arial",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran} | Fase/Kelas: ${meta.fase} / XI`,
                    bold: true,
                    color: "FFFFFF",
                    size: 20, // 10pt
                    font: "Arial",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(bannerTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // 2. Identity Table below the banner
  const identityTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" }, // light mint green
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Nama Guru / Penyusun", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: meta.namaGuru || "", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Instansi / Sekolah", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Mata Pelajaran", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: meta.mataPelajaran || "", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Jenjang / Fase / Kelas", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: `SMK/MAK / ${meta.fase} / XI`, font: "Arial", size: 18 })] })],
          }),
        ],
      }),
    ],
  });

  children.push(identityTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // 3. Heading: CAPAIAN PEMBELAJARAN
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "CAPAIAN PEMBELAJARAN", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );

  // 4. Table 1: Capaian Pembelajaran per Elemen
  const cpTableRows = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 30, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ children: [new TextRun({ text: "Elemen", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 70, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ children: [new TextRun({ text: "Capaian Pembelajaran (CP)", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })],
        }),
      ],
    })
  ];

  if (Array.isArray(data.atp)) {
    data.atp.forEach((item: any) => {
      cpTableRows.push(
        new TableRow({
          children: [
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              shading: { fill: "F9FBF9" },
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18, color: "111111" })]
                })
              ],
            }),
            new TableCell({
              width: { size: 70, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: item.capaian || item.deskripsi || "", font: "Arial", size: 18 })]
                })
              ],
            }),
          ],
        })
      );
    });
  }

  children.push(new Table({ rows: cpTableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // 5. Page Break and Heading: ALUR TUJUAN PEMBELAJARAN (ATP)
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      pageBreakBefore: true,
      children: [new TextRun({ text: "ALUR TUJUAN PEMBELAJARAN (ATP)", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );

  // 6. Table 2: Alur Tujuan Pembelajaran
  const atpTableRows = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 15, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Elemen", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Capaian Pembelajaran (CP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 32, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tujuan Pembelajaran (TP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 13, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Kata Kunci Materi", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 10, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Profil Lulusan (SKL)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Perkiraan JP", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
      ],
    })
  ];

  if (Array.isArray(data.atp)) {
    data.atp.forEach((item: any) => {
      const tps = Array.isArray(item.tujuanPembelajaran) ? item.tujuanPembelajaran : [];
      const rowSpanVal = tps.length || 1;

      if (tps.length === 0) {
        atpTableRows.push(
          new TableRow({
            children: [
              new TableCell({
                width: { size: 15, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18 })] })],
              }),
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: item.capaian || "", font: "Arial", size: 18 })] })],
              }),
              new TableCell({
                width: { size: 32, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "", font: "Arial", size: 18 })] })],
              }),
              new TableCell({
                width: { size: 13, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "", font: "Arial", size: 18 })] })],
              }),
              new TableCell({
                width: { size: 10, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "", font: "Arial", size: 16 })] })],
              }),
              new TableCell({
                width: { size: 5, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "", font: "Arial", size: 18 })] })],
              }),
            ],
          })
        );
      } else {
        tps.forEach((tpObj: any, index: number) => {
          const cells = [];

          if (index === 0) {
            cells.push(
              new TableCell({
                width: { size: 15, type: WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18 })] })],
              })
            );
            cells.push(
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: item.capaian || "", font: "Arial", size: 18 })] })],
              })
            );
          }

          const formattedTp = formatTpText(tpObj.tp, index);
          cells.push(
            new TableCell({
              width: { size: 32, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: formattedTp, font: "Arial", size: 18 })
                  ]
                })
              ],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 13, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: tpObj.topik || "", font: "Arial", size: 18 })] })],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 10, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: tpObj.profilLulusan || tpObj.profilPancasila || "", font: "Arial", size: 16 })] })],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: tpObj.alokasiWaktu || "", font: "Arial", size: 18 })] })],
            })
          );

          atpTableRows.push(new TableRow({ children: cells }));
        });
      }
    });
  }

  children.push(new Table({ rows: atpTableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
  children.push(...createSignatures(meta));

  return new Document({
    sections: [{ properties: {}, children }],
  });
}

// 3. PROTA & PROMES
export function generateProtaPromesDoc(meta: Metadata, data: any): Document {
  const children: any[] = [];

  // Create a beautiful header banner block
  const bannerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "007A54" }, // Deep green
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new TextRun({
                    text: "PROGRAM TAHUNAN & PROGRAM SEMESTER",
                    bold: true,
                    color: "FFFFFF",
                    size: 28, // 14pt
                    font: "Arial",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran} | Fase/Kelas: ${meta.fase} / XI`,
                    bold: true,
                    color: "FFFFFF",
                    size: 20, // 10pt
                    font: "Arial",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(bannerTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // Identity Table below the banner
  const identityTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: "Nama Guru / Penyusun", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: meta.namaGuru || "-", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: "Instansi / Sekolah", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: meta.jenjangSekolah || "-", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: "Mata Pelajaran", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: meta.mataPelajaran || "-", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: "Jenjang / Fase / Kelas", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            borders: commonBorder,
            children: [new Paragraph({ children: [new TextRun({ text: `${meta.jenjangSekolah || "SMK/MAK"} / ${meta.fase} / XI`, font: "Arial", size: 18 })] })],
          }),
        ],
      }),
    ],
  });

  children.push(identityTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // Part A: PROTA Title
  const protaHeading = new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [
      new TextRun({
        text: "PROGRAM TAHUNAN (PROTA)",
        bold: true,
        color: "007A54",
        size: 24, // 12pt
        font: "Arial",
      }),
    ],
  });
  children.push(protaHeading);

  // PROTA Table Rows
  const protaRows = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 8, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ children: [new TextRun({ text: "Elemen / Lingkup Materi", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          width: { size: 52, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ children: [new TextRun({ text: "Tujuan Pembelajaran", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          width: { size: 15, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Alokasi Waktu (JP)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
      ],
    }),
  ];

  let totalJp = 0;
  if (Array.isArray(data.prota)) {
    data.prota.forEach((item: any, idx: number) => {
      const numJp = parseInt(item.alokasiWaktu) || 0;
      totalJp += numJp;
      protaRows.push(
        new TableRow({
          children: [
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(item.no || (idx + 1)), font: "Arial", size: 18 })] })],
            }),
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ children: [new TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18 })] })],
            }),
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ children: [new TextRun({ text: item.tujuanPembelajaran || "", font: "Arial", size: 18 })] })],
            }),
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.alokasiWaktu || "", font: "Arial", size: 18 })] })],
            }),
          ],
        })
      );
    });
  }

  // Add the total row for PROTA
  protaRows.push(
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 3,
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: "TOTAL ALOKASI WAKTU SATU TAHUN", bold: true, font: "Arial", size: 18 })],
            }),
          ],
        }),
        new TableCell({
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `${totalJp} JP`, bold: true, color: "007A54", font: "Arial", size: 18 })],
            }),
          ],
        }),
      ],
    })
  );

  const protaTable = new Table({ rows: protaRows, width: { size: 100, type: WidthType.PERCENTAGE } });
  children.push(protaTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // Part B: PROSEM Title
  const promesHeading = new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [
      new TextRun({
        text: "PROGRAM SEMESTER (PROSEM)",
        bold: true,
        color: "007A54",
        size: 24, // 12pt
        font: "Arial",
      }),
    ],
  });
  children.push(promesHeading);

  // PROSEM Table Header Row 1 & 2
  const promesRows = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          rowSpan: 2,
          width: { size: 5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          rowSpan: 2,
          width: { size: 40, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ children: [new TextRun({ text: "Tujuan Pembelajaran", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          rowSpan: 2,
          width: { size: 8, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Alokasi (JP)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          columnSpan: 6,
          width: { size: 23.5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Semester 1 (Ganjil)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
        new TableCell({
          columnSpan: 6,
          width: { size: 23.5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Semester 2 (Genap)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })],
        }),
      ],
    }),
    new TableRow({
      tableHeader: true,
      children: [
        ...["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => {
          return new TableCell({
            width: { size: 3.91, type: WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m, bold: true, color: "FFFFFF", font: "Arial", size: 16 })] })],
          });
        }),
      ],
    }),
  ];

  // Map promes items
  if (Array.isArray(data.promes)) {
    data.promes.forEach((item: any, idx: number) => {
      const distCells = ["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => {
        const val = (item.distribusi && item.distribusi[m]) || "-";
        const isNum = !isNaN(parseInt(val));
        const cellShading = isNum 
          ? { fill: "E6F0FA" } // Light sky blue
          : val === "ASG" || val === "ASE" 
            ? { fill: "F2F2F2" } // Light grey
            : undefined;

        return new TableCell({
          shading: cellShading,
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: val,
                  bold: isNum || val === "ASG" || val === "ASE",
                  font: "Arial",
                  size: 16,
                  color: isNum ? "1A365D" : val === "ASG" || val === "ASE" ? "4A5568" : "718096",
                }),
              ],
            }),
          ],
        });
      });

      promesRows.push(
        new TableRow({
          children: [
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(item.no || (idx + 1)), font: "Arial", size: 18 })] })],
            }),
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ children: [new TextRun({ text: item.tujuanPembelajaran || "", font: "Arial", size: 18 })] })],
            }),
            new TableCell({
              borders: commonBorder,
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(item.alokasiWaktu || ""), font: "Arial", size: 18 })] })],
            }),
            ...distCells,
          ],
        })
      );
    });
  }

  // Add total row under PROSEM
  const totalPromesJp = Array.isArray(data.promes) ? data.promes.reduce((sum: number, item: any) => sum + (Number(item.alokasiWaktu) || 0), 0) : 0;
  promesRows.push(
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 2,
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: "Jumlah JP per Semester", bold: true, font: "Arial", size: 18 })],
            }),
          ],
        }),
        new TableCell({
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `${totalPromesJp}`, bold: true, font: "Arial", size: 18 })],
            }),
          ],
        }),
        new TableCell({
          columnSpan: 6,
          shading: { fill: "EBF5FA" },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `Semester Ganjil: ${data.jumlahJpSemester1 || 172} JP`, bold: true, color: "007A54", font: "Arial", size: 18 })],
            }),
          ],
        }),
        new TableCell({
          columnSpan: 6,
          shading: { fill: "EBF5FA" },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `Semester Genap: ${data.jumlahJpSemester2 || 188} JP`, bold: true, color: "007A54", font: "Arial", size: 18 })],
            }),
          ],
        }),
      ],
    }),
  );

  const promesTable = new Table({ rows: promesRows, width: { size: 100, type: WidthType.PERCENTAGE } });
  children.push(promesTable);

  const promesFooterNote = new Paragraph({
    spacing: { before: 120, after: 240 },
    children: [
      new TextRun({
        text: "*Keterangan: ASG = Asesmen Sumatif Ganjil / Cadangan; ASE = Asesmen Sumatif Akhir Tahun / Kelulusan Fase. Angka menunjukkan distribusi alokasi waktu JP per bulan efektif.",
        italics: true,
        size: 16,
        font: "Arial",
        color: "718096",
      }),
    ],
  });
  children.push(promesFooterNote);

  // Add side-by-side signatures helper call
  const sigTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                spacing: { before: 240 },
                children: [new TextRun({ text: "Mengetahui,", font: "Arial", size: 20 })]
              }),
              new Paragraph({
                children: [new TextRun({ text: "Kepala Sekolah", bold: true, font: "Arial", size: 20 })]
              }),
              new Paragraph({
                spacing: { before: 1200 }, // Signature space
                children: [new TextRun({ text: `( ${meta.namaKepalaSekolah || "__________________________"} )`, bold: true, font: "Arial", size: 20 })]
              }),
              new Paragraph({
                children: [new TextRun({ text: `NIP. ${meta.nipKepalaSekolah || "..................................."}`, font: "Arial", size: 20 })]
              }),
            ]
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                spacing: { before: 240 },
                children: [new TextRun({ text: "......................, 29 Juni 2026", font: "Arial", size: 20 })]
              }),
              new Paragraph({
                children: [new TextRun({ text: "Guru Mata Pelajaran", bold: true, font: "Arial", size: 20 })]
              }),
              new Paragraph({
                spacing: { before: 1200 }, // Signature space
                children: [new TextRun({ text: meta.namaGuru || "__________________________", bold: true, font: "Arial", size: 20 })]
              }),
              new Paragraph({
                children: [new TextRun({ text: `NIP. ${meta.nipGuru || "..................................."}`, font: "Arial", size: 20 })]
              }),
            ]
          })
        ]
      })
    ]
  });

  children.push(sigTable);

  return new Document({
    sections: [{ properties: {}, children }],
  });
}

// 4. KKTP
export function generateKKTPDoc(meta: Metadata, data: any): Document {
  const children: any[] = [];

  // 1. Header Banner
  const bannerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "007A54" }, // Deep green from the PDF image
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new TextRun({
                    text: "ANALISIS KRITERIA KETERCAPAIAN (KKTP)",
                    bold: true,
                    color: "FFFFFF",
                    size: 28, // 14pt
                    font: "Arial",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran} | Fase/Kelas: ${meta.fase} / XI`,
                    bold: true,
                    color: "FFFFFF",
                    size: 20, // 10pt
                    font: "Arial",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(bannerTable);
  children.push(new Paragraph({ spacing: { before: 120, after: 120 } }));

  // 2. Identity Table
  const identityTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" }, // light mint green
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Nama Guru / Penyusun", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: meta.namaGuru || "", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Instansi / Sekolah", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Mata Pelajaran", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: meta.mataPelajaran || "", font: "Arial", size: 18 })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Jenjang / Fase / Kelas", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: `SMK/MAK / ${meta.fase} / XI`, font: "Arial", size: 18 })] })],
          }),
        ],
      }),
    ],
  });

  children.push(identityTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // 3. Heading: PENDEKATAN KKTP
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "PENDEKATAN KKTP", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );

  // 4. Table 1: Pendekatan KKTP
  const pendekatanTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Metode Pendekatan", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Deskripsi dan Alasan Penggunaan", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Pendekatan Interval Nilai", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Menggunakan interval kriteria untuk menentukan ketuntasan belajar peserta didik. Pendekatan ini dipilih karena memberikan gambaran objektif, terukur, dan mempermudah pemetaan tingkat kompetensi peserta didik (Mulai Berkembang, Layak, Cakap, Mahir) pada setiap indikator mata pelajaran ${meta.mataPelajaran} Fase ${meta.fase} (Kelas XI).`,
                    font: "Arial",
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Ketentuan Ketuntasan", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Peserta didik dinyatakan mencapai ketuntasan minimum jika mencapai kriteria minimal Layak (61-70) pada indikator-indikator kunci kompetensi esensial.",
                    font: "Arial",
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(pendekatanTable);
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // 5. Heading: ANALISIS KKTP
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "ANALISIS KKTP", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );

  // 6. Table 2: Analisis KKTP (7 columns: No, TP, IKTP, Mulai Berkembang, Layak, Cakap, Mahir)
  const kktpTableRows = [
    // Header Row 1
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          rowSpan: 2,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          rowSpan: 2,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tujuan Pembelajaran (TP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          rowSpan: 2,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Indikator Ketercapaian (IKTP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 50, type: WidthType.PERCENTAGE },
          columnSpan: 4,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Kriteria Interval Nilai (%)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
        }),
      ],
    }),
    // Header Row 2
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 12.5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Mulai Berkembang (0-60)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 12.5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Layak (61-70)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 12.5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Cakap (71-80)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })],
        }),
        new TableCell({
          width: { size: 12.5, type: WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Mahir (81-100)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })],
        }),
      ],
    }),
  ];

  if (Array.isArray(data.kktp)) {
    data.kktp.forEach((item: any, idx: number) => {
      const criteriaList = Array.isArray(item.kriteria) ? item.kriteria : [];
      const rowSpanVal = criteriaList.length || 1;

      if (criteriaList.length === 0) {
        kktpTableRows.push(
          new TableRow({
            children: [
              new TableCell({
                width: { size: 5, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(idx + 1), font: "Arial", size: 18 })] })],
              }),
              new TableCell({
                width: { size: 20, type: WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: `${item.tpKode || ""}: `, bold: true, font: "Arial", size: 18 }),
                      new TextRun({ text: item.tpTeks || "", font: "Arial", size: 18 })
                    ]
                  })
                ],
              }),
              new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, borders: commonBorder, children: [new Paragraph({ text: "" })] }),
              new TableCell({ width: { size: 12.5, type: WidthType.PERCENTAGE }, borders: commonBorder, children: [new Paragraph({ text: "" })] }),
              new TableCell({ width: { size: 12.5, type: WidthType.PERCENTAGE }, borders: commonBorder, children: [new Paragraph({ text: "" })] }),
              new TableCell({ width: { size: 12.5, type: WidthType.PERCENTAGE }, borders: commonBorder, children: [new Paragraph({ text: "" })] }),
              new TableCell({ width: { size: 12.5, type: WidthType.PERCENTAGE }, borders: commonBorder, children: [new Paragraph({ text: "" })] }),
            ],
          })
        );
      } else {
        criteriaList.forEach((crit: any, cIdx: number) => {
          const cells = [];

          if (cIdx === 0) {
            cells.push(
              new TableCell({
                width: { size: 5, type: WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(idx + 1), font: "Arial", size: 18 })] })],
              })
            );
            cells.push(
              new TableCell({
                width: { size: 20, type: WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: `${item.tpKode || ""}: `, bold: true, font: "Arial", size: 18 }),
                      new TextRun({ text: item.tpTeks || "", font: "Arial", size: 18 })
                    ]
                  })
                ],
              })
            );
          }

          cells.push(
            new TableCell({
              width: { size: 25, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: crit.aspek || "", bold: true, font: "Arial", size: 18 })] })],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 12.5, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: crit.baruBerkembang || "", font: "Arial", size: 16 })] })],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 12.5, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: crit.layak || "", font: "Arial", size: 16 })] })],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 12.5, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: crit.cakap || "", font: "Arial", size: 16 })] })],
            })
          );

          cells.push(
            new TableCell({
              width: { size: 12.5, type: WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new Paragraph({ children: [new TextRun({ text: crit.mahir || "", font: "Arial", size: 16 })] })],
            })
          );

          kktpTableRows.push(new TableRow({ children: cells }));
        });
      }
    });
  }

  children.push(new Table({ rows: kktpTableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
  children.push(new Paragraph({ spacing: { before: 240, after: 120 } }));

  // 7. Heading: TINDAK LANJUT
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      pageBreakBefore: true,
      children: [new TextRun({ text: "TINDAK LANJUT", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );

  // 8. Table 3: Tindak Lanjut
  const tindakLanjutTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Kategori Interval", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Status Kelulusan", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
          }),
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Tindak Lanjut / Intervensi", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "0 - 60 %\n(Mulai Berkembang)", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Belum Tuntas", bold: true, color: "FF0000", font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Peserta didik wajib mengikuti program bimbingan intensif/remedial individu atau kelompok kecil pada materi yang belum dikuasai (misalnya, simulasi routing/perhitungan subnetting ulang menggunakan alat peraga interaktif sebelum dilakukan asesmen ulang).`,
                    font: "Arial",
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "61 - 70 %\n(Layak)", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Tuntas (Syarat)", bold: true, color: "D97706", font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Peserta didik diberikan pendampingan berupa latihan terbimbing tambahan (peer tutoring) atau tugas mandiri terfokus untuk memperkuat pemahaman operasional dasar dari kompetensi yang belum sepenuhnya lancar.`,
                    font: "Arial",
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "71 - 80 %\n(Cakap)", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Tuntas", bold: true, color: "059669", font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Peserta didik langsung melanjutkan ke materi pembelajaran berikutnya dan diarahkan untuk mengeksplorasi studi kasus penyelesaian masalah jaringan riil dengan skala yang lebih luas.`,
                    font: "Arial",
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "81 - 100 %\n(Mahir)", bold: true, font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new Paragraph({ children: [new TextRun({ text: "Tuntas (Sangat Baik)", bold: true, color: "2563EB", font: "Arial", size: 18 })] })],
          }),
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Peserta didik diberikan program pengayaan berupa tantangan proyek desain infrastruktur jaringan industri skala enterprise, diajarkan teknologi advanced (seperti implementasi jaringan nirkabel cerdas/SDN), atau ditunjuk sebagai tutor sebaya (peer tutor) untuk membantu rekan sekelas yang berada di kategori mulai berkembang.`,
                    font: "Arial",
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(tindakLanjutTable);
  children.push(...createSignatures(meta));

  return new Document({
    sections: [{ properties: {}, children }],
  });
}

// 5. RENCANA PEMBELAJARAN MENDALAM (RPM)
export function generateRPMDoc(meta: Metadata, data: any): Document {
  // Helper function to build formatted table cells easily
  function makeCell(
    content: string | string[] | Paragraph[],
    widthPercent: number,
    shadingColor?: string,
    boldText = false,
    textSize = 18,
    isItalic = false,
    textColor = "111111",
    cellBorders = commonBorder
  ): TableCell {
    let cellChildren: Paragraph[] = [];
    
    if (Array.isArray(content)) {
      if (content.length === 0) {
        cellChildren.push(new Paragraph({
          children: [new TextRun({ text: "-", font: "Arial", size: textSize, color: textColor })]
        }));
      } else if (typeof content[0] === "string") {
        (content as string[]).forEach((str) => {
          cellChildren.push(new Paragraph({
            spacing: { before: 20, after: 20 },
            children: [new TextRun({ text: str, font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
          }));
        });
      } else {
        cellChildren = content as Paragraph[];
      }
    } else if (typeof content === "string") {
      cellChildren.push(new Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: content || "-", font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
      }));
    } else {
      cellChildren.push(content as Paragraph);
    }

    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      shading: shadingColor ? { fill: shadingColor } : undefined,
      borders: cellBorders,
      children: cellChildren,
    });
  }

  // Helper to format bullets nicely with custom symbols
  function formatBullets(items: any): string[] {
    if (Array.isArray(items)) {
      return items.map(item => `• ${item}`);
    }
    if (typeof items === "string") {
      return items.split("\n").filter(Boolean).map(x => x.startsWith("•") || x.startsWith("-") ? x : `• ${x}`);
    }
    return ["-"];
  }

  const identitas = data.identitas || {};
  const identifikasi = data.identifikasi || {};
  const sarana = identifikasi.saranaPrasarana || {};
  const desain = data.desainPembelajaran || {};
  const langkah = data.langkahPembelajaran || {};
  const asesmenObj = data.asesmen || {};
  const remedialObj = data.remedialPengayaan || {};
  const refleksiObj = data.refleksi || {};

  // Title box with deep emerald green background and white text
  const titleBlockTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "auto" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
      left: { style: BorderStyle.NONE, size: 0, color: "auto" },
      right: { style: BorderStyle.NONE, size: 0, color: "auto" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "007C58" },
            width: { size: 100, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 180, after: 60 },
                children: [
                  new TextRun({
                    text: "MODUL AJAR PEMBELAJARAN MENDALAM",
                    bold: true,
                    size: 26, // 13pt
                    font: "Arial",
                    color: "FFFFFF",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 60, after: 180 },
                children: [
                  new TextRun({
                    text: `Mata Pelajaran: ${identitas.mataPelajaran || meta.mataPelajaran || "-"}  |  Fase: ${meta.fase.toUpperCase()}  |  Kelas/Semester: ${identitas.kelasSemester || "-"}`,
                    size: 18, // 9pt
                    font: "Arial",
                    color: "FFFFFF",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  const children: any[] = [
    titleBlockTable,
    new Paragraph({ spacing: { before: 120, after: 120 } }), // spacer

    // 1. INFORMASI ADMINISTRATIF Table
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "I. INFORMASI ADMINISTRATIF", bold: true, size: 22, font: "Arial", color: "111111" })]
    }),
  ];

  const infoRows = [
    new TableRow({
      children: [
        makeCell("Nama Penyusun / Guru", 30, "F1F5F9", true, 18),
        makeCell(identitas.namaGuru || meta.namaGuru || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Institusi / Sekolah", 30, "F1F5F9", true, 18),
        makeCell(identitas.sekolah || meta.jenjangSekolah || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Mata Pelajaran", 30, "F1F5F9", true, 18),
        makeCell(identitas.mataPelajaran || meta.mataPelajaran || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Fase / Kelas", 30, "F1F5F9", true, 18),
        makeCell(`Fase ${meta.fase.toUpperCase()} / ${identitas.kelasSemester || "-"}`, 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Semester & Tahun Ajaran", 30, "F1F5F9", true, 18),
        makeCell(identitas.semesterTahun || `${identitas.kelasSemester?.includes("Semester") ? "" : "Semester Ganjil"} - ${meta.tahunAjaran}`, 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Topik / Materi Pokok", 30, "F1F5F9", true, 18),
        makeCell(identitas.topik || data.judul?.replace("Rencana Pembelajaran Mendalam (RPM) - Topik ", "") || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Model Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(identitas.modelPembelajaran || "Tatap Muka / Reguler", 70, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: infoRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  // 2. A. IDENTIFIKASI Table
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "II. DESAIN PENGEMBANGAN DAN IMPLEMENTASI", bold: true, size: 22, font: "Arial", color: "111111" })]
    }),
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "A. IDENTIFIKASI", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );

  const identifikasiRows = [
    new TableRow({
      children: [
        makeCell("Aspek Identifikasi", 30, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Deskripsi Analisis Pengembangan", 70, "007C58", true, 18, false, "FFFFFF"),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Target Peserta Didik", 30, "F1F5F9", true, 18),
        makeCell(identifikasi.targetPesertaDidik || "Peserta didik reguler/tipikal kelas terkait.", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Kompetensi Prasyarat", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(identifikasi.kompetensiPrasyarat), 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Profil Lulusan (8 Dimensi)", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(identifikasi.profilLulusan), 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Sarana dan Prasarana", 30, "F1F5F9", true, 18),
        makeCell([
          `• Fasilitas Fisik: ${sarana.fasilitasFisik || "-"}`,
          `• Piranti Lunak: ${sarana.pirantiLunak || "-"}`,
          `• Sumber Belajar: ${sarana.sumberBelajar || "-"}`
        ], 70, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: identifikasiRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  // 3. B. DESAIN PEMBELAJARAN Table
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "B. DESAIN PEMBELAJARAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );

  const desainRows = [
    new TableRow({
      children: [
        makeCell("Aspek Desain", 30, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Rincian Implementasi Kurikulum Merdeka", 70, "007C58", true, 18, false, "FFFFFF"),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Capaian Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(desain.capaianPembelajaran || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Lintas Disiplin Ilmu", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(desain.lintasDisiplin), 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Tujuan Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(desain.tujuanPembelajaran), 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Topik Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(desain.topikPembelajaran || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Praktik Pedagogis", 30, "F1F5F9", true, 18),
        makeCell(desain.praktikPedagogis || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Kemitraan Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(desain.kemitraanPembelajaran || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Lingkungan Belajar", 30, "F1F5F9", true, 18),
        makeCell(desain.lingkunganBelajar || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Pemanfaatan Digital", 30, "F1F5F9", true, 18),
        makeCell(desain.pemanfaatanDigital || "-", 70, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: desainRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  // 4. C. LANGKAH PEMBELAJARAN Table
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "C. LANGKAH PEMBELAJARAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );

  const formatLangkahBullets = (guruList: string[], siswaList: string[]): Paragraph[] => {
    const paras: Paragraph[] = [];
    paras.push(new Paragraph({
      spacing: { before: 40, after: 20 },
      children: [new TextRun({ text: "Aktivitas Guru:", bold: true, font: "Arial", size: 18 })]
    }));
    guruList.forEach(item => {
      paras.push(new Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: `• ${item}`, font: "Arial", size: 18 })]
      }));
    });

    paras.push(new Paragraph({
      spacing: { before: 80, after: 20 },
      children: [new TextRun({ text: "Aktivitas Peserta Didik:", bold: true, font: "Arial", size: 18 })]
    }));
    siswaList.forEach(item => {
      paras.push(new Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: `• ${item}`, font: "Arial", size: 18 })]
      }));
    });

    return paras;
  };

  const pFase = langkah.pendahuluan || { durasi: "15 Menit", guru: [], siswa: [], pendekatan: "MINDFUL" };
  const ieFase = langkah.intiEksplorasi || { durasi: "45 Menit", guru: [], siswa: [], pendekatan: "MEANINGFUL" };
  const irFase = langkah.intiRancangan || { durasi: "45 Menit", guru: [], siswa: [], pendekatan: "JOYFUL" };
  const penFase = langkah.penutup || { durasi: "15 Menit", guru: [], siswa: [], pendekatan: "MINDFUL & MEANINGFUL" };

  const langkahRows = [
    new TableRow({
      children: [
        makeCell("Fase Kegiatan", 20, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Aktivitas Pendidik & Peserta Didik", 65, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Pendekatan Deep Learning", 15, "007C58", true, 18, false, "FFFFFF"),
      ],
    }),
    new TableRow({
      children: [
        makeCell(`Pendahuluan\n(${pFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(pFase.guru || [], pFase.siswa || []), 65, undefined, false, 18),
        makeCell(pFase.pendekatan || "MINDFUL", 15, undefined, true, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell(`Eksplorasi Konsep & Analisis\n(${ieFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(ieFase.guru || [], ieFase.siswa || []), 65, undefined, false, 18),
        makeCell(ieFase.pendekatan || "MEANINGFUL", 15, undefined, true, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell(`Perancangan & Simulasi\n(${irFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(irFase.guru || [], irFase.siswa || []), 65, undefined, false, 18),
        makeCell(irFase.pendekatan || "JOYFUL", 15, undefined, true, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell(`Penutup & Evaluasi\n(${penFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(penFase.guru || [], penFase.siswa || []), 65, undefined, false, 18),
        makeCell(penFase.pendekatan || "MINDFUL & MEANINGFUL", 15, undefined, true, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: langkahRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  // 5. D. ASESMEN & PENILAIAN Table
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "D. ASESMEN & PENILAIAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );

  const diagAs = asesmenObj.diagnostik || { metode: "-", instrumen: "-", kriteria: "-" };
  const formAs = asesmenObj.formatif || { metode: "-", instrumen: "-", kriteria: "-" };
  const sumAs = asesmenObj.sumatif || { metode: "-", instrumen: "-", kriteria: "-" };

  const asesmenRows = [
    new TableRow({
      children: [
        makeCell("Kategori Asesmen", 20, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Metode Penilaian", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Instrumen Penilaian", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Kriteria Keberhasilan", 30, "007C58", true, 18, false, "FFFFFF"),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Asesmen Diagnostik (Awal)", 20, "F1F5F9", true, 18),
        makeCell(diagAs.metode || "-", 25, undefined, false, 18),
        makeCell(diagAs.instrumen || "-", 25, undefined, false, 18),
        makeCell(diagAs.kriteria || "-", 30, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Asesmen Formatif (Proses)", 20, "F1F5F9", true, 18),
        makeCell(formAs.metode || "-", 25, undefined, false, 18),
        makeCell(formAs.instrumen || "-", 25, undefined, false, 18),
        makeCell(formAs.kriteria || "-", 30, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Asesmen Sumatif (Akhir)", 20, "F1F5F9", true, 18),
        makeCell(sumAs.metode || "-", 25, undefined, false, 18),
        makeCell(sumAs.instrumen || "-", 25, undefined, false, 18),
        makeCell(sumAs.kriteria || "-", 30, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: asesmenRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  // 6. E. REMEDIAL & PENGAYAAN Table
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "E. REMEDIAL & PENGAYAAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );

  const remObj = remedialObj.remedial || { sasaran: "Siswa yang belum mencapai kriteria ketuntasan.", aktivitas: [] };
  const pengObj = remedialObj.pengayaan || { sasaran: "Siswa yang telah melampaui kriteria ketuntasan.", aktivitas: [] };

  const remedialRows = [
    new TableRow({
      children: [
        makeCell("Program Kegiatan", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Sasaran Peserta Didik", 35, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Bentuk Aktivitas Intervensi", 40, "007C58", true, 18, false, "FFFFFF"),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Program Remedial", 25, "F1F5F9", true, 18),
        makeCell(remObj.sasaran || "-", 35, undefined, false, 18),
        makeCell(formatBullets(remObj.aktivitas), 40, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Program Pengayaan", 25, "F1F5F9", true, 18),
        makeCell(pengObj.sasaran || "-", 35, undefined, false, 18),
        makeCell(formatBullets(pengObj.aktivitas), 40, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: remedialRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  // 7. F. REFLEKSI Table
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "F. REFLEKSI", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );

  const sRef = refleksiObj.siswa || [];
  const gRef = refleksiObj.guru || [];

  const refleksiRows = [
    new TableRow({
      children: [
        makeCell("Subjek Refleksi", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Pertanyaan Reflektif Pemicu Kesadaran Belajar", 75, "007C58", true, 18, false, "FFFFFF"),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Refleksi Peserta Didik", 25, "F1F5F9", true, 18),
        makeCell(formatBullets(sRef), 75, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Refleksi Pendidik (Guru)", 25, "F1F5F9", true, 18),
        makeCell(formatBullets(gRef), 75, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: refleksiRows,
    }),
    new Paragraph({ spacing: { before: 240, after: 120 } })
  );

  children.push(...createSignatures(meta));

  return new Document({
    sections: [{ properties: {}, children }],
  });
}

export function generateTopikDoc(meta: Metadata, data: any): Document {
  const children: any[] = [...createDocHeader(data.judul || "DAFTAR TOPIK PEMBELAJARAN", meta)];

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 240, after: 120 },
      children: [new TextRun({ text: "DAFTAR TOPIK / MATERI UTAMA", bold: true, size: 24, font: "Arial", color: "111111" })]
    })
  );

  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ width: { size: 5, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "No", bold: true, font: "Arial" })] })] }),
        new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Topik Utama", bold: true, font: "Arial" })] })] }),
        new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Deskripsi Cakupan", bold: true, font: "Arial" })] })] }),
        new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Tujuan Pembelajaran (TP)", bold: true, font: "Arial" })] })] }),
        new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Durasi & Sem.", bold: true, font: "Arial" })] })] }),
      ],
    }),
  ];

  if (data && Array.isArray(data.daftarTopik)) {
    data.daftarTopik.forEach((t: any) => {
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(t.no), font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t.topik || "", bold: true, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t.deskripsi || "", font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `[${t.tpKode}] ${t.tpTeks}`, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${t.alokasiWaktu} (Sem. ${t.semester})`, font: "Arial" })] })] }),
          ],
        })
      );
    });
  }

  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, borders: commonBorder }));
  children.push(...createSignatures(meta));

  return new Document({ sections: [{ properties: {}, children }] });
}

export function generateLKPDDoc(meta: Metadata, data: any): Document {
  const children: any[] = [];

  // Helper for cell padding and formatting
  function makeCell(
    content: string | string[] | Paragraph[],
    widthPercent: number,
    shadingColor?: string,
    boldText = false,
    textSize = 18,
    isItalic = false,
    textColor = "111111",
    align = AlignmentType.LEFT,
    cellBorders = commonBorder
  ): TableCell {
    let cellChildren: Paragraph[] = [];
    
    if (Array.isArray(content)) {
      if (content.length === 0) {
        cellChildren.push(new Paragraph({
          alignment: align,
          children: [new TextRun({ text: "-", font: "Arial", size: textSize, color: textColor })]
        }));
      } else if (typeof content[0] === "string") {
        (content as string[]).forEach((str) => {
          cellChildren.push(new Paragraph({
            alignment: align,
            spacing: { before: 40, after: 40 },
            children: [new TextRun({ text: str, font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
          }));
        });
      } else {
        cellChildren = content as Paragraph[];
      }
    } else if (typeof content === "string") {
      cellChildren.push(new Paragraph({
        alignment: align,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: content || "-", font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
      }));
    } else {
      cellChildren.push(content as Paragraph);
    }

    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      shading: shadingColor ? { fill: shadingColor } : undefined,
      borders: cellBorders,
      margins: { top: 100, bottom: 100, left: 150, right: 150 },
      children: cellChildren,
    });
  }

  // Accent Header generator helper
  function createAccentHeader(text: string): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.SINGLE, size: 24, color: "007A54" }, // Green bar
                right: { style: BorderStyle.NONE },
              },
              margins: { left: 120, top: 120, bottom: 120 },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: text.toUpperCase(), bold: true, size: 22, font: "Arial", color: "007A54" })
                  ]
                })
              ]
            })
          ]
        })
      ]
    });
  }

  // Green Title Banner Table
  const titleBannerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "007C58" },
            margins: { top: 240, bottom: 240, left: 150, right: 150 },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: data.judul || "LEMBAR KERJA PESERTA DIDIK (LKPD)",
                    bold: true,
                    size: 26,
                    font: "Arial",
                    color: "FFFFFF",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran || "-"} | Fase/Kelas: ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`,
                    bold: true,
                    size: 18,
                    font: "Arial",
                    color: "FFFFFF",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(titleBannerTable);
  children.push(new Paragraph({ spacing: { after: 120 } })); // Spacer

  // Metadata / Identitas Table
  const metadataRows = [
    new TableRow({
      children: [
        makeCell("Nama Guru / Penyusun", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.namaGuru || meta.namaGuru || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Instansi / Sekolah", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.sekolah || meta.jenjangSekolah || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Mata Pelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.mataPelajaran || meta.mataPelajaran || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Jenjang / Fase / Kelas", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(`SMK/MAK / ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`, 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Semester & Tahun Ajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.semesterTahun || `Ganjil - ${meta.tahunAjaran || ""}`, 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Topik / Materi Pokok", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.topikMateri || data.topik || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Model Pembelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.modelPembelajaran || "Tatap Muka / Reguler", 70, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: metadataRows,
    })
  );
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  // TUJUAN & PETUNJUK Section
  children.push(createAccentHeader("TUJUAN & PETUNJUK"));
  children.push(new Paragraph({ spacing: { after: 120 } })); // Spacer

  // Bullet items for Tujuan Pembelajaran
  const tpList = Array.isArray(data.tujuanPembelajaran) ? data.tujuanPembelajaran : [];
  const tpParagraphs = tpList.map(item => new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: "•  ", font: "Arial", size: 18, bold: true }),
      new TextRun({ text: item, font: "Arial", size: 18 })
    ]
  }));

  // Numbered items for Petunjuk Pengerjaan
  const petunjukList = Array.isArray(data.petunjukPengerjaan) ? data.petunjukPengerjaan : (Array.isArray(data.petunjukBelajar) ? data.petunjukBelajar : []);
  const petunjukParagraphs = petunjukList.map((item, idx) => new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: `${idx + 1}.  `, font: "Arial", size: 18, bold: true }),
      new TextRun({ text: item, font: "Arial", size: 18 })
    ]
  }));

  const tujuanPetunjukTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          makeCell("Aspek Kerja", 30, "007C58", true, 18, false, "FFFFFF", AlignmentType.LEFT),
          makeCell("Deskripsi Panduan", 70, "007C58", true, 18, false, "FFFFFF", AlignmentType.LEFT),
        ]
      }),
      new TableRow({
        children: [
          makeCell("Tujuan Pembelajaran", 30, "F0FDF4", true, 18, false, "111111"),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            borders: commonBorder,
            children: tpParagraphs.length > 0 ? tpParagraphs : [new Paragraph({ children: [new TextRun({ text: "-", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new TableRow({
        children: [
          makeCell("Petunjuk Pengerjaan", 30, "F0FDF4", true, 18, false, "111111"),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            borders: commonBorder,
            children: petunjukParagraphs.length > 0 ? petunjukParagraphs : [new Paragraph({ children: [new TextRun({ text: "-", font: "Arial", size: 18 })] })]
          })
        ]
      })
    ]
  });

  children.push(tujuanPetunjukTable);
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  // RINGKASAN MATERI Section
  children.push(createAccentHeader("RINGKASAN MATERI"));
  children.push(new Paragraph({ spacing: { after: 120 } })); // Spacer

  const ringkasanRows = [
    new TableRow({
      children: [
        makeCell("Komponen Utama", 30, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Ringkasan Konseptual " + (data.topik || "Materi"), 70, "007C58", true, 18, false, "FFFFFF"),
      ]
    })
  ];

  if (Array.isArray(data.ringkasanMateri)) {
    data.ringkasanMateri.forEach(item => {
      let ringkasanParagraphs: Paragraph[] = [];
      const lines = typeof item.ringkasan === "string" ? item.ringkasan.split("\n") : [];
      if (lines.length > 1) {
        lines.forEach(line => {
          const cleanLine = line.trim();
          if (cleanLine.startsWith("•") || cleanLine.startsWith("-")) {
            const textPart = cleanLine.replace(/^[•\-]\s*/, "");
            const boldMatch = textPart.match(/^([^:]+):(.*)$/);
            if (boldMatch) {
              ringkasanParagraphs.push(new Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new TextRun({ text: "•  ", font: "Arial", size: 18, bold: true }),
                  new TextRun({ text: boldMatch[1] + ": ", font: "Arial", size: 18, bold: true }),
                  new TextRun({ text: boldMatch[2], font: "Arial", size: 18 })
                ]
              }));
            } else {
              ringkasanParagraphs.push(new Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new TextRun({ text: "•  ", font: "Arial", size: 18, bold: true }),
                  new TextRun({ text: textPart, font: "Arial", size: 18 })
                ]
              }));
            }
          } else {
            ringkasanParagraphs.push(new Paragraph({
              spacing: { before: 40, after: 40 },
              children: [new TextRun({ text: cleanLine, font: "Arial", size: 18 })]
            }));
          }
        });
      } else {
        ringkasanParagraphs.push(new Paragraph({
          spacing: { before: 40, after: 40 },
          children: [new TextRun({ text: item.ringkasan || "-", font: "Arial", size: 18 })]
        }));
      }

      ringkasanRows.push(
        new TableRow({
          children: [
            makeCell(item.komponen || "-", 30, "F0FDF4", true, 18, false, "111111"),
            new TableCell({
              width: { size: 70, type: WidthType.PERCENTAGE },
              margins: { top: 100, bottom: 100, left: 150, right: 150 },
              borders: commonBorder,
              children: ringkasanParagraphs
            })
          ]
        })
      );
    });
  }

  const ringkasanTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: ringkasanRows,
  });

  children.push(ringkasanTable);
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  // AKTIVITAS / SOAL Section
  children.push(createAccentHeader("AKTIVITAS / SOAL"));
  children.push(new Paragraph({ spacing: { after: 120 } })); // Spacer

  if (Array.isArray(data.aktivitasList)) {
    data.aktivitasList.forEach((item, idx) => {
      const activityTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              makeCell(item.judul || `Aktivitas ${idx + 1}`, 100, "007C58", true, 20, false, "FFFFFF"),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                width: { size: 100, type: WidthType.PERCENTAGE },
                shading: { fill: "F0FDF4" },
                margins: { top: 150, bottom: 150, left: 150, right: 150 },
                borders: commonBorder,
                children: (item.studiKasus || "").split("\n").map((line: string) => {
                  const cleanLine = line.trim();
                  const isStudiKasusHeader = cleanLine.toLowerCase().startsWith("studi kasus:");
                  return new Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [
                      new TextRun({
                        text: cleanLine,
                        font: "Arial",
                        size: 18,
                        bold: isStudiKasusHeader,
                      })
                    ]
                  });
                })
              })
            ]
          })
        ]
      });

      children.push(activityTable);
      children.push(new Paragraph({ spacing: { after: 180 } }));
    });
  } else if (data.tugasAktivitas && Array.isArray(data.tugasAktivitas.pertanyaan)) {
    // Fallback for legacy data
    const legacyActivityTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            makeCell(data.judul || "Aktivitas Pembelajaran", 100, "007C58", true, 20, false, "FFFFFF"),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              width: { size: 100, type: WidthType.PERCENTAGE },
              shading: { fill: "F0FDF4" },
              margins: { top: 150, bottom: 150, left: 150, right: 150 },
              borders: commonBorder,
              children: [
                new Paragraph({
                  spacing: { before: 40, after: 80 },
                  children: [new TextRun({ text: "Studi Kasus / Instruksi Kerja:", bold: true, font: "Arial", size: 18 })]
                }),
                new Paragraph({
                  spacing: { after: 120 },
                  children: [new TextRun({ text: data.tugasAktivitas.instruksi || "", font: "Arial", size: 18 })]
                }),
                ...data.tugasAktivitas.pertanyaan.map((q: string, idx: number) => {
                  return new Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [
                      new TextRun({ text: `Pertanyaan ${idx + 1}: `, bold: true, font: "Arial", size: 18 }),
                      new TextRun({ text: q, font: "Arial", size: 18 })
                    ]
                  });
                })
              ]
            })
          ]
        })
      ]
    });
    children.push(legacyActivityTable);
  }

  children.push(...createSignatures(meta));

  return new Document({ sections: [{ properties: {}, children }] });
}

export function generateAsesmenDoc(meta: Metadata, data: any): Document {
  // Helper function to build formatted table cells easily
  function makeCell(
    content: string | string[] | Paragraph[],
    widthPercent: number,
    shadingColor?: string,
    boldText = false,
    textSize = 18,
    isItalic = false,
    textColor = "111111",
    align: any = AlignmentType.LEFT,
    cellBorders = commonBorder
  ): TableCell {
    let cellChildren: Paragraph[] = [];
    
    if (Array.isArray(content)) {
      if (content.length === 0) {
        cellChildren.push(new Paragraph({
          alignment: align,
          children: [new TextRun({ text: "-", font: "Arial", size: textSize, color: textColor })]
        }));
      } else {
        if (content[0] instanceof Paragraph) {
          cellChildren = content as Paragraph[];
        } else {
          (content as string[]).forEach(line => {
            cellChildren.push(new Paragraph({
              alignment: align,
              children: [new TextRun({ text: line, font: "Arial", size: textSize, color: textColor })]
            }));
          });
        }
      }
    } else {
      cellChildren.push(new Paragraph({
        alignment: align,
        children: [new TextRun({ 
          text: content as string, 
          bold: boldText, 
          italics: isItalic,
          size: textSize, 
          font: "Arial", 
          color: textColor 
        })]
      }));
    }

    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      shading: shadingColor ? { fill: shadingColor } : undefined,
      margins: { top: 120, bottom: 120, left: 150, right: 150 },
      borders: cellBorders,
      children: cellChildren,
    });
  }

  const children: any[] = [];

  // 1. Beautiful Title Banner in emerald green
  const titleBannerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "007C58" },
            margins: { top: 240, bottom: 240, left: 150, right: 150 },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: "KISI-KISI DAN SOAL ASESMEN",
                    bold: true,
                    size: 26,
                    font: "Arial",
                    color: "FFFFFF",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran || "-"} | Fase/Kelas: ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`,
                    bold: true,
                    size: 18,
                    font: "Arial",
                    color: "FFFFFF",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(titleBannerTable);
  children.push(new Paragraph({ spacing: { after: 120 } })); // Spacer

  // 2. Metadata / Identitas Table
  const identitasRows = [
    new TableRow({
      children: [
        makeCell("Nama Guru / Penyusun", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.namaGuru || meta.namaGuru || "Rudi Akbar Saragih, S.Kom.", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Instansi / Sekolah", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.sekolah || meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Mata Pelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.mataPelajaran || meta.mataPelajaran || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Jenjang / Fase / Kelas", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(`SMK/MAK / ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`, 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Semester & Tahun Ajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.semesterTahun || `Ganjil - ${meta.tahunAjaran || "2026/2027"}`, 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Topik / Materi Pokok", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.topikMateri || data.topik || "-", 70, undefined, false, 18),
      ],
    }),
    new TableRow({
      children: [
        makeCell("Model Pembelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.modelPembelajaran || "Tatap Muka / Reguler", 70, undefined, false, 18),
      ],
    }),
  ];

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: identitasRows,
    })
  );
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  // 3. KISI-KISI SOAL Section
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new TextRun({ text: "KISI-KISI SOAL", bold: true, size: 22, font: "Arial", color: "007C58" })
      ]
    })
  );

  const kisiKisiRows = [
    new TableRow({
      children: [
        makeCell("No", 5, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Tujuan Pembelajaran (TP)", 25, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Indikator Soal", 45, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Level Kognitif", 12, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Bentuk Soal", 8, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("No Soal", 5, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
      ]
    })
  ];

  const listKisi = data.kisiKisi || [];
  listKisi.forEach((k: any) => {
    kisiKisiRows.push(
      new TableRow({
        children: [
          makeCell(String(k.no), 5, undefined, false, 18, false, "111111", AlignmentType.CENTER),
          makeCell(k.tujuanPembelajaran || "-", 25, undefined, false, 18),
          makeCell(k.indikatorSoal || "-", 45, undefined, false, 18),
          makeCell(k.levelKognitif || "-", 12, undefined, false, 18, false, "111111", AlignmentType.CENTER),
          makeCell(k.bentukSoal || "-", 8, undefined, false, 18, false, "111111", AlignmentType.CENTER),
          makeCell(String(k.noSoal), 5, undefined, false, 18, false, "111111", AlignmentType.CENTER),
        ]
      })
    );
  });

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: kisiKisiRows,
    })
  );
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  // 4. BUTIR SOAL ASESMEN Section
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new TextRun({ text: "BUTIR SOAL ASESMEN", bold: true, size: 22, font: "Arial", color: "007C58" })
      ]
    })
  );

  const butirSoalRows = [
    new TableRow({
      children: [
        makeCell("No", 5, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Tipe & Level", 15, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Butir Soal & Pilihan Jawaban / Instruksi Soal", 80, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
      ]
    })
  ];

  const listSoal = data.daftarSoal || data.soal || [];
  listSoal.forEach((s: any) => {
    const contentParagraphs: Paragraph[] = [];
    contentParagraphs.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 80 },
      children: [new TextRun({ text: s.soal || s.pertanyaan || "-", font: "Arial", size: 18 })]
    }));

    if (s.pilihan) {
      Object.keys(s.pilihan).forEach((opt: string) => {
        contentParagraphs.push(new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { before: 20, after: 20 },
          indent: { left: 240 },
          children: [
            new TextRun({ text: `${opt}. `, bold: true, font: "Arial", size: 18 }),
            new TextRun({ text: s.pilihan[opt] || "", font: "Arial", size: 18 }),
          ]
        }));
      });
    } else if (s.tipe === "Pilihan Ganda" || (!s.tipe && (s.a || s.b || s.c || s.d || s.e))) {
      const opts = ["A", "B", "C", "D", "E"];
      opts.forEach((opt) => {
        const val = s[opt.toLowerCase()];
        if (val) {
          contentParagraphs.push(new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 20, after: 20 },
            indent: { left: 240 },
            children: [
              new TextRun({ text: `${opt}. `, bold: true, font: "Arial", size: 18 }),
              new TextRun({ text: val, font: "Arial", size: 18 }),
            ]
          }));
        }
      });
    }

    const tipeLevel = `${s.tipe || "Pilihan Ganda"}\n${s.levelKognitif || ""}`;

    butirSoalRows.push(
      new TableRow({
        children: [
          makeCell(String(s.no), 5, undefined, false, 18, false, "111111", AlignmentType.CENTER),
          makeCell(tipeLevel, 15, undefined, true, 18, false, "111111", AlignmentType.CENTER),
          new TableCell({
            width: { size: 80, type: WidthType.PERCENTAGE },
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            borders: commonBorder,
            children: contentParagraphs,
          })
        ]
      })
    );
  });

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: butirSoalRows,
    })
  );
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  // 5. KUNCI JAWABAN & RUBRIK PENILAIAN Section
  children.push(
    new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new TextRun({ text: "KUNCI JAWABAN & RUBRIK PENILAIAN", bold: true, size: 22, font: "Arial", color: "007C58" })
      ]
    })
  );

  const kunciJawabanRows = [
    new TableRow({
      children: [
        makeCell("No Soal", 10, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Bentuk Soal", 15, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Kunci Jawaban / Kriteria Penilaian Rubrik", 65, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
        makeCell("Skor Maksimal", 10, "007C58", true, 18, false, "FFFFFF", AlignmentType.CENTER),
      ]
    })
  ];

  const pgSoal = listSoal.filter((s: any) => s.tipe === "Pilihan Ganda" || (!s.tipe && (s.pilihan || s.a)));
  const uraianSoal = listSoal.filter((s: any) => s.tipe === "Uraian" || (!s.tipe && !s.pilihan && !s.a));

  if (pgSoal.length > 0) {
    kunciJawabanRows.push(
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 4,
            shading: { fill: "E6F2ED" },
            margins: { top: 80, bottom: 80, left: 150, right: 150 },
            borders: commonBorder,
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [new TextRun({ text: "BAGIAN I: PILIHAN GANDA", bold: true, size: 18, font: "Arial", color: "007C58" })]
              })
            ]
          })
        ]
      })
    );

    pgSoal.forEach((s: any) => {
      kunciJawabanRows.push(
        new TableRow({
          children: [
            makeCell(String(s.no), 10, undefined, false, 18, false, "111111", AlignmentType.CENTER),
            makeCell(s.tipe || "Pilihan Ganda", 15, undefined, false, 18, false, "111111", AlignmentType.CENTER),
            makeCell(s.kunciJawaban || "-", 65, undefined, false, 18),
            makeCell(String(s.skorMaksimal || 5), 10, undefined, false, 18, false, "111111", AlignmentType.CENTER),
          ]
        })
      );
    });
  }

  if (uraianSoal.length > 0) {
    kunciJawabanRows.push(
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 4,
            shading: { fill: "E6F2ED" },
            margins: { top: 80, bottom: 80, left: 150, right: 150 },
            borders: commonBorder,
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [new TextRun({ text: "BAGIAN II: URAIAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
              })
            ]
          })
        ]
      })
    );

    uraianSoal.forEach((s: any) => {
      const lines = (s.kunciJawaban || s.pembahasan || "-").split("\n");
      const answerParagraphs = lines.map((l: string) => new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: l.trim(), font: "Arial", size: 18 })]
      }));

      kunciJawabanRows.push(
        new TableRow({
          children: [
            makeCell(String(s.no), 10, undefined, false, 18, false, "111111", AlignmentType.CENTER),
            makeCell(s.tipe || "Uraian", 15, undefined, false, 18, false, "111111", AlignmentType.CENTER),
            new TableCell({
              width: { size: 65, type: WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              borders: commonBorder,
              children: answerParagraphs,
            }),
            makeCell(String(s.skorMaksimal || 15), 10, undefined, false, 18, false, "111111", AlignmentType.CENTER),
          ]
        })
      );
    });
  }

  // Total Row
  const totalScore = listSoal.reduce((acc: number, s: any) => acc + (s.skorMaksimal || 5), 0);
  kunciJawabanRows.push(
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 3,
          shading: { fill: "E6F2ED" },
          margins: { top: 100, bottom: 100, left: 150, right: 150 },
          borders: commonBorder,
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: "TOTAL SKOR MAKSIMAL KESELURUHAN SOAL", bold: true, size: 18, font: "Arial", color: "007C58" })]
            })
          ]
        }),
        makeCell(String(totalScore || 100), 10, "E6F2ED", true, 18, false, "007C58", AlignmentType.CENTER),
      ]
    })
  );

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: kunciJawabanRows,
    })
  );
  children.push(new Paragraph({ spacing: { after: 240 } })); // Spacer

  children.push(...createSignatures(meta));

  return new Document({ sections: [{ properties: {}, children }] });
}
