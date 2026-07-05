var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// server.ts
var import_express = __toESM(require("express"), 1);
var import_path = __toESM(require("path"), 1);
var import_dotenv = __toESM(require("dotenv"), 1);
var import_genai = require("@google/genai");
var import_vite = require("vite");
var import_docx2 = require("docx");

// src/utils/docxGenerator.ts
var import_docx = require("docx");
var commonBorder = {
  top: { style: import_docx.BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
  bottom: { style: import_docx.BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
  left: { style: import_docx.BorderStyle.SINGLE, size: 8, color: "CCCCCC" },
  right: { style: import_docx.BorderStyle.SINGLE, size: 8, color: "CCCCCC" }
};
function createDocHeader(title, meta) {
  return [
    new import_docx.Paragraph({
      alignment: import_docx.AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [
        new import_docx.TextRun({
          text: title.toUpperCase(),
          bold: true,
          size: 28,
          // 14pt
          font: "Arial"
        })
      ]
    }),
    new import_docx.Paragraph({
      alignment: import_docx.AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [
        new import_docx.TextRun({
          text: `MATA PELAJARAN: ${meta.mataPelajaran.toUpperCase()} | FASE ${meta.fase.toUpperCase()}`,
          bold: true,
          size: 22,
          // 11pt
          font: "Arial",
          color: "555555"
        })
      ]
    }),
    // Metadata Table / Block
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [
        new import_docx.TextRun({ text: "IDENTITAS ADMINISTRATIF", bold: true, size: 18, font: "Arial", color: "444444" })
      ]
    }),
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [
        new import_docx.TextRun({ text: `Nama Guru : `, bold: true, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: meta.namaGuru || "-", size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: `   |   NIP : `, bold: true, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: meta.nipGuru || "-", size: 20, font: "Arial" })
      ]
    }),
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [
        new import_docx.TextRun({ text: `Sekolah / Jenjang : `, bold: true, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: `${meta.jenjangSekolah || "-"} (Fase ${meta.fase || "-"})`, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: `   |   Tahun Ajaran : `, bold: true, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: meta.tahunAjaran || "-", size: 20, font: "Arial" })
      ]
    }),
    new import_docx.Paragraph({
      spacing: { after: 240 },
      children: [
        new import_docx.TextRun({ text: `Kepala Sekolah : `, bold: true, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: meta.namaKepalaSekolah || "-", size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: `   |   NIP Kepsek : `, bold: true, size: 20, font: "Arial" }),
        new import_docx.TextRun({ text: meta.nipKepalaSekolah || "-", size: 20, font: "Arial" })
      ]
    }),
    new import_docx.Paragraph({
      spacing: { after: 240 },
      children: [
        new import_docx.TextRun({
          text: "_________________________________________________________________________________",
          color: "CCCCCC"
        })
      ]
    })
  ];
}
function createSignatures(meta) {
  return [
    new import_docx.Paragraph({ spacing: { before: 400, after: 120 } }),
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      borders: {
        top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        insideHorizontal: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        insideVertical: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
      },
      rows: [
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
              borders: {
                top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
              },
              children: [
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: "Mengetahui,", font: "Arial", size: 20 })]
                }),
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: "Kepala Sekolah", font: "Arial", size: 20 })],
                  spacing: { after: 1200 }
                  // Space for signature
                }),
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: `( ${meta.namaKepalaSekolah || "__________________________"} )`, bold: true, font: "Arial", size: 20 })]
                }),
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: `NIP. ${meta.nipKepalaSekolah || "..................................."}`, font: "Arial", size: 20 })]
                })
              ]
            }),
            new import_docx.TableCell({
              width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
              borders: {
                top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
              },
              children: [
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: "......................, 29 Juni 2026", font: "Arial", size: 20 })]
                }),
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: "Guru Mata Pelajaran", font: "Arial", size: 20 })],
                  spacing: { after: 1200 }
                  // Space for signature
                }),
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: meta.namaGuru || "__________________________", bold: true, font: "Arial", size: 20 })]
                }),
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.LEFT,
                  children: [new import_docx.TextRun({ text: `NIP. ${meta.nipGuru || "..................................."}`, font: "Arial", size: 20 })]
                })
              ]
            })
          ]
        })
      ]
    })
  ];
}
function generateCPDoc(meta, data) {
  const children = [
    ...createDocHeader("Capaian Pembelajaran (CP)", meta),
    // Regulasi rujukan
    new import_docx.Paragraph({
      spacing: { after: 120 },
      children: [
        new import_docx.TextRun({
          text: "Regulasi Rujukan: ",
          bold: true,
          size: 22,
          font: "Arial"
        }),
        new import_docx.TextRun({
          text: data.regulasi || "Keputusan Kepala BSKAP Nomor 046/H/KR/2025",
          italics: true,
          size: 22,
          font: "Arial"
        })
      ]
    }),
    // Rasional
    new import_docx.Paragraph({
      heading: import_docx.HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "1. Rasional Mata Pelajaran", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new import_docx.Paragraph({
      spacing: { after: 180 },
      children: [new import_docx.TextRun({ text: data.rasional || "", size: 22, font: "Arial" })]
    }),
    // Tujuan
    new import_docx.Paragraph({
      heading: import_docx.HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "2. Tujuan Mata Pelajaran", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new import_docx.Paragraph({
      spacing: { after: 180 },
      children: [new import_docx.TextRun({ text: data.tujuan || "", size: 22, font: "Arial" })]
    }),
    // Karakteristik
    new import_docx.Paragraph({
      heading: import_docx.HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "3. Karakteristik Mata Pelajaran", bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new import_docx.Paragraph({
      spacing: { after: 180 },
      children: [new import_docx.TextRun({ text: data.karakteristik || "", size: 22, font: "Arial" })]
    }),
    // Capaian Umum
    new import_docx.Paragraph({
      heading: import_docx.HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: `4. Capaian Umum Fase ${meta.fase.toUpperCase()}`, bold: true, size: 24, font: "Arial", color: "111111" })]
    }),
    new import_docx.Paragraph({
      spacing: { after: 180 },
      children: [new import_docx.TextRun({ text: data.capaianUmum || "", size: 22, font: "Arial" })]
    }),
    // Capaian per Elemen (Table)
    new import_docx.Paragraph({
      heading: import_docx.HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "5. Capaian Pembelajaran per Elemen", bold: true, size: 24, font: "Arial", color: "111111" })]
    })
  ];
  const tableRows = [
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "F2F2F2" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Elemen", bold: true, font: "Arial", size: 20 })] })]
        }),
        new import_docx.TableCell({
          width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "F2F2F2" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Capaian Pembelajaran Elemen", bold: true, font: "Arial", size: 20 })] })]
        })
      ]
    })
  ];
  if (Array.isArray(data.elemen)) {
    data.elemen.forEach((el) => {
      tableRows.push(
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              borders: commonBorder,
              children: [
                new import_docx.Paragraph({
                  children: [new import_docx.TextRun({ text: el.nama || "", bold: true, font: "Arial", size: 20 })],
                  spacing: { after: 60 }
                }),
                new import_docx.Paragraph({
                  children: [new import_docx.TextRun({ text: el.deskripsi || "", italics: true, font: "Arial", size: 18, color: "555555" })]
                })
              ]
            }),
            new import_docx.TableCell({
              borders: commonBorder,
              children: [
                new import_docx.Paragraph({
                  children: [new import_docx.TextRun({ text: el.capaian || "", font: "Arial", size: 20 })]
                })
              ]
            })
          ]
        })
      );
    });
  }
  children.push(new import_docx.Table({ rows: tableRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
  children.push(...createSignatures(meta));
  return new import_docx.Document({
    sections: [{ properties: {}, children }]
  });
}
function generateATPDoc(meta, data) {
  const children = [];
  const formatTpText = (tpText, idx) => {
    const trimmed = (tpText || "").trim();
    if (/^\d+[\.\)]/.test(trimmed)) {
      return trimmed;
    }
    return `${idx + 1}. ${trimmed}`;
  };
  const bannerTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "007A54" },
            // Deep green from the PDF image
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new import_docx.TextRun({
                    text: "CAPAIAN & ALUR TUJUAN PEMBELAJARAN",
                    bold: true,
                    color: "FFFFFF",
                    size: 28,
                    // 14pt
                    font: "Arial"
                  })
                ]
              }),
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                children: [
                  new import_docx.TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran} | Fase/Kelas: ${meta.fase} / XI`,
                    bold: true,
                    color: "FFFFFF",
                    size: 20,
                    // 10pt
                    font: "Arial"
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(bannerTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  const identityTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            // light mint green
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Nama Guru / Penyusun", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.namaGuru || "", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Instansi / Sekolah", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Mata Pelajaran", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.mataPelajaran || "", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Jenjang / Fase / Kelas", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: `SMK/MAK / ${meta.fase} / XI`, font: "Arial", size: 18 })] })]
          })
        ]
      })
    ]
  });
  children.push(identityTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "CAPAIAN PEMBELAJARAN", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );
  const cpTableRows = [
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Elemen", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Capaian Pembelajaran (CP)", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })]
        })
      ]
    })
  ];
  if (Array.isArray(data.atp)) {
    data.atp.forEach((item) => {
      cpTableRows.push(
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              shading: { fill: "F9FBF9" },
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [
                new import_docx.Paragraph({
                  children: [new import_docx.TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18, color: "111111" })]
                })
              ]
            }),
            new import_docx.TableCell({
              width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [
                new import_docx.Paragraph({
                  children: [new import_docx.TextRun({ text: item.capaian || item.deskripsi || "", font: "Arial", size: 18 })]
                })
              ]
            })
          ]
        })
      );
    });
  }
  children.push(new import_docx.Table({ rows: cpTableRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      pageBreakBefore: true,
      children: [new import_docx.TextRun({ text: "ALUR TUJUAN PEMBELAJARAN (ATP)", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );
  const atpTableRows = [
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Elemen", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Capaian Pembelajaran (CP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 32, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Tujuan Pembelajaran (TP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 13, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Kata Kunci Materi", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 10, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Profil Lulusan (SKL)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Perkiraan JP", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        })
      ]
    })
  ];
  if (Array.isArray(data.atp)) {
    data.atp.forEach((item) => {
      const tps = Array.isArray(item.tujuanPembelajaran) ? item.tujuanPembelajaran : [];
      const rowSpanVal = tps.length || 1;
      if (tps.length === 0) {
        atpTableRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18 })] })]
              }),
              new import_docx.TableCell({
                width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.capaian || "", font: "Arial", size: 18 })] })]
              }),
              new import_docx.TableCell({
                width: { size: 32, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "", font: "Arial", size: 18 })] })]
              }),
              new import_docx.TableCell({
                width: { size: 13, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "", font: "Arial", size: 18 })] })]
              }),
              new import_docx.TableCell({
                width: { size: 10, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "", font: "Arial", size: 16 })] })]
              }),
              new import_docx.TableCell({
                width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "", font: "Arial", size: 18 })] })]
              })
            ]
          })
        );
      } else {
        tps.forEach((tpObj, index) => {
          const cells = [];
          if (index === 0) {
            cells.push(
              new import_docx.TableCell({
                width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18 })] })]
              })
            );
            cells.push(
              new import_docx.TableCell({
                width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.capaian || "", font: "Arial", size: 18 })] })]
              })
            );
          }
          const formattedTp = formatTpText(tpObj.tp, index);
          cells.push(
            new import_docx.TableCell({
              width: { size: 32, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [
                new import_docx.Paragraph({
                  children: [
                    new import_docx.TextRun({ text: formattedTp, font: "Arial", size: 18 })
                  ]
                })
              ]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 13, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: tpObj.topik || "", font: "Arial", size: 18 })] })]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 10, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: tpObj.profilLulusan || tpObj.profilPancasila || "", font: "Arial", size: 16 })] })]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: tpObj.alokasiWaktu || "", font: "Arial", size: 18 })] })]
            })
          );
          atpTableRows.push(new import_docx.TableRow({ children: cells }));
        });
      }
    });
  }
  children.push(new import_docx.Table({ rows: atpTableRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
  children.push(...createSignatures(meta));
  return new import_docx.Document({
    sections: [{ properties: {}, children }]
  });
}
function generateProtaPromesDoc(meta, data) {
  const children = [];
  const bannerTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "007A54" },
            // Deep green
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new import_docx.TextRun({
                    text: "PROGRAM TAHUNAN & PROGRAM SEMESTER",
                    bold: true,
                    color: "FFFFFF",
                    size: 28,
                    // 14pt
                    font: "Arial"
                  })
                ]
              }),
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                children: [
                  new import_docx.TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran} | Fase/Kelas: ${meta.fase} / XI`,
                    bold: true,
                    color: "FFFFFF",
                    size: 20,
                    // 10pt
                    font: "Arial"
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(bannerTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  const identityTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Nama Guru / Penyusun", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.namaGuru || "-", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Instansi / Sekolah", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.jenjangSekolah || "-", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Mata Pelajaran", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.mataPelajaran || "-", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "F5F5F5" },
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Jenjang / Fase / Kelas", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            borders: commonBorder,
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: `${meta.jenjangSekolah || "SMK/MAK"} / ${meta.fase} / XI`, font: "Arial", size: 18 })] })]
          })
        ]
      })
    ]
  });
  children.push(identityTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  const protaHeading = new import_docx.Paragraph({
    spacing: { before: 240, after: 120 },
    children: [
      new import_docx.TextRun({
        text: "PROGRAM TAHUNAN (PROTA)",
        bold: true,
        color: "007A54",
        size: 24,
        // 12pt
        font: "Arial"
      })
    ]
  });
  children.push(protaHeading);
  const protaRows = [
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          width: { size: 8, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "No", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Elemen / Lingkup Materi", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          width: { size: 52, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tujuan Pembelajaran", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Alokasi Waktu (JP)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        })
      ]
    })
  ];
  let totalJp = 0;
  if (Array.isArray(data.prota)) {
    data.prota.forEach((item, idx) => {
      const numJp = parseInt(item.alokasiWaktu) || 0;
      totalJp += numJp;
      protaRows.push(
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: String(item.no || idx + 1), font: "Arial", size: 18 })] })]
            }),
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.elemen || "", bold: true, font: "Arial", size: 18 })] })]
            }),
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.tujuanPembelajaran || "", font: "Arial", size: 18 })] })]
            }),
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: item.alokasiWaktu || "", font: "Arial", size: 18 })] })]
            })
          ]
        })
      );
    });
  }
  protaRows.push(
    new import_docx.TableRow({
      children: [
        new import_docx.TableCell({
          columnSpan: 3,
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.RIGHT,
              children: [new import_docx.TextRun({ text: "TOTAL ALOKASI WAKTU SATU TAHUN", bold: true, font: "Arial", size: 18 })]
            })
          ]
        }),
        new import_docx.TableCell({
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.CENTER,
              children: [new import_docx.TextRun({ text: `${totalJp} JP`, bold: true, color: "007A54", font: "Arial", size: 18 })]
            })
          ]
        })
      ]
    })
  );
  const protaTable = new import_docx.Table({ rows: protaRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } });
  children.push(protaTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  const promesHeading = new import_docx.Paragraph({
    spacing: { before: 240, after: 120 },
    children: [
      new import_docx.TextRun({
        text: "PROGRAM SEMESTER (PROSEM)",
        bold: true,
        color: "007A54",
        size: 24,
        // 12pt
        font: "Arial"
      })
    ]
  });
  children.push(promesHeading);
  const promesRows = [
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          rowSpan: 2,
          width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "No", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          rowSpan: 2,
          width: { size: 40, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tujuan Pembelajaran", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          rowSpan: 2,
          width: { size: 8, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Alokasi (JP)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          columnSpan: 6,
          width: { size: 23.5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Semester 1 (Ganjil)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        }),
        new import_docx.TableCell({
          columnSpan: 6,
          width: { size: 23.5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Semester 2 (Genap)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
        })
      ]
    }),
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        ...["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => {
          return new import_docx.TableCell({
            width: { size: 3.91, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: m, bold: true, color: "FFFFFF", font: "Arial", size: 16 })] })]
          });
        })
      ]
    })
  ];
  if (Array.isArray(data.promes)) {
    data.promes.forEach((item, idx) => {
      const distCells = ["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"].map((m) => {
        const val = item.distribusi && item.distribusi[m] || "-";
        const isNum = !isNaN(parseInt(val));
        const cellShading = isNum ? { fill: "E6F0FA" } : val === "ASG" || val === "ASE" ? { fill: "F2F2F2" } : void 0;
        return new import_docx.TableCell({
          shading: cellShading,
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.CENTER,
              children: [
                new import_docx.TextRun({
                  text: val,
                  bold: isNum || val === "ASG" || val === "ASE",
                  font: "Arial",
                  size: 16,
                  color: isNum ? "1A365D" : val === "ASG" || val === "ASE" ? "4A5568" : "718096"
                })
              ]
            })
          ]
        });
      });
      promesRows.push(
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: String(item.no || idx + 1), font: "Arial", size: 18 })] })]
            }),
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: item.tujuanPembelajaran || "", font: "Arial", size: 18 })] })]
            }),
            new import_docx.TableCell({
              borders: commonBorder,
              children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: String(item.alokasiWaktu || ""), font: "Arial", size: 18 })] })]
            }),
            ...distCells
          ]
        })
      );
    });
  }
  const totalPromesJp = Array.isArray(data.promes) ? data.promes.reduce((sum, item) => sum + (Number(item.alokasiWaktu) || 0), 0) : 0;
  promesRows.push(
    new import_docx.TableRow({
      children: [
        new import_docx.TableCell({
          columnSpan: 2,
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.RIGHT,
              children: [new import_docx.TextRun({ text: "Jumlah JP per Semester", bold: true, font: "Arial", size: 18 })]
            })
          ]
        }),
        new import_docx.TableCell({
          shading: { fill: "F5F5F5" },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.CENTER,
              children: [new import_docx.TextRun({ text: `${totalPromesJp}`, bold: true, font: "Arial", size: 18 })]
            })
          ]
        }),
        new import_docx.TableCell({
          columnSpan: 6,
          shading: { fill: "EBF5FA" },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.CENTER,
              children: [new import_docx.TextRun({ text: `Semester Ganjil: ${data.jumlahJpSemester1 || 172} JP`, bold: true, color: "007A54", font: "Arial", size: 18 })]
            })
          ]
        }),
        new import_docx.TableCell({
          columnSpan: 6,
          shading: { fill: "EBF5FA" },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.CENTER,
              children: [new import_docx.TextRun({ text: `Semester Genap: ${data.jumlahJpSemester2 || 188} JP`, bold: true, color: "007A54", font: "Arial", size: 18 })]
            })
          ]
        })
      ]
    })
  );
  const promesTable = new import_docx.Table({ rows: promesRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } });
  children.push(promesTable);
  const promesFooterNote = new import_docx.Paragraph({
    spacing: { before: 120, after: 240 },
    children: [
      new import_docx.TextRun({
        text: "*Keterangan: ASG = Asesmen Sumatif Ganjil / Cadangan; ASE = Asesmen Sumatif Akhir Tahun / Kelulusan Fase. Angka menunjukkan distribusi alokasi waktu JP per bulan efektif.",
        italics: true,
        size: 16,
        font: "Arial",
        color: "718096"
      })
    ]
  });
  children.push(promesFooterNote);
  const sigTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
            children: [
              new import_docx.Paragraph({
                spacing: { before: 240 },
                children: [new import_docx.TextRun({ text: "Mengetahui,", font: "Arial", size: 20 })]
              }),
              new import_docx.Paragraph({
                children: [new import_docx.TextRun({ text: "Kepala Sekolah", bold: true, font: "Arial", size: 20 })]
              }),
              new import_docx.Paragraph({
                spacing: { before: 1200 },
                // Signature space
                children: [new import_docx.TextRun({ text: `( ${meta.namaKepalaSekolah || "__________________________"} )`, bold: true, font: "Arial", size: 20 })]
              }),
              new import_docx.Paragraph({
                children: [new import_docx.TextRun({ text: `NIP. ${meta.nipKepalaSekolah || "..................................."}`, font: "Arial", size: 20 })]
              })
            ]
          }),
          new import_docx.TableCell({
            width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
            children: [
              new import_docx.Paragraph({
                spacing: { before: 240 },
                children: [new import_docx.TextRun({ text: "......................, 29 Juni 2026", font: "Arial", size: 20 })]
              }),
              new import_docx.Paragraph({
                children: [new import_docx.TextRun({ text: "Guru Mata Pelajaran", bold: true, font: "Arial", size: 20 })]
              }),
              new import_docx.Paragraph({
                spacing: { before: 1200 },
                // Signature space
                children: [new import_docx.TextRun({ text: meta.namaGuru || "__________________________", bold: true, font: "Arial", size: 20 })]
              }),
              new import_docx.Paragraph({
                children: [new import_docx.TextRun({ text: `NIP. ${meta.nipGuru || "..................................."}`, font: "Arial", size: 20 })]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(sigTable);
  return new import_docx.Document({
    sections: [{ properties: {}, children }]
  });
}
function generateKKTPDoc(meta, data) {
  const children = [];
  const bannerTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: import_docx.BorderStyle.NONE, size: 0, color: "FFFFFF" }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "007A54" },
            // Deep green from the PDF image
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new import_docx.TextRun({
                    text: "ANALISIS KRITERIA KETERCAPAIAN (KKTP)",
                    bold: true,
                    color: "FFFFFF",
                    size: 28,
                    // 14pt
                    font: "Arial"
                  })
                ]
              }),
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                children: [
                  new import_docx.TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran} | Fase/Kelas: ${meta.fase} / XI`,
                    bold: true,
                    color: "FFFFFF",
                    size: 20,
                    // 10pt
                    font: "Arial"
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(bannerTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 120, after: 120 } }));
  const identityTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            // light mint green
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Nama Guru / Penyusun", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.namaGuru || "", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Instansi / Sekolah", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Mata Pelajaran", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: meta.mataPelajaran || "", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "E2F2ED" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Jenjang / Fase / Kelas", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: `SMK/MAK / ${meta.fase} / XI`, font: "Arial", size: 18 })] })]
          })
        ]
      })
    ]
  });
  children.push(identityTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "PENDEKATAN KKTP", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );
  const pendekatanTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: [
      new import_docx.TableRow({
        tableHeader: true,
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Metode Pendekatan", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Deskripsi dan Alasan Penggunaan", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Pendekatan Interval Nilai", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new import_docx.Paragraph({
                children: [
                  new import_docx.TextRun({
                    text: `Menggunakan interval kriteria untuk menentukan ketuntasan belajar peserta didik. Pendekatan ini dipilih karena memberikan gambaran objektif, terukur, dan mempermudah pemetaan tingkat kompetensi peserta didik (Mulai Berkembang, Layak, Cakap, Mahir) pada setiap indikator mata pelajaran ${meta.mataPelajaran} Fase ${meta.fase} (Kelas XI).`,
                    font: "Arial",
                    size: 18
                  })
                ]
              })
            ]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Ketentuan Ketuntasan", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new import_docx.Paragraph({
                children: [
                  new import_docx.TextRun({
                    text: "Peserta didik dinyatakan mencapai ketuntasan minimum jika mencapai kriteria minimal Layak (61-70) pada indikator-indikator kunci kompetensi esensial.",
                    font: "Arial",
                    size: 18
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(pendekatanTable);
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "ANALISIS KKTP", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );
  const kktpTableRows = [
    // Header Row 1
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
          rowSpan: 2,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "No", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
          rowSpan: 2,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Tujuan Pembelajaran (TP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
          rowSpan: 2,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Indikator Ketercapaian (IKTP)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
          columnSpan: 4,
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Kriteria Interval Nilai (%)", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
        })
      ]
    }),
    // Header Row 2
    new import_docx.TableRow({
      tableHeader: true,
      children: [
        new import_docx.TableCell({
          width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Mulai Berkembang (0-60)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Layak (61-70)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Cakap (71-80)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })]
        }),
        new import_docx.TableCell({
          width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
          shading: { fill: "007A54" },
          borders: commonBorder,
          margins: { top: 120, bottom: 120, left: 150, right: 150 },
          children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: "Mahir (81-100)", bold: true, font: "Arial", size: 16, color: "FFFFFF" })] })]
        })
      ]
    })
  ];
  if (Array.isArray(data.kktp)) {
    data.kktp.forEach((item, idx) => {
      const criteriaList = Array.isArray(item.kriteria) ? item.kriteria : [];
      const rowSpanVal = criteriaList.length || 1;
      if (criteriaList.length === 0) {
        kktpTableRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: String(idx + 1), font: "Arial", size: 18 })] })]
              }),
              new import_docx.TableCell({
                width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [
                  new import_docx.Paragraph({
                    children: [
                      new import_docx.TextRun({ text: `${item.tpKode || ""}: `, bold: true, font: "Arial", size: 18 }),
                      new import_docx.TextRun({ text: item.tpTeks || "", font: "Arial", size: 18 })
                    ]
                  })
                ]
              }),
              new import_docx.TableCell({ width: { size: 25, type: import_docx.WidthType.PERCENTAGE }, borders: commonBorder, children: [new import_docx.Paragraph({ text: "" })] }),
              new import_docx.TableCell({ width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE }, borders: commonBorder, children: [new import_docx.Paragraph({ text: "" })] }),
              new import_docx.TableCell({ width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE }, borders: commonBorder, children: [new import_docx.Paragraph({ text: "" })] }),
              new import_docx.TableCell({ width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE }, borders: commonBorder, children: [new import_docx.Paragraph({ text: "" })] }),
              new import_docx.TableCell({ width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE }, borders: commonBorder, children: [new import_docx.Paragraph({ text: "" })] })
            ]
          })
        );
      } else {
        criteriaList.forEach((crit, cIdx) => {
          const cells = [];
          if (cIdx === 0) {
            cells.push(
              new import_docx.TableCell({
                width: { size: 5, type: import_docx.WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [new import_docx.Paragraph({ alignment: import_docx.AlignmentType.CENTER, children: [new import_docx.TextRun({ text: String(idx + 1), font: "Arial", size: 18 })] })]
              })
            );
            cells.push(
              new import_docx.TableCell({
                width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
                rowSpan: rowSpanVal,
                borders: commonBorder,
                margins: { top: 120, bottom: 120, left: 150, right: 150 },
                children: [
                  new import_docx.Paragraph({
                    children: [
                      new import_docx.TextRun({ text: `${item.tpKode || ""}: `, bold: true, font: "Arial", size: 18 }),
                      new import_docx.TextRun({ text: item.tpTeks || "", font: "Arial", size: 18 })
                    ]
                  })
                ]
              })
            );
          }
          cells.push(
            new import_docx.TableCell({
              width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: crit.aspek || "", bold: true, font: "Arial", size: 18 })] })]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: crit.baruBerkembang || "", font: "Arial", size: 16 })] })]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: crit.layak || "", font: "Arial", size: 16 })] })]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: crit.cakap || "", font: "Arial", size: 16 })] })]
            })
          );
          cells.push(
            new import_docx.TableCell({
              width: { size: 12.5, type: import_docx.WidthType.PERCENTAGE },
              borders: commonBorder,
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: crit.mahir || "", font: "Arial", size: 16 })] })]
            })
          );
          kktpTableRows.push(new import_docx.TableRow({ children: cells }));
        });
      }
    });
  }
  children.push(new import_docx.Table({ rows: kktpTableRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
  children.push(new import_docx.Paragraph({ spacing: { before: 240, after: 120 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      pageBreakBefore: true,
      children: [new import_docx.TextRun({ text: "TINDAK LANJUT", bold: true, size: 24, font: "Arial", color: "007A54" })]
    })
  );
  const tindakLanjutTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: [
      new import_docx.TableRow({
        tableHeader: true,
        children: [
          new import_docx.TableCell({
            width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Kategori Interval", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
          }),
          new import_docx.TableCell({
            width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Status Kelulusan", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
          }),
          new import_docx.TableCell({
            width: { size: 55, type: import_docx.WidthType.PERCENTAGE },
            shading: { fill: "007A54" },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tindak Lanjut / Intervensi", bold: true, font: "Arial", size: 18, color: "FFFFFF" })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "0 - 60 %\n(Mulai Berkembang)", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Belum Tuntas", bold: true, color: "FF0000", font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 55, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new import_docx.Paragraph({
                children: [
                  new import_docx.TextRun({
                    text: `Peserta didik wajib mengikuti program bimbingan intensif/remedial individu atau kelompok kecil pada materi yang belum dikuasai (misalnya, simulasi routing/perhitungan subnetting ulang menggunakan alat peraga interaktif sebelum dilakukan asesmen ulang).`,
                    font: "Arial",
                    size: 18
                  })
                ]
              })
            ]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "61 - 70 %\n(Layak)", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tuntas (Syarat)", bold: true, color: "D97706", font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 55, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new import_docx.Paragraph({
                children: [
                  new import_docx.TextRun({
                    text: `Peserta didik diberikan pendampingan berupa latihan terbimbing tambahan (peer tutoring) atau tugas mandiri terfokus untuk memperkuat pemahaman operasional dasar dari kompetensi yang belum sepenuhnya lancar.`,
                    font: "Arial",
                    size: 18
                  })
                ]
              })
            ]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "71 - 80 %\n(Cakap)", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tuntas", bold: true, color: "059669", font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 55, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new import_docx.Paragraph({
                children: [
                  new import_docx.TextRun({
                    text: `Peserta didik langsung melanjutkan ke materi pembelajaran berikutnya dan diarahkan untuk mengeksplorasi studi kasus penyelesaian masalah jaringan riil dengan skala yang lebih luas.`,
                    font: "Arial",
                    size: 18
                  })
                ]
              })
            ]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "81 - 100 %\n(Mahir)", bold: true, font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tuntas (Sangat Baik)", bold: true, color: "2563EB", font: "Arial", size: 18 })] })]
          }),
          new import_docx.TableCell({
            width: { size: 55, type: import_docx.WidthType.PERCENTAGE },
            borders: commonBorder,
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            children: [
              new import_docx.Paragraph({
                children: [
                  new import_docx.TextRun({
                    text: `Peserta didik diberikan program pengayaan berupa tantangan proyek desain infrastruktur jaringan industri skala enterprise, diajarkan teknologi advanced (seperti implementasi jaringan nirkabel cerdas/SDN), atau ditunjuk sebagai tutor sebaya (peer tutor) untuk membantu rekan sekelas yang berada di kategori mulai berkembang.`,
                    font: "Arial",
                    size: 18
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(tindakLanjutTable);
  children.push(...createSignatures(meta));
  return new import_docx.Document({
    sections: [{ properties: {}, children }]
  });
}
function generateRPMDoc(meta, data) {
  function makeCell(content, widthPercent, shadingColor, boldText = false, textSize = 18, isItalic = false, textColor = "111111", cellBorders = commonBorder) {
    let cellChildren = [];
    if (Array.isArray(content)) {
      if (content.length === 0) {
        cellChildren.push(new import_docx.Paragraph({
          children: [new import_docx.TextRun({ text: "-", font: "Arial", size: textSize, color: textColor })]
        }));
      } else if (typeof content[0] === "string") {
        content.forEach((str) => {
          cellChildren.push(new import_docx.Paragraph({
            spacing: { before: 20, after: 20 },
            children: [new import_docx.TextRun({ text: str, font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
          }));
        });
      } else {
        cellChildren = content;
      }
    } else if (typeof content === "string") {
      cellChildren.push(new import_docx.Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new import_docx.TextRun({ text: content || "-", font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
      }));
    } else {
      cellChildren.push(content);
    }
    return new import_docx.TableCell({
      width: { size: widthPercent, type: import_docx.WidthType.PERCENTAGE },
      shading: shadingColor ? { fill: shadingColor } : void 0,
      borders: cellBorders,
      children: cellChildren
    });
  }
  function formatBullets(items) {
    if (Array.isArray(items)) {
      return items.map((item) => `\u2022 ${item}`);
    }
    if (typeof items === "string") {
      return items.split("\n").filter(Boolean).map((x) => x.startsWith("\u2022") || x.startsWith("-") ? x : `\u2022 ${x}`);
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
  const titleBlockTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE, size: 0, color: "auto" },
      bottom: { style: import_docx.BorderStyle.NONE, size: 0, color: "auto" },
      left: { style: import_docx.BorderStyle.NONE, size: 0, color: "auto" },
      right: { style: import_docx.BorderStyle.NONE, size: 0, color: "auto" }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "007C58" },
            width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { before: 180, after: 60 },
                children: [
                  new import_docx.TextRun({
                    text: "MODUL AJAR PEMBELAJARAN MENDALAM",
                    bold: true,
                    size: 26,
                    // 13pt
                    font: "Arial",
                    color: "FFFFFF"
                  })
                ]
              }),
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { before: 60, after: 180 },
                children: [
                  new import_docx.TextRun({
                    text: `Mata Pelajaran: ${identitas.mataPelajaran || meta.mataPelajaran || "-"}  |  Fase: ${meta.fase.toUpperCase()}  |  Kelas/Semester: ${identitas.kelasSemester || "-"}`,
                    size: 18,
                    // 9pt
                    font: "Arial",
                    color: "FFFFFF"
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  const children = [
    titleBlockTable,
    new import_docx.Paragraph({ spacing: { before: 120, after: 120 } }),
    // spacer
    // 1. INFORMASI ADMINISTRATIF Table
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "I. INFORMASI ADMINISTRATIF", bold: true, size: 22, font: "Arial", color: "111111" })]
    })
  ];
  const infoRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Nama Penyusun / Guru", 30, "F1F5F9", true, 18),
        makeCell(identitas.namaGuru || meta.namaGuru || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Institusi / Sekolah", 30, "F1F5F9", true, 18),
        makeCell(identitas.sekolah || meta.jenjangSekolah || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Mata Pelajaran", 30, "F1F5F9", true, 18),
        makeCell(identitas.mataPelajaran || meta.mataPelajaran || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Fase / Kelas", 30, "F1F5F9", true, 18),
        makeCell(`Fase ${meta.fase.toUpperCase()} / ${identitas.kelasSemester || "-"}`, 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Semester & Tahun Ajaran", 30, "F1F5F9", true, 18),
        makeCell(identitas.semesterTahun || `${identitas.kelasSemester?.includes("Semester") ? "" : "Semester Ganjil"} - ${meta.tahunAjaran}`, 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Topik / Materi Pokok", 30, "F1F5F9", true, 18),
        makeCell(identitas.topik || data.judul?.replace("Rencana Pembelajaran Mendalam (RPM) - Topik ", "") || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Model Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(identitas.modelPembelajaran || "Tatap Muka / Reguler", 70, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: infoRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "II. DESAIN PENGEMBANGAN DAN IMPLEMENTASI", bold: true, size: 22, font: "Arial", color: "111111" })]
    }),
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "A. IDENTIFIKASI", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );
  const identifikasiRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Aspek Identifikasi", 30, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Deskripsi Analisis Pengembangan", 70, "007C58", true, 18, false, "FFFFFF")
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Target Peserta Didik", 30, "F1F5F9", true, 18),
        makeCell(identifikasi.targetPesertaDidik || "Peserta didik reguler/tipikal kelas terkait.", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Kompetensi Prasyarat", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(identifikasi.kompetensiPrasyarat), 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Profil Lulusan (8 Dimensi)", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(identifikasi.profilLulusan), 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Sarana dan Prasarana", 30, "F1F5F9", true, 18),
        makeCell([
          `\u2022 Fasilitas Fisik: ${sarana.fasilitasFisik || "-"}`,
          `\u2022 Piranti Lunak: ${sarana.pirantiLunak || "-"}`,
          `\u2022 Sumber Belajar: ${sarana.sumberBelajar || "-"}`
        ], 70, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: identifikasiRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "B. DESAIN PEMBELAJARAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );
  const desainRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Aspek Desain", 30, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Rincian Implementasi Kurikulum Merdeka", 70, "007C58", true, 18, false, "FFFFFF")
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Capaian Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(desain.capaianPembelajaran || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Lintas Disiplin Ilmu", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(desain.lintasDisiplin), 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Tujuan Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(formatBullets(desain.tujuanPembelajaran), 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Topik Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(desain.topikPembelajaran || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Praktik Pedagogis", 30, "F1F5F9", true, 18),
        makeCell(desain.praktikPedagogis || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Kemitraan Pembelajaran", 30, "F1F5F9", true, 18),
        makeCell(desain.kemitraanPembelajaran || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Lingkungan Belajar", 30, "F1F5F9", true, 18),
        makeCell(desain.lingkunganBelajar || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Pemanfaatan Digital", 30, "F1F5F9", true, 18),
        makeCell(desain.pemanfaatanDigital || "-", 70, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: desainRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "C. LANGKAH PEMBELAJARAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );
  const formatLangkahBullets = (guruList, siswaList) => {
    const paras = [];
    paras.push(new import_docx.Paragraph({
      spacing: { before: 40, after: 20 },
      children: [new import_docx.TextRun({ text: "Aktivitas Guru:", bold: true, font: "Arial", size: 18 })]
    }));
    guruList.forEach((item) => {
      paras.push(new import_docx.Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new import_docx.TextRun({ text: `\u2022 ${item}`, font: "Arial", size: 18 })]
      }));
    });
    paras.push(new import_docx.Paragraph({
      spacing: { before: 80, after: 20 },
      children: [new import_docx.TextRun({ text: "Aktivitas Peserta Didik:", bold: true, font: "Arial", size: 18 })]
    }));
    siswaList.forEach((item) => {
      paras.push(new import_docx.Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new import_docx.TextRun({ text: `\u2022 ${item}`, font: "Arial", size: 18 })]
      }));
    });
    return paras;
  };
  const pFase = langkah.pendahuluan || { durasi: "15 Menit", guru: [], siswa: [], pendekatan: "MINDFUL" };
  const ieFase = langkah.intiEksplorasi || { durasi: "45 Menit", guru: [], siswa: [], pendekatan: "MEANINGFUL" };
  const irFase = langkah.intiRancangan || { durasi: "45 Menit", guru: [], siswa: [], pendekatan: "JOYFUL" };
  const penFase = langkah.penutup || { durasi: "15 Menit", guru: [], siswa: [], pendekatan: "MINDFUL & MEANINGFUL" };
  const langkahRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Fase Kegiatan", 20, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Aktivitas Pendidik & Peserta Didik", 65, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Pendekatan Deep Learning", 15, "007C58", true, 18, false, "FFFFFF")
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell(`Pendahuluan
(${pFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(pFase.guru || [], pFase.siswa || []), 65, void 0, false, 18),
        makeCell(pFase.pendekatan || "MINDFUL", 15, void 0, true, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell(`Eksplorasi Konsep & Analisis
(${ieFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(ieFase.guru || [], ieFase.siswa || []), 65, void 0, false, 18),
        makeCell(ieFase.pendekatan || "MEANINGFUL", 15, void 0, true, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell(`Perancangan & Simulasi
(${irFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(irFase.guru || [], irFase.siswa || []), 65, void 0, false, 18),
        makeCell(irFase.pendekatan || "JOYFUL", 15, void 0, true, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell(`Penutup & Evaluasi
(${penFase.durasi})`, 20, "F1F5F9", true, 18),
        makeCell(formatLangkahBullets(penFase.guru || [], penFase.siswa || []), 65, void 0, false, 18),
        makeCell(penFase.pendekatan || "MINDFUL & MEANINGFUL", 15, void 0, true, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: langkahRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "D. ASESMEN & PENILAIAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );
  const diagAs = asesmenObj.diagnostik || { metode: "-", instrumen: "-", kriteria: "-" };
  const formAs = asesmenObj.formatif || { metode: "-", instrumen: "-", kriteria: "-" };
  const sumAs = asesmenObj.sumatif || { metode: "-", instrumen: "-", kriteria: "-" };
  const asesmenRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Kategori Asesmen", 20, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Metode Penilaian", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Instrumen Penilaian", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Kriteria Keberhasilan", 30, "007C58", true, 18, false, "FFFFFF")
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Asesmen Diagnostik (Awal)", 20, "F1F5F9", true, 18),
        makeCell(diagAs.metode || "-", 25, void 0, false, 18),
        makeCell(diagAs.instrumen || "-", 25, void 0, false, 18),
        makeCell(diagAs.kriteria || "-", 30, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Asesmen Formatif (Proses)", 20, "F1F5F9", true, 18),
        makeCell(formAs.metode || "-", 25, void 0, false, 18),
        makeCell(formAs.instrumen || "-", 25, void 0, false, 18),
        makeCell(formAs.kriteria || "-", 30, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Asesmen Sumatif (Akhir)", 20, "F1F5F9", true, 18),
        makeCell(sumAs.metode || "-", 25, void 0, false, 18),
        makeCell(sumAs.instrumen || "-", 25, void 0, false, 18),
        makeCell(sumAs.kriteria || "-", 30, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: asesmenRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "E. REMEDIAL & PENGAYAAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );
  const remObj = remedialObj.remedial || { sasaran: "Siswa yang belum mencapai kriteria ketuntasan.", aktivitas: [] };
  const pengObj = remedialObj.pengayaan || { sasaran: "Siswa yang telah melampaui kriteria ketuntasan.", aktivitas: [] };
  const remedialRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Program Kegiatan", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Sasaran Peserta Didik", 35, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Bentuk Aktivitas Intervensi", 40, "007C58", true, 18, false, "FFFFFF")
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Program Remedial", 25, "F1F5F9", true, 18),
        makeCell(remObj.sasaran || "-", 35, void 0, false, 18),
        makeCell(formatBullets(remObj.aktivitas), 40, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Program Pengayaan", 25, "F1F5F9", true, 18),
        makeCell(pengObj.sasaran || "-", 35, void 0, false, 18),
        makeCell(formatBullets(pengObj.aktivitas), 40, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: remedialRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(
    new import_docx.Paragraph({
      spacing: { after: 60 },
      children: [new import_docx.TextRun({ text: "F. REFLEKSI", bold: true, size: 18, font: "Arial", color: "007C58" })]
    })
  );
  const sRef = refleksiObj.siswa || [];
  const gRef = refleksiObj.guru || [];
  const refleksiRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Subjek Refleksi", 25, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Pertanyaan Reflektif Pemicu Kesadaran Belajar", 75, "007C58", true, 18, false, "FFFFFF")
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Refleksi Peserta Didik", 25, "F1F5F9", true, 18),
        makeCell(formatBullets(sRef), 75, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Refleksi Pendidik (Guru)", 25, "F1F5F9", true, 18),
        makeCell(formatBullets(gRef), 75, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: refleksiRows
    }),
    new import_docx.Paragraph({ spacing: { before: 240, after: 120 } })
  );
  children.push(...createSignatures(meta));
  return new import_docx.Document({
    sections: [{ properties: {}, children }]
  });
}
function generateTopikDoc(meta, data) {
  const children = [...createDocHeader(data.judul || "DAFTAR TOPIK PEMBELAJARAN", meta)];
  children.push(
    new import_docx.Paragraph({
      heading: import_docx.HeadingLevel.HEADING_1,
      spacing: { before: 240, after: 120 },
      children: [new import_docx.TextRun({ text: "DAFTAR TOPIK / MATERI UTAMA", bold: true, size: 24, font: "Arial", color: "111111" })]
    })
  );
  const tableRows = [
    new import_docx.TableRow({
      children: [
        new import_docx.TableCell({ width: { size: 5, type: import_docx.WidthType.PERCENTAGE }, children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "No", bold: true, font: "Arial" })] })] }),
        new import_docx.TableCell({ width: { size: 25, type: import_docx.WidthType.PERCENTAGE }, children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Topik Utama", bold: true, font: "Arial" })] })] }),
        new import_docx.TableCell({ width: { size: 30, type: import_docx.WidthType.PERCENTAGE }, children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Deskripsi Cakupan", bold: true, font: "Arial" })] })] }),
        new import_docx.TableCell({ width: { size: 25, type: import_docx.WidthType.PERCENTAGE }, children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Tujuan Pembelajaran (TP)", bold: true, font: "Arial" })] })] }),
        new import_docx.TableCell({ width: { size: 15, type: import_docx.WidthType.PERCENTAGE }, children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Durasi & Sem.", bold: true, font: "Arial" })] })] })
      ]
    })
  ];
  if (data && Array.isArray(data.daftarTopik)) {
    data.daftarTopik.forEach((t) => {
      tableRows.push(
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({ children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: String(t.no), font: "Arial" })] })] }),
            new import_docx.TableCell({ children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: t.topik || "", bold: true, font: "Arial" })] })] }),
            new import_docx.TableCell({ children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: t.deskripsi || "", font: "Arial" })] })] }),
            new import_docx.TableCell({ children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: `[${t.tpKode}] ${t.tpTeks}`, font: "Arial" })] })] }),
            new import_docx.TableCell({ children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: `${t.alokasiWaktu} (Sem. ${t.semester})`, font: "Arial" })] })] })
          ]
        })
      );
    });
  }
  children.push(new import_docx.Table({ width: { size: 100, type: import_docx.WidthType.PERCENTAGE }, rows: tableRows, borders: commonBorder }));
  children.push(...createSignatures(meta));
  return new import_docx.Document({ sections: [{ properties: {}, children }] });
}
function generateLKPDDoc(meta, data) {
  const children = [];
  function makeCell(content, widthPercent, shadingColor, boldText = false, textSize = 18, isItalic = false, textColor = "111111", align = import_docx.AlignmentType.LEFT, cellBorders = commonBorder) {
    let cellChildren = [];
    if (Array.isArray(content)) {
      if (content.length === 0) {
        cellChildren.push(new import_docx.Paragraph({
          alignment: align,
          children: [new import_docx.TextRun({ text: "-", font: "Arial", size: textSize, color: textColor })]
        }));
      } else if (typeof content[0] === "string") {
        content.forEach((str) => {
          cellChildren.push(new import_docx.Paragraph({
            alignment: align,
            spacing: { before: 40, after: 40 },
            children: [new import_docx.TextRun({ text: str, font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
          }));
        });
      } else {
        cellChildren = content;
      }
    } else if (typeof content === "string") {
      cellChildren.push(new import_docx.Paragraph({
        alignment: align,
        spacing: { before: 40, after: 40 },
        children: [new import_docx.TextRun({ text: content || "-", font: "Arial", size: textSize, bold: boldText, italics: isItalic, color: textColor })]
      }));
    } else {
      cellChildren.push(content);
    }
    return new import_docx.TableCell({
      width: { size: widthPercent, type: import_docx.WidthType.PERCENTAGE },
      shading: shadingColor ? { fill: shadingColor } : void 0,
      borders: cellBorders,
      margins: { top: 100, bottom: 100, left: 150, right: 150 },
      children: cellChildren
    });
  }
  function createAccentHeader(text) {
    return new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      borders: {
        top: { style: import_docx.BorderStyle.NONE },
        bottom: { style: import_docx.BorderStyle.NONE },
        left: { style: import_docx.BorderStyle.NONE },
        right: { style: import_docx.BorderStyle.NONE }
      },
      rows: [
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              borders: {
                top: { style: import_docx.BorderStyle.NONE },
                bottom: { style: import_docx.BorderStyle.NONE },
                left: { style: import_docx.BorderStyle.SINGLE, size: 24, color: "007A54" },
                // Green bar
                right: { style: import_docx.BorderStyle.NONE }
              },
              margins: { left: 120, top: 120, bottom: 120 },
              children: [
                new import_docx.Paragraph({
                  children: [
                    new import_docx.TextRun({ text: text.toUpperCase(), bold: true, size: 22, font: "Arial", color: "007A54" })
                  ]
                })
              ]
            })
          ]
        })
      ]
    });
  }
  const titleBannerTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE },
      bottom: { style: import_docx.BorderStyle.NONE },
      left: { style: import_docx.BorderStyle.NONE },
      right: { style: import_docx.BorderStyle.NONE }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "007C58" },
            margins: { top: 240, bottom: 240, left: 150, right: 150 },
            borders: {
              top: { style: import_docx.BorderStyle.NONE },
              bottom: { style: import_docx.BorderStyle.NONE },
              left: { style: import_docx.BorderStyle.NONE },
              right: { style: import_docx.BorderStyle.NONE }
            },
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new import_docx.TextRun({
                    text: data.judul || "LEMBAR KERJA PESERTA DIDIK (LKPD)",
                    bold: true,
                    size: 26,
                    font: "Arial",
                    color: "FFFFFF"
                  })
                ]
              }),
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                children: [
                  new import_docx.TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran || "-"} | Fase/Kelas: ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`,
                    bold: true,
                    size: 18,
                    font: "Arial",
                    color: "FFFFFF"
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(titleBannerTable);
  children.push(new import_docx.Paragraph({ spacing: { after: 120 } }));
  const metadataRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Nama Guru / Penyusun", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.namaGuru || meta.namaGuru || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Instansi / Sekolah", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.sekolah || meta.jenjangSekolah || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Mata Pelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.mataPelajaran || meta.mataPelajaran || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Jenjang / Fase / Kelas", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(`SMK/MAK / ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`, 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Semester & Tahun Ajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.semesterTahun || `Ganjil - ${meta.tahunAjaran || ""}`, 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Topik / Materi Pokok", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.topikMateri || data.topik || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Model Pembelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.modelPembelajaran || "Tatap Muka / Reguler", 70, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: metadataRows
    })
  );
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(createAccentHeader("TUJUAN & PETUNJUK"));
  children.push(new import_docx.Paragraph({ spacing: { after: 120 } }));
  const tpList = Array.isArray(data.tujuanPembelajaran) ? data.tujuanPembelajaran : [];
  const tpParagraphs = tpList.map((item) => new import_docx.Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new import_docx.TextRun({ text: "\u2022  ", font: "Arial", size: 18, bold: true }),
      new import_docx.TextRun({ text: item, font: "Arial", size: 18 })
    ]
  }));
  const petunjukList = Array.isArray(data.petunjukPengerjaan) ? data.petunjukPengerjaan : Array.isArray(data.petunjukBelajar) ? data.petunjukBelajar : [];
  const petunjukParagraphs = petunjukList.map((item, idx) => new import_docx.Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new import_docx.TextRun({ text: `${idx + 1}.  `, font: "Arial", size: 18, bold: true }),
      new import_docx.TextRun({ text: item, font: "Arial", size: 18 })
    ]
  }));
  const tujuanPetunjukTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: [
      new import_docx.TableRow({
        children: [
          makeCell("Aspek Kerja", 30, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.LEFT),
          makeCell("Deskripsi Panduan", 70, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.LEFT)
        ]
      }),
      new import_docx.TableRow({
        children: [
          makeCell("Tujuan Pembelajaran", 30, "F0FDF4", true, 18, false, "111111"),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            borders: commonBorder,
            children: tpParagraphs.length > 0 ? tpParagraphs : [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "-", font: "Arial", size: 18 })] })]
          })
        ]
      }),
      new import_docx.TableRow({
        children: [
          makeCell("Petunjuk Pengerjaan", 30, "F0FDF4", true, 18, false, "111111"),
          new import_docx.TableCell({
            width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            borders: commonBorder,
            children: petunjukParagraphs.length > 0 ? petunjukParagraphs : [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "-", font: "Arial", size: 18 })] })]
          })
        ]
      })
    ]
  });
  children.push(tujuanPetunjukTable);
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(createAccentHeader("RINGKASAN MATERI"));
  children.push(new import_docx.Paragraph({ spacing: { after: 120 } }));
  const ringkasanRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Komponen Utama", 30, "007C58", true, 18, false, "FFFFFF"),
        makeCell("Ringkasan Konseptual " + (data.topik || "Materi"), 70, "007C58", true, 18, false, "FFFFFF")
      ]
    })
  ];
  if (Array.isArray(data.ringkasanMateri)) {
    data.ringkasanMateri.forEach((item) => {
      let ringkasanParagraphs = [];
      const lines = typeof item.ringkasan === "string" ? item.ringkasan.split("\n") : [];
      if (lines.length > 1) {
        lines.forEach((line) => {
          const cleanLine = line.trim();
          if (cleanLine.startsWith("\u2022") || cleanLine.startsWith("-")) {
            const textPart = cleanLine.replace(/^[•\-]\s*/, "");
            const boldMatch = textPart.match(/^([^:]+):(.*)$/);
            if (boldMatch) {
              ringkasanParagraphs.push(new import_docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new import_docx.TextRun({ text: "\u2022  ", font: "Arial", size: 18, bold: true }),
                  new import_docx.TextRun({ text: boldMatch[1] + ": ", font: "Arial", size: 18, bold: true }),
                  new import_docx.TextRun({ text: boldMatch[2], font: "Arial", size: 18 })
                ]
              }));
            } else {
              ringkasanParagraphs.push(new import_docx.Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new import_docx.TextRun({ text: "\u2022  ", font: "Arial", size: 18, bold: true }),
                  new import_docx.TextRun({ text: textPart, font: "Arial", size: 18 })
                ]
              }));
            }
          } else {
            ringkasanParagraphs.push(new import_docx.Paragraph({
              spacing: { before: 40, after: 40 },
              children: [new import_docx.TextRun({ text: cleanLine, font: "Arial", size: 18 })]
            }));
          }
        });
      } else {
        ringkasanParagraphs.push(new import_docx.Paragraph({
          spacing: { before: 40, after: 40 },
          children: [new import_docx.TextRun({ text: item.ringkasan || "-", font: "Arial", size: 18 })]
        }));
      }
      ringkasanRows.push(
        new import_docx.TableRow({
          children: [
            makeCell(item.komponen || "-", 30, "F0FDF4", true, 18, false, "111111"),
            new import_docx.TableCell({
              width: { size: 70, type: import_docx.WidthType.PERCENTAGE },
              margins: { top: 100, bottom: 100, left: 150, right: 150 },
              borders: commonBorder,
              children: ringkasanParagraphs
            })
          ]
        })
      );
    });
  }
  const ringkasanTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    rows: ringkasanRows
  });
  children.push(ringkasanTable);
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(createAccentHeader("AKTIVITAS / SOAL"));
  children.push(new import_docx.Paragraph({ spacing: { after: 120 } }));
  if (Array.isArray(data.aktivitasList)) {
    data.aktivitasList.forEach((item, idx) => {
      const activityTable = new import_docx.Table({
        width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
        rows: [
          new import_docx.TableRow({
            children: [
              makeCell(item.judul || `Aktivitas ${idx + 1}`, 100, "007C58", true, 20, false, "FFFFFF")
            ]
          }),
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
                shading: { fill: "F0FDF4" },
                margins: { top: 150, bottom: 150, left: 150, right: 150 },
                borders: commonBorder,
                children: (item.studiKasus || "").split("\n").map((line) => {
                  const cleanLine = line.trim();
                  const isStudiKasusHeader = cleanLine.toLowerCase().startsWith("studi kasus:");
                  return new import_docx.Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [
                      new import_docx.TextRun({
                        text: cleanLine,
                        font: "Arial",
                        size: 18,
                        bold: isStudiKasusHeader
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
      children.push(new import_docx.Paragraph({ spacing: { after: 180 } }));
    });
  } else if (data.tugasAktivitas && Array.isArray(data.tugasAktivitas.pertanyaan)) {
    const legacyActivityTable = new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: [
        new import_docx.TableRow({
          children: [
            makeCell(data.judul || "Aktivitas Pembelajaran", 100, "007C58", true, 20, false, "FFFFFF")
          ]
        }),
        new import_docx.TableRow({
          children: [
            new import_docx.TableCell({
              width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
              shading: { fill: "F0FDF4" },
              margins: { top: 150, bottom: 150, left: 150, right: 150 },
              borders: commonBorder,
              children: [
                new import_docx.Paragraph({
                  spacing: { before: 40, after: 80 },
                  children: [new import_docx.TextRun({ text: "Studi Kasus / Instruksi Kerja:", bold: true, font: "Arial", size: 18 })]
                }),
                new import_docx.Paragraph({
                  spacing: { after: 120 },
                  children: [new import_docx.TextRun({ text: data.tugasAktivitas.instruksi || "", font: "Arial", size: 18 })]
                }),
                ...data.tugasAktivitas.pertanyaan.map((q, idx) => {
                  return new import_docx.Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [
                      new import_docx.TextRun({ text: `Pertanyaan ${idx + 1}: `, bold: true, font: "Arial", size: 18 }),
                      new import_docx.TextRun({ text: q, font: "Arial", size: 18 })
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
  return new import_docx.Document({ sections: [{ properties: {}, children }] });
}
function generateAsesmenDoc(meta, data) {
  function makeCell(content, widthPercent, shadingColor, boldText = false, textSize = 18, isItalic = false, textColor = "111111", align = import_docx.AlignmentType.LEFT, cellBorders = commonBorder) {
    let cellChildren = [];
    if (Array.isArray(content)) {
      if (content.length === 0) {
        cellChildren.push(new import_docx.Paragraph({
          alignment: align,
          children: [new import_docx.TextRun({ text: "-", font: "Arial", size: textSize, color: textColor })]
        }));
      } else {
        if (content[0] instanceof import_docx.Paragraph) {
          cellChildren = content;
        } else {
          content.forEach((line) => {
            cellChildren.push(new import_docx.Paragraph({
              alignment: align,
              children: [new import_docx.TextRun({ text: line, font: "Arial", size: textSize, color: textColor })]
            }));
          });
        }
      }
    } else {
      cellChildren.push(new import_docx.Paragraph({
        alignment: align,
        children: [new import_docx.TextRun({
          text: content,
          bold: boldText,
          italics: isItalic,
          size: textSize,
          font: "Arial",
          color: textColor
        })]
      }));
    }
    return new import_docx.TableCell({
      width: { size: widthPercent, type: import_docx.WidthType.PERCENTAGE },
      shading: shadingColor ? { fill: shadingColor } : void 0,
      margins: { top: 120, bottom: 120, left: 150, right: 150 },
      borders: cellBorders,
      children: cellChildren
    });
  }
  const children = [];
  const titleBannerTable = new import_docx.Table({
    width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: import_docx.BorderStyle.NONE },
      bottom: { style: import_docx.BorderStyle.NONE },
      left: { style: import_docx.BorderStyle.NONE },
      right: { style: import_docx.BorderStyle.NONE }
    },
    rows: [
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            shading: { fill: "007C58" },
            margins: { top: 240, bottom: 240, left: 150, right: 150 },
            borders: {
              top: { style: import_docx.BorderStyle.NONE },
              bottom: { style: import_docx.BorderStyle.NONE },
              left: { style: import_docx.BorderStyle.NONE },
              right: { style: import_docx.BorderStyle.NONE }
            },
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new import_docx.TextRun({
                    text: "KISI-KISI DAN SOAL ASESMEN",
                    bold: true,
                    size: 26,
                    font: "Arial",
                    color: "FFFFFF"
                  })
                ]
              }),
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.CENTER,
                children: [
                  new import_docx.TextRun({
                    text: `Mata Pelajaran: ${meta.mataPelajaran || "-"} | Fase/Kelas: ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`,
                    bold: true,
                    size: 18,
                    font: "Arial",
                    color: "FFFFFF"
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
  children.push(titleBannerTable);
  children.push(new import_docx.Paragraph({ spacing: { after: 120 } }));
  const identitasRows = [
    new import_docx.TableRow({
      children: [
        makeCell("Nama Guru / Penyusun", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.namaGuru || meta.namaGuru || "Rudi Akbar Saragih, S.Kom.", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Instansi / Sekolah", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.sekolah || meta.jenjangSekolah || "SMKS TI Muhammadiyah 11 Sibuluan", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Mata Pelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.mataPelajaran || meta.mataPelajaran || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Jenjang / Fase / Kelas", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(`SMK/MAK / ${meta.fase ? meta.fase.toUpperCase() : "F"} / ${meta.fase === "f" ? "XI" : "X"}`, 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Semester & Tahun Ajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.semesterTahun || `Ganjil - ${meta.tahunAjaran || "2026/2027"}`, 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Topik / Materi Pokok", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.topikMateri || data.topik || "-", 70, void 0, false, 18)
      ]
    }),
    new import_docx.TableRow({
      children: [
        makeCell("Model Pembelajaran", 30, "E6F2ED", true, 18, false, "111111"),
        makeCell(data.identitas?.modelPembelajaran || "Tatap Muka / Reguler", 70, void 0, false, 18)
      ]
    })
  ];
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: identitasRows
    })
  );
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new import_docx.TextRun({ text: "KISI-KISI SOAL", bold: true, size: 22, font: "Arial", color: "007C58" })
      ]
    })
  );
  const kisiKisiRows = [
    new import_docx.TableRow({
      children: [
        makeCell("No", 5, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Tujuan Pembelajaran (TP)", 25, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Indikator Soal", 45, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Level Kognitif", 12, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Bentuk Soal", 8, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("No Soal", 5, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER)
      ]
    })
  ];
  const listKisi = data.kisiKisi || [];
  listKisi.forEach((k) => {
    kisiKisiRows.push(
      new import_docx.TableRow({
        children: [
          makeCell(String(k.no), 5, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
          makeCell(k.tujuanPembelajaran || "-", 25, void 0, false, 18),
          makeCell(k.indikatorSoal || "-", 45, void 0, false, 18),
          makeCell(k.levelKognitif || "-", 12, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
          makeCell(k.bentukSoal || "-", 8, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
          makeCell(String(k.noSoal), 5, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER)
        ]
      })
    );
  });
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: kisiKisiRows
    })
  );
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new import_docx.TextRun({ text: "BUTIR SOAL ASESMEN", bold: true, size: 22, font: "Arial", color: "007C58" })
      ]
    })
  );
  const butirSoalRows = [
    new import_docx.TableRow({
      children: [
        makeCell("No", 5, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Tipe & Level", 15, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Butir Soal & Pilihan Jawaban / Instruksi Soal", 80, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER)
      ]
    })
  ];
  const listSoal = data.daftarSoal || data.soal || [];
  listSoal.forEach((s) => {
    const contentParagraphs = [];
    contentParagraphs.push(new import_docx.Paragraph({
      alignment: import_docx.AlignmentType.LEFT,
      spacing: { after: 80 },
      children: [new import_docx.TextRun({ text: s.soal || s.pertanyaan || "-", font: "Arial", size: 18 })]
    }));
    if (s.pilihan) {
      Object.keys(s.pilihan).forEach((opt) => {
        contentParagraphs.push(new import_docx.Paragraph({
          alignment: import_docx.AlignmentType.LEFT,
          spacing: { before: 20, after: 20 },
          indent: { left: 240 },
          children: [
            new import_docx.TextRun({ text: `${opt}. `, bold: true, font: "Arial", size: 18 }),
            new import_docx.TextRun({ text: s.pilihan[opt] || "", font: "Arial", size: 18 })
          ]
        }));
      });
    } else if (s.tipe === "Pilihan Ganda" || !s.tipe && (s.a || s.b || s.c || s.d || s.e)) {
      const opts = ["A", "B", "C", "D", "E"];
      opts.forEach((opt) => {
        const val = s[opt.toLowerCase()];
        if (val) {
          contentParagraphs.push(new import_docx.Paragraph({
            alignment: import_docx.AlignmentType.LEFT,
            spacing: { before: 20, after: 20 },
            indent: { left: 240 },
            children: [
              new import_docx.TextRun({ text: `${opt}. `, bold: true, font: "Arial", size: 18 }),
              new import_docx.TextRun({ text: val, font: "Arial", size: 18 })
            ]
          }));
        }
      });
    }
    const tipeLevel = `${s.tipe || "Pilihan Ganda"}
${s.levelKognitif || ""}`;
    butirSoalRows.push(
      new import_docx.TableRow({
        children: [
          makeCell(String(s.no), 5, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
          makeCell(tipeLevel, 15, void 0, true, 18, false, "111111", import_docx.AlignmentType.CENTER),
          new import_docx.TableCell({
            width: { size: 80, type: import_docx.WidthType.PERCENTAGE },
            margins: { top: 120, bottom: 120, left: 150, right: 150 },
            borders: commonBorder,
            children: contentParagraphs
          })
        ]
      })
    );
  });
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: butirSoalRows
    })
  );
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(
    new import_docx.Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new import_docx.TextRun({ text: "KUNCI JAWABAN & RUBRIK PENILAIAN", bold: true, size: 22, font: "Arial", color: "007C58" })
      ]
    })
  );
  const kunciJawabanRows = [
    new import_docx.TableRow({
      children: [
        makeCell("No Soal", 10, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Bentuk Soal", 15, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Kunci Jawaban / Kriteria Penilaian Rubrik", 65, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER),
        makeCell("Skor Maksimal", 10, "007C58", true, 18, false, "FFFFFF", import_docx.AlignmentType.CENTER)
      ]
    })
  ];
  const pgSoal = listSoal.filter((s) => s.tipe === "Pilihan Ganda" || !s.tipe && (s.pilihan || s.a));
  const uraianSoal = listSoal.filter((s) => s.tipe === "Uraian" || !s.tipe && !s.pilihan && !s.a);
  if (pgSoal.length > 0) {
    kunciJawabanRows.push(
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            columnSpan: 4,
            shading: { fill: "E6F2ED" },
            margins: { top: 80, bottom: 80, left: 150, right: 150 },
            borders: commonBorder,
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.LEFT,
                children: [new import_docx.TextRun({ text: "BAGIAN I: PILIHAN GANDA", bold: true, size: 18, font: "Arial", color: "007C58" })]
              })
            ]
          })
        ]
      })
    );
    pgSoal.forEach((s) => {
      kunciJawabanRows.push(
        new import_docx.TableRow({
          children: [
            makeCell(String(s.no), 10, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
            makeCell(s.tipe || "Pilihan Ganda", 15, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
            makeCell(s.kunciJawaban || "-", 65, void 0, false, 18),
            makeCell(String(s.skorMaksimal || 5), 10, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER)
          ]
        })
      );
    });
  }
  if (uraianSoal.length > 0) {
    kunciJawabanRows.push(
      new import_docx.TableRow({
        children: [
          new import_docx.TableCell({
            columnSpan: 4,
            shading: { fill: "E6F2ED" },
            margins: { top: 80, bottom: 80, left: 150, right: 150 },
            borders: commonBorder,
            children: [
              new import_docx.Paragraph({
                alignment: import_docx.AlignmentType.LEFT,
                children: [new import_docx.TextRun({ text: "BAGIAN II: URAIAN", bold: true, size: 18, font: "Arial", color: "007C58" })]
              })
            ]
          })
        ]
      })
    );
    uraianSoal.forEach((s) => {
      const lines = (s.kunciJawaban || s.pembahasan || "-").split("\n");
      const answerParagraphs = lines.map((l) => new import_docx.Paragraph({
        alignment: import_docx.AlignmentType.LEFT,
        spacing: { before: 20, after: 20 },
        children: [new import_docx.TextRun({ text: l.trim(), font: "Arial", size: 18 })]
      }));
      kunciJawabanRows.push(
        new import_docx.TableRow({
          children: [
            makeCell(String(s.no), 10, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
            makeCell(s.tipe || "Uraian", 15, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER),
            new import_docx.TableCell({
              width: { size: 65, type: import_docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 150, right: 150 },
              borders: commonBorder,
              children: answerParagraphs
            }),
            makeCell(String(s.skorMaksimal || 15), 10, void 0, false, 18, false, "111111", import_docx.AlignmentType.CENTER)
          ]
        })
      );
    });
  }
  const totalScore = listSoal.reduce((acc, s) => acc + (s.skorMaksimal || 5), 0);
  kunciJawabanRows.push(
    new import_docx.TableRow({
      children: [
        new import_docx.TableCell({
          columnSpan: 3,
          shading: { fill: "E6F2ED" },
          margins: { top: 100, bottom: 100, left: 150, right: 150 },
          borders: commonBorder,
          children: [
            new import_docx.Paragraph({
              alignment: import_docx.AlignmentType.RIGHT,
              children: [new import_docx.TextRun({ text: "TOTAL SKOR MAKSIMAL KESELURUHAN SOAL", bold: true, size: 18, font: "Arial", color: "007C58" })]
            })
          ]
        }),
        makeCell(String(totalScore || 100), 10, "E6F2ED", true, 18, false, "007C58", import_docx.AlignmentType.CENTER)
      ]
    })
  );
  children.push(
    new import_docx.Table({
      width: { size: 100, type: import_docx.WidthType.PERCENTAGE },
      rows: kunciJawabanRows
    })
  );
  children.push(new import_docx.Paragraph({ spacing: { after: 240 } }));
  children.push(...createSignatures(meta));
  return new import_docx.Document({ sections: [{ properties: {}, children }] });
}

// server.ts
import_dotenv.default.config();
var app = (0, import_express.default)();
var PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3e3;
app.use(import_express.default.json({ limit: "50mb" }));
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
var apiKey = process.env.GEMINI_API_KEY;
if (!apiKey) {
  console.warn("WARNING: GEMINI_API_KEY is not defined in environment variables. Gemini calls will fail.");
}
var ai = new import_genai.GoogleGenAI({
  apiKey: apiKey || "MOCK_KEY_IF_MISSING_TO_PREVENT_LOAD_ERROR",
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build"
    }
  }
});
function ensureApiKey(res) {
  if (!process.env.GEMINI_API_KEY) {
    res.status(500).json({
      error: "GEMINI_API_KEY is missing. Please configure it in the Secrets panel in the AI Studio UI under Settings > Secrets."
    });
    return false;
  }
  return true;
}
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
    tahunAjaran
  } = req.body;
  const isReligious = /agama/i.test(mataPelajaran || "") || /islam/i.test(mataPelajaran || "") || /kristen/i.test(mataPelajaran || "") || /katolik/i.test(mataPelajaran || "") || /hindu/i.test(mataPelajaran || "") || /buddha/i.test(mataPelajaran || "") || /konghucu/i.test(mataPelajaran || "") || /budi pekerti/i.test(mataPelajaran || "");
  const regulasiUtama = isReligious ? "Keputusan Kepala BKPDM Nomor 020 Tahun 2026 tentang Capaian Pembelajaran Pendidikan Agama dan Budi Pekerti" : "Keputusan Kepala BSKAP Nomor 046/H/KR/2025 tentang Capaian Pembelajaran pada Pendidikan Anak Usia Dini, Jenjang Pendidikan Dasar, dan Jenjang Pendidikan Menengah pada Kurikulum Merdeka";
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
          type: import_genai.Type.OBJECT,
          properties: {
            judul: { type: import_genai.Type.STRING },
            regulasi: { type: import_genai.Type.STRING },
            rasional: { type: import_genai.Type.STRING },
            tujuan: { type: import_genai.Type.STRING },
            karakteristik: { type: import_genai.Type.STRING },
            capaianUmum: { type: import_genai.Type.STRING },
            elemen: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  nama: { type: import_genai.Type.STRING },
                  deskripsi: { type: import_genai.Type.STRING },
                  capaian: { type: import_genai.Type.STRING }
                },
                required: ["nama", "deskripsi", "capaian"]
              }
            }
          },
          required: ["judul", "regulasi", "rasional", "tujuan", "karakteristik", "capaianUmum", "elemen"]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating CP:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Capaian Pembelajaran." });
  }
});
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
    ${cpData.elemen.map((el) => `- Elemen: ${el.nama}
  Capaian Elemen: ${el.capaian}`).join("\n")}
    
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
          type: import_genai.Type.OBJECT,
          properties: {
            judul: { type: import_genai.Type.STRING },
            fase: { type: import_genai.Type.STRING },
            mataPelajaran: { type: import_genai.Type.STRING },
            atp: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  elemen: { type: import_genai.Type.STRING },
                  tujuanPembelajaran: {
                    type: import_genai.Type.ARRAY,
                    items: {
                      type: import_genai.Type.OBJECT,
                      properties: {
                        kode: { type: import_genai.Type.STRING },
                        tp: { type: import_genai.Type.STRING },
                        topik: { type: import_genai.Type.STRING },
                        alokasiWaktu: { type: import_genai.Type.STRING },
                        profilLulusan: { type: import_genai.Type.STRING },
                        glosarium: { type: import_genai.Type.STRING }
                      },
                      required: ["kode", "tp", "topik", "alokasiWaktu", "profilLulusan", "glosarium"]
                    }
                  }
                },
                required: ["elemen", "tujuanPembelajaran"]
              }
            }
          },
          required: ["judul", "fase", "mataPelajaran", "atp"]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating ATP:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Analisis CP & ATP." });
  }
});
app.post("/api/generate/prota-promes", async (req, res) => {
  if (!ensureApiKey(res)) return;
  const { meta, atpData } = req.body;
  const tpList = [];
  atpData.atp.forEach((item) => {
    item.tujuanPembelajaran.forEach((tpObj) => {
      tpList.push({
        elemen: item.elemen,
        kode: tpObj.kode,
        tp: tpObj.tp,
        topik: tpObj.topik,
        alokasiWaktu: tpObj.alokasiWaktu
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
          type: import_genai.Type.OBJECT,
          properties: {
            judulProta: { type: import_genai.Type.STRING },
            judulPromes: { type: import_genai.Type.STRING },
            prota: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  no: { type: import_genai.Type.INTEGER },
                  elemen: { type: import_genai.Type.STRING },
                  tujuanPembelajaran: { type: import_genai.Type.STRING },
                  alokasiWaktu: { type: import_genai.Type.STRING }
                },
                required: ["no", "elemen", "tujuanPembelajaran", "alokasiWaktu"]
              }
            },
            promes: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  no: { type: import_genai.Type.INTEGER },
                  tujuanPembelajaran: { type: import_genai.Type.STRING },
                  alokasiWaktu: { type: import_genai.Type.INTEGER },
                  distribusi: {
                    type: import_genai.Type.OBJECT,
                    properties: {
                      Jul: { type: import_genai.Type.STRING },
                      Agt: { type: import_genai.Type.STRING },
                      Sep: { type: import_genai.Type.STRING },
                      Okt: { type: import_genai.Type.STRING },
                      Nov: { type: import_genai.Type.STRING },
                      Des: { type: import_genai.Type.STRING },
                      Jan: { type: import_genai.Type.STRING },
                      Feb: { type: import_genai.Type.STRING },
                      Mar: { type: import_genai.Type.STRING },
                      Apr: { type: import_genai.Type.STRING },
                      Mei: { type: import_genai.Type.STRING },
                      Jun: { type: import_genai.Type.STRING }
                    },
                    required: ["Jul", "Agt", "Sep", "Okt", "Nov", "Des", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun"]
                  }
                },
                required: ["no", "tujuanPembelajaran", "alokasiWaktu", "distribusi"]
              }
            },
            jumlahJpSemester1: { type: import_genai.Type.INTEGER },
            jumlahJpSemester2: { type: import_genai.Type.INTEGER }
          },
          required: ["judulProta", "judulPromes", "prota", "promes", "jumlahJpSemester1", "jumlahJpSemester2"]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating Prota/Promes:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate PROTA & PROMES." });
  }
});
app.post("/api/generate/kktp", async (req, res) => {
  if (!ensureApiKey(res)) return;
  const { meta, atpData } = req.body;
  const tpList = [];
  atpData.atp.forEach((item) => {
    item.tujuanPembelajaran.forEach((tpObj) => {
      tpList.push({
        kode: tpObj.kode,
        tp: tpObj.tp,
        topik: tpObj.topik
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
          type: import_genai.Type.OBJECT,
          properties: {
            judul: { type: import_genai.Type.STRING },
            kktp: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  tpKode: { type: import_genai.Type.STRING },
                  tpTeks: { type: import_genai.Type.STRING },
                  kriteria: {
                    type: import_genai.Type.ARRAY,
                    items: {
                      type: import_genai.Type.OBJECT,
                      properties: {
                        aspek: { type: import_genai.Type.STRING },
                        baruBerkembang: { type: import_genai.Type.STRING },
                        layak: { type: import_genai.Type.STRING },
                        cakap: { type: import_genai.Type.STRING },
                        mahir: { type: import_genai.Type.STRING }
                      },
                      required: ["aspek", "baruBerkembang", "layak", "cakap", "mahir"]
                    }
                  },
                  intervalNilai: {
                    type: import_genai.Type.OBJECT,
                    properties: {
                      kurang: { type: import_genai.Type.STRING },
                      cukup: { type: import_genai.Type.STRING },
                      baik: { type: import_genai.Type.STRING },
                      sangatBaik: { type: import_genai.Type.STRING }
                    },
                    required: ["kurang", "cukup", "baik", "sangatBaik"]
                  }
                },
                required: ["tpKode", "tpTeks", "kriteria", "intervalNilai"]
              }
            }
          },
          required: ["judul", "kktp"]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating KKTP:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate KKTP." });
  }
});
app.post("/api/generate/topik", async (req, res) => {
  if (!ensureApiKey(res)) return;
  const { meta, atpData } = req.body;
  const tpList = [];
  atpData.atp.forEach((item) => {
    item.tujuanPembelajaran.forEach((tpObj) => {
      tpList.push({
        elemen: item.elemen,
        kode: tpObj.kode,
        tp: tpObj.tp,
        topik: tpObj.topik,
        alokasiWaktu: tpObj.alokasiWaktu
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
          type: import_genai.Type.OBJECT,
          properties: {
            judul: { type: import_genai.Type.STRING },
            daftarTopik: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  no: { type: import_genai.Type.INTEGER },
                  topik: { type: import_genai.Type.STRING },
                  deskripsi: { type: import_genai.Type.STRING },
                  tpKode: { type: import_genai.Type.STRING },
                  tpTeks: { type: import_genai.Type.STRING },
                  alokasiWaktu: { type: import_genai.Type.STRING },
                  semester: { type: import_genai.Type.INTEGER }
                },
                required: ["no", "topik", "deskripsi", "tpKode", "tpTeks", "alokasiWaktu", "semester"]
              }
            }
          },
          required: ["judul", "daftarTopik"]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating Topik:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Daftar Topik." });
  }
});
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
       - kelasSemester: Kelas ${meta.fase === "f" ? "XI/XII" : meta.fase === "e" ? "X" : "VII/VIII/IX"} / Semester ${topikItem.semester === 1 ? "Ganjil" : "Genap"}
       - semesterTahun: ${topikItem.semester === 1 ? "Ganjil" : "Genap"} - ${meta.tahunAjaran}
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
          type: import_genai.Type.OBJECT,
          properties: {
            judul: { type: import_genai.Type.STRING },
            identitas: {
              type: import_genai.Type.OBJECT,
              properties: {
                namaGuru: { type: import_genai.Type.STRING },
                sekolah: { type: import_genai.Type.STRING },
                mataPelajaran: { type: import_genai.Type.STRING },
                kelasSemester: { type: import_genai.Type.STRING },
                semesterTahun: { type: import_genai.Type.STRING },
                topik: { type: import_genai.Type.STRING },
                modelPembelajaran: { type: import_genai.Type.STRING }
              },
              required: ["namaGuru", "sekolah", "mataPelajaran", "kelasSemester", "semesterTahun", "topik", "modelPembelajaran"]
            },
            identifikasi: {
              type: import_genai.Type.OBJECT,
              properties: {
                targetPesertaDidik: { type: import_genai.Type.STRING },
                kompetensiPrasyarat: {
                  type: import_genai.Type.ARRAY,
                  items: { type: import_genai.Type.STRING }
                },
                profilLulusan: {
                  type: import_genai.Type.ARRAY,
                  items: { type: import_genai.Type.STRING }
                },
                saranaPrasarana: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    fasilitasFisik: { type: import_genai.Type.STRING },
                    pirantiLunak: { type: import_genai.Type.STRING },
                    sumberBelajar: { type: import_genai.Type.STRING }
                  },
                  required: ["fasilitasFisik", "pirantiLunak", "sumberBelajar"]
                }
              },
              required: ["targetPesertaDidik", "kompetensiPrasyarat", "profilLulusan", "saranaPrasarana"]
            },
            desainPembelajaran: {
              type: import_genai.Type.OBJECT,
              properties: {
                capaianPembelajaran: { type: import_genai.Type.STRING },
                lintasDisiplin: {
                  type: import_genai.Type.ARRAY,
                  items: { type: import_genai.Type.STRING }
                },
                tujuanPembelajaran: {
                  type: import_genai.Type.ARRAY,
                  items: { type: import_genai.Type.STRING }
                },
                topikPembelajaran: { type: import_genai.Type.STRING },
                praktikPedagogis: { type: import_genai.Type.STRING },
                kemitraanPembelajaran: { type: import_genai.Type.STRING },
                lingkunganBelajar: { type: import_genai.Type.STRING },
                pemanfaatanDigital: { type: import_genai.Type.STRING }
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
              ]
            },
            langkahPembelajaran: {
              type: import_genai.Type.OBJECT,
              properties: {
                pendahuluan: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    durasi: { type: import_genai.Type.STRING },
                    guru: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    siswa: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    pendekatan: { type: import_genai.Type.STRING }
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"]
                },
                intiEksplorasi: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    durasi: { type: import_genai.Type.STRING },
                    guru: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    siswa: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    pendekatan: { type: import_genai.Type.STRING }
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"]
                },
                intiRancangan: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    durasi: { type: import_genai.Type.STRING },
                    guru: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    siswa: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    pendekatan: { type: import_genai.Type.STRING }
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"]
                },
                penutup: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    durasi: { type: import_genai.Type.STRING },
                    guru: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    siswa: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                    pendekatan: { type: import_genai.Type.STRING }
                  },
                  required: ["durasi", "guru", "siswa", "pendekatan"]
                }
              },
              required: ["pendahuluan", "intiEksplorasi", "intiRancangan", "penutup"]
            },
            asesmen: {
              type: import_genai.Type.OBJECT,
              properties: {
                diagnostik: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    metode: { type: import_genai.Type.STRING },
                    instrumen: { type: import_genai.Type.STRING },
                    kriteria: { type: import_genai.Type.STRING }
                  },
                  required: ["metode", "instrumen", "kriteria"]
                },
                formatif: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    metode: { type: import_genai.Type.STRING },
                    instrumen: { type: import_genai.Type.STRING },
                    kriteria: { type: import_genai.Type.STRING }
                  },
                  required: ["metode", "instrumen", "kriteria"]
                },
                sumatif: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    metode: { type: import_genai.Type.STRING },
                    instrumen: { type: import_genai.Type.STRING },
                    kriteria: { type: import_genai.Type.STRING }
                  },
                  required: ["metode", "instrumen", "kriteria"]
                }
              },
              required: ["diagnostik", "formatif", "sumatif"]
            },
            remedialPengayaan: {
              type: import_genai.Type.OBJECT,
              properties: {
                remedial: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    sasaran: { type: import_genai.Type.STRING },
                    aktivitas: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } }
                  },
                  required: ["sasaran", "aktivitas"]
                },
                pengayaan: {
                  type: import_genai.Type.OBJECT,
                  properties: {
                    sasaran: { type: import_genai.Type.STRING },
                    aktivitas: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } }
                  },
                  required: ["sasaran", "aktivitas"]
                }
              },
              required: ["remedial", "pengayaan"]
            },
            refleksi: {
              type: import_genai.Type.OBJECT,
              properties: {
                siswa: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } },
                guru: { type: import_genai.Type.ARRAY, items: { type: import_genai.Type.STRING } }
              },
              required: ["siswa", "guru"]
            }
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
          ]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating RPM-single:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Rencana Pembelajaran Mendalam." });
  }
});
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
        "kelasSemester": "Kelas ${meta.fase === "f" ? "XI" : "X"} / Semester ${topikItem.semester}",
        "alokasiWaktu": "${topikItem.alokasiWaktu}",
        "namaGuru": "${meta.namaGuru}",
        "semesterTahun": "Semester ${topikItem.semester === 1 ? "Ganjil" : "Genap"} - ${meta.tahunAjaran}",
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
          type: import_genai.Type.OBJECT,
          properties: {
            topik: { type: import_genai.Type.STRING },
            tpKode: { type: import_genai.Type.STRING },
            judul: { type: import_genai.Type.STRING },
            identitas: {
              type: import_genai.Type.OBJECT,
              properties: {
                sekolah: { type: import_genai.Type.STRING },
                mataPelajaran: { type: import_genai.Type.STRING },
                kelasSemester: { type: import_genai.Type.STRING },
                alokasiWaktu: { type: import_genai.Type.STRING },
                namaGuru: { type: import_genai.Type.STRING },
                semesterTahun: { type: import_genai.Type.STRING },
                topikMateri: { type: import_genai.Type.STRING },
                modelPembelajaran: { type: import_genai.Type.STRING }
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
              ]
            },
            tujuanPembelajaran: {
              type: import_genai.Type.ARRAY,
              items: { type: import_genai.Type.STRING }
            },
            petunjukPengerjaan: {
              type: import_genai.Type.ARRAY,
              items: { type: import_genai.Type.STRING }
            },
            ringkasanMateri: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  komponen: { type: import_genai.Type.STRING },
                  ringkasan: { type: import_genai.Type.STRING }
                },
                required: ["komponen", "ringkasan"]
              }
            },
            aktivitasList: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  judul: { type: import_genai.Type.STRING },
                  studiKasus: { type: import_genai.Type.STRING }
                },
                required: ["judul", "studiKasus"]
              }
            },
            rubrikPenilaian: { type: import_genai.Type.STRING }
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
          ]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating LKPD-single:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate LKPD." });
  }
});
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
          type: import_genai.Type.OBJECT,
          properties: {
            topik: { type: import_genai.Type.STRING },
            tpKode: { type: import_genai.Type.STRING },
            judul: { type: import_genai.Type.STRING },
            kisiKisiUmum: { type: import_genai.Type.STRING },
            identitas: {
              type: import_genai.Type.OBJECT,
              properties: {
                namaGuru: { type: import_genai.Type.STRING },
                sekolah: { type: import_genai.Type.STRING },
                mataPelajaran: { type: import_genai.Type.STRING },
                semesterTahun: { type: import_genai.Type.STRING },
                topikMateri: { type: import_genai.Type.STRING },
                modelPembelajaran: { type: import_genai.Type.STRING }
              },
              required: ["namaGuru", "sekolah", "mataPelajaran", "semesterTahun", "topikMateri", "modelPembelajaran"]
            },
            kisiKisi: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  no: { type: import_genai.Type.INTEGER },
                  tujuanPembelajaran: { type: import_genai.Type.STRING },
                  indikatorSoal: { type: import_genai.Type.STRING },
                  levelKognitif: { type: import_genai.Type.STRING },
                  bentukSoal: { type: import_genai.Type.STRING },
                  noSoal: { type: import_genai.Type.INTEGER }
                },
                required: ["no", "tujuanPembelajaran", "indikatorSoal", "levelKognitif", "bentukSoal", "noSoal"]
              }
            },
            daftarSoal: {
              type: import_genai.Type.ARRAY,
              items: {
                type: import_genai.Type.OBJECT,
                properties: {
                  no: { type: import_genai.Type.INTEGER },
                  tipe: { type: import_genai.Type.STRING },
                  levelKognitif: { type: import_genai.Type.STRING },
                  soal: { type: import_genai.Type.STRING },
                  pilihan: {
                    type: import_genai.Type.OBJECT,
                    properties: {
                      A: { type: import_genai.Type.STRING },
                      B: { type: import_genai.Type.STRING },
                      C: { type: import_genai.Type.STRING },
                      D: { type: import_genai.Type.STRING },
                      E: { type: import_genai.Type.STRING }
                    }
                  },
                  kunciJawaban: { type: import_genai.Type.STRING },
                  skorMaksimal: { type: import_genai.Type.INTEGER }
                },
                required: ["no", "tipe", "levelKognitif", "soal", "kunciJawaban", "skorMaksimal"]
              }
            }
          },
          required: ["topik", "tpKode", "judul", "kisiKisiUmum", "identitas", "kisiKisi", "daftarSoal"]
        }
      }
    });
    const data = JSON.parse(response.text || "{}");
    res.json(data);
  } catch (error) {
    console.error("Error generating Asesmen-single:", error);
    res.status(500).json({ error: error.message || "Gagal menggenerate Asesmen." });
  }
});
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
    const buffer = await import_docx2.Packer.toBuffer(doc);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (error) {
    console.error("Error exporting DOCX:", error);
    res.status(500).json({ error: error.message || "Gagal mengekspor dokumen DOCX." });
  }
});
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await (0, import_vite.createServer)({
      server: { middlewareMode: true },
      appType: "spa"
    });
    app.use(vite.middlewares);
  } else {
    const distPath = import_path.default.join(process.cwd(), "dist");
    app.use(import_express.default.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(import_path.default.join(distPath, "index.html"));
    });
  }
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}
startServer();
//# sourceMappingURL=server.cjs.map
