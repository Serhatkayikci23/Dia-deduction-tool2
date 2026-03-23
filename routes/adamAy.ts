import { Client } from "pg";
import express from "express";
import ExcelJS from "exceljs";

// proje.ts ilk hali

const app = express();
app.use(express.json());

const baseUrl = "https://diademo.ws.dia.com.tr/api/v3";
const firmCode = 1;
const periodCode = 1;

const db = new Client({
  host: "127.0.0.1",
  port: 5432,
  database: "dia-deduction-tool2",
  user: "postgres",
  password: "123",
});

let sessionId: string | null = null;

async function login() {
  const response = await fetch(baseUrl + "/sis/json", {
    method: "POST",
    body: JSON.stringify({
      login: {
        username: "ws",
        password: "ws",
        disconnect_same_user: "true",
        lang: "tr",
        params: { apikey: "" },
      },
    }),
  });
  const data = (await response.json()) as { code: string; msg: string };
  sessionId = data.msg;
}

await login();
await db.connect();

// Personel tablosu — DIA'dan gelen alanlar
await db.query(`DROP TABLE IF EXISTS proje_personel`);
await db.query(`
  CREATE TABLE IF NOT EXISTS proje_personel (
    id SERIAL PRIMARY KEY,
    tckimlikno TEXT,
    personeladisoyadi TEXT,
    persdepartmanaciklama TEXT,
    gorevi TEXT
  )
`);

// Proje tablosu — elle girilen proje bilgileri
await db.query(`
  CREATE TABLE IF NOT EXISTS projeler (
    id SERIAL PRIMARY KEY,
    proje_adi TEXT,
    proje_kodu TEXT,
    baslangic_tarihi TEXT,
    bitis_tarihi TEXT,
    renk TEXT
  )
`);

// Personel-proje oranları — elle girilen
await db.query(`
  CREATE TABLE IF NOT EXISTS personel_proje_oran (
    id SERIAL PRIMARY KEY,
    tckimlikno TEXT,
    proje_kodu TEXT,
    oran NUMERIC
  )
`);

// DIA'dan personel çek
try {
  const response = await fetch(baseUrl + "/per/json", {
    method: "POST",
    body: JSON.stringify({
      per_personel_puantaj_listele: {
        session_id: sessionId,
        firma_kodu: firmCode,
        donem_kodu: periodCode,
        params: {
          selectedcolumns: [
            "tckimlikno",
            "personeladisoyadi",
            "persdepartmanaciklama",
            "gorevi",
          ],
        },
      },
    }),
  });

  const data = (await response.json()).result;

  for (const p of data) {
    await db.query(
      `INSERT INTO proje_personel (tckimlikno, personeladisoyadi, persdepartmanaciklama, gorevi)
       VALUES ($1, $2, $3, $4)
       ON CONFLICT DO NOTHING`,
      [p.tckimlikno, p.personeladisoyadi, p.persdepartmanaciklama, p.gorevi],
    );
  }
  console.log(`${data.length} personel DIA'dan aktarıldı`);
} catch (error) {
  console.error("Personel çekilirken hata:", error);
}

// ─── PROJE CRUD ───────────────────────────────────────────────

// Tüm projeleri listele
app.get("/api/projeler", async (req, res) => {
  const result = await db.query(`SELECT * FROM projeler ORDER BY id`);
  res.json(result.rows);
});

// Yeni proje ekle
// Body: { proje_adi, proje_kodu, baslangic_tarihi, bitis_tarihi, renk }
app.post("/api/projeler", async (req, res) => {
  const { proje_adi, proje_kodu, baslangic_tarihi, bitis_tarihi, renk } =
    req.body;
  const result = await db.query(
    `INSERT INTO projeler (proje_adi, proje_kodu, baslangic_tarihi, bitis_tarihi, renk)
     VALUES ($1, $2, $3, $4, $5) RETURNING *`,
    [proje_adi, proje_kodu, baslangic_tarihi, bitis_tarihi, renk],
  );
  res.json(result.rows[0]);
});

// Proje güncelle
app.put("/api/projeler/:id", async (req, res) => {
  const { proje_adi, proje_kodu, baslangic_tarihi, bitis_tarihi, renk } =
    req.body;
  const result = await db.query(
    `UPDATE projeler SET proje_adi=$1, proje_kodu=$2, baslangic_tarihi=$3, bitis_tarihi=$4, renk=$5
     WHERE id=$6 RETURNING *`,
    [
      proje_adi,
      proje_kodu,
      baslangic_tarihi,
      bitis_tarihi,
      renk,
      req.params.id,
    ],
  );
  res.json(result.rows[0]);
});

// Proje sil
app.delete("/api/projeler/:id", async (req, res) => {
  await db.query(`DELETE FROM projeler WHERE id=$1`, [req.params.id]);
  res.json({ ok: true });
});

// ─── PERSONEL-PROJE ORAN CRUD ─────────────────────────────────

// Tüm oranları listele
app.get("/api/oranlar", async (req, res) => {
  const result = await db.query(
    `SELECT * FROM personel_proje_oran ORDER BY tckimlikno`,
  );
  res.json(result.rows);
});

// Oran ekle/güncelle
// Body: { tckimlikno, proje_kodu, oran }
app.post("/api/oranlar", async (req, res) => {
  const { tckimlikno, proje_kodu, oran } = req.body;
  const result = await db.query(
    `INSERT INTO personel_proje_oran (tckimlikno, proje_kodu, oran)
     VALUES ($1, $2, $3)
     ON CONFLICT (tckimlikno, proje_kodu) DO UPDATE SET oran = EXCLUDED.oran
     RETURNING *`,
    [tckimlikno, proje_kodu, oran],
  );
  res.json(result.rows[0]);
});

// Oran sil
app.delete("/api/oranlar/:id", async (req, res) => {
  await db.query(`DELETE FROM personel_proje_oran WHERE id=$1`, [
    req.params.id,
  ]);
  res.json({ ok: true });
});

// ─── PERSONEL LİSTESİ ─────────────────────────────────────────

app.get("/api/personel", async (req, res) => {
  const result = await db.query(
    `SELECT tckimlikno AS "TC KİMLİK NO",
            personeladisoyadi AS "AD SOYAD",
            persdepartmanaciklama AS "DEPARTMAN",
            gorevi AS "GÖREVİ"
     FROM proje_personel ORDER BY personeladisoyadi`,
  );
  res.json(result.rows);
});

// ─── EXCEL EXPORT ─────────────────────────────────────────────

app.get("/api/personel/excel", async (req, res) => {
  // Verileri çek
  const personelRes = await db.query(
    `SELECT tckimlikno, personeladisoyadi, persdepartmanaciklama, gorevi
     FROM proje_personel ORDER BY personeladisoyadi`,
  );
  const projelerRes = await db.query(`SELECT * FROM projeler ORDER BY id`);
  const oranlarRes = await db.query(`SELECT * FROM personel_proje_oran`);

  const personelList = personelRes.rows;
  const projeler = projelerRes.rows;

  // Oranları map'e al: { tckimlikno: { proje_kodu: oran } }
  const oranMap: Record<string, Record<string, number>> = {};
  for (const o of oranlarRes.rows) {
    if (!oranMap[o.tckimlikno]) oranMap[o.tckimlikno] = {};
    oranMap[o.tckimlikno][o.proje_kodu] = parseFloat(o.oran);
  }

  // Renk sabitleri
  const COLORS: Record<string, string> = {
    MOR: "FF7030A0",
    TURUNCU: "FFFF6600",
    MAVI: "FF00B0F0",
    SARI: "FFFFFF00",
    GRAY: "FFD9D9D9",
    WHITE: "FFFFFFFF",
    LIGHT_YELLOW: "FFFFFFCC",
    GREEN_NEW: "FF92D050",
    RED_REMOVED: "FFFF0000",
  };

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Personel Dağılım");

  // Yardımcı: hücre stili
  function styleCell(
    cell: ExcelJS.Cell,
    opts: {
      bold?: boolean;
      fontSize?: number;
      bgColor?: string;
      alignH?: ExcelJS.Alignment["horizontal"];
      wrap?: boolean;
      fontColor?: string;
      borderTop?: boolean;
    } = {},
  ) {
    const {
      bold = false,
      fontSize = 9,
      bgColor,
      alignH = "center",
      wrap = true,
      fontColor = "FF000000",
      borderTop = true,
    } = opts;

    cell.font = {
      bold,
      size: fontSize,
      name: "Arial",
      color: { argb: fontColor },
    };
    cell.alignment = { horizontal: alignH, vertical: "middle", wrapText: wrap };

    if (bgColor) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: bgColor },
      };
    }

    const thin = { style: "thin" as const };
    const none = { style: undefined };
    cell.border = {
      left: thin,
      right: thin,
      top: borderTop ? thin : none,
      bottom: thin,
    };
  }

  // Sütun genişlikleri: A=TC, B=Ad, C=Departman, D=Bölüm, E...(projeler), son=TOPLAM
  sheet.getColumn(1).width = 15; // TC
  sheet.getColumn(2).width = 22; // Ad Soyad
  sheet.getColumn(3).width = 22; // Departman
  sheet.getColumn(4).width = 20; // Bölüm
  const projStartCol = 5;
  projeler.forEach((_, i) => {
    sheet.getColumn(projStartCol + i).width = 14;
  });
  const totalCol = projStartCol + projeler.length;
  sheet.getColumn(totalCol).width = 10;

  const getColor = (renk: string) => COLORS[renk?.toUpperCase()] ?? "FFD9D9D9";

  // ── Satır 1: Proje adları ──
  sheet.getRow(1).height = 45;
  for (let i = 0; i < projeler.length; i++) {
    const col = projStartCol + i;
    const cell = sheet.getRow(1).getCell(col);
    cell.value = projeler[i].proje_adi;
    styleCell(cell, {
      bold: true,
      bgColor: getColor(projeler[i].renk),
      borderTop: true,
    });
  }

  // ── Satır 2: Proje kodları ──
  sheet.getRow(2).height = 20;
  for (let i = 0; i < projeler.length; i++) {
    const col = projStartCol + i;
    const cell = sheet.getRow(2).getCell(col);
    cell.value = `Proje Kodu:\n${projeler[i].proje_kodu}`;
    styleCell(cell, {
      bold: true,
      bgColor: getColor(projeler[i].renk),
      borderTop: true,
    });
  }

  // ── Satır 3: Tarihler ──
  sheet.getRow(3).height = 60;
  for (let i = 0; i < projeler.length; i++) {
    const col = projStartCol + i;
    const cell = sheet.getRow(3).getCell(col);
    cell.value = `Başlangıç: ${projeler[i].baslangic_tarihi}\nBitiş Tarihi: ${projeler[i].bitis_tarihi}`;
    styleCell(cell, {
      bgColor: getColor(projeler[i].renk),
      fontSize: 8,
      borderTop: true,
    });
  }

  // ── Satır 4: Başlıklar ──
  sheet.getRow(4).height = 25;
  const headers = [
    "TC KİMLİK NO",
    "PERSONEL ADI SOYADI",
    "DEPARTMAN",
    "GÖREVİ",
  ];
  headers.forEach((h, i) => {
    const cell = sheet.getRow(4).getCell(i + 1);
    cell.value = h;
    styleCell(cell, { bold: true, bgColor: COLORS.GRAY, borderTop: true });
  });
  for (let i = 0; i < projeler.length; i++) {
    const cell = sheet.getRow(4).getCell(projStartCol + i);
    styleCell(cell, { bgColor: getColor(projeler[i].renk), borderTop: true });
  }
  const totalHeaderCell = sheet.getRow(4).getCell(totalCol);
  totalHeaderCell.value = "TOPLAM";
  styleCell(totalHeaderCell, {
    bold: true,
    bgColor: COLORS.GRAY,
    borderTop: true,
  });

  // ── Veri satırları ──
  personelList.forEach((p, idx) => {
    const rowNum = 5 + idx;
    sheet.getRow(rowNum).height = 18;
    const bg = idx % 2 === 0 ? COLORS.WHITE : COLORS.LIGHT_YELLOW;

    const rowData = [
      p.tckimlikno,
      p.personeladisoyadi,
      p.persdepartmanaciklama,
      p.gorevi ?? "",
    ];
    rowData.forEach((val, ci) => {
      const cell = sheet.getRow(rowNum).getCell(ci + 1);
      cell.value = val;
      styleCell(cell, { bgColor: bg, alignH: "left", borderTop: false });
    });

    const oranlar = oranMap[p.tckimlikno] ?? {};
    let toplamFormulaParts: string[] = [];

    projeler.forEach((proj, pi) => {
      const col = projStartCol + pi;
      const colLetter = sheet.getColumn(col).letter;
      const cell = sheet.getRow(rowNum).getCell(col);
      const oran = oranlar[proj.proje_kodu];
      cell.value = oran ?? null;
      if (oran !== undefined) cell.numFmt = "0.00";
      styleCell(cell, { bgColor: bg, borderTop: false });
      toplamFormulaParts.push(`${colLetter}${rowNum}`);
    });

    const totalCell = sheet.getRow(rowNum).getCell(totalCol);
    totalCell.value = {
      formula: `IFERROR(SUM(${toplamFormulaParts.join(",")}),0)`,
    };
    styleCell(totalCell, { bgColor: bg, borderTop: false });
  });

  // ── Özet bölümü ──
  const lastDataRow = 4 + personelList.length;
  const sumStart = lastDataRow + 5;

  const summaryLabels = [
    "Personel gideri çarpanı",
    "PROJE ADAM/AY",
    "PROJE SÜRESİ (AY)",
    "PROJE GÖREVLİ PERSONEL SAYISI",
    "Proje Toplam Adam/ay",
  ];

  summaryLabels.forEach((label, i) => {
    const rowNum = sumStart + i;
    sheet.getRow(rowNum).height = 18;
    const cell = sheet.getRow(rowNum).getCell(4);
    cell.value = label;
    styleCell(cell, {
      bold: i === 1,
      bgColor: i === 1 ? COLORS.LIGHT_YELLOW : COLORS.WHITE,
      alignH: "left",
      borderTop: false,
    });

    // PROJE ADAM/AY satırı için otomatik formül
    if (i === 1) {
      projeler.forEach((_, pi) => {
        const col = projStartCol + pi;
        const colLetter = sheet.getColumn(col).letter;
        const cell = sheet.getRow(rowNum).getCell(col);
        cell.value = {
          formula: `SUMPRODUCT((${colLetter}5:${colLetter}${lastDataRow})*(${colLetter}5:${colLetter}${lastDataRow}<>""))`,
        };
        styleCell(cell, {
          bold: true,
          bgColor: COLORS.LIGHT_YELLOW,
          borderTop: false,
        });
      });
    }
  });

  // Yeni eklenen / Ayrılan göstergesi
  const legendRow1 = sheet.getRow(sumStart + 1);
  const legendRow2 = sheet.getRow(sumStart + 2);
  legendRow1.getCell(1).value = "Yeni eklenen";
  styleCell(legendRow1.getCell(1), {
    bgColor: COLORS.GREEN_NEW,
    alignH: "left",
    borderTop: false,
  });
  legendRow2.getCell(1).value = "Ayrılan";
  styleCell(legendRow2.getCell(1), {
    bgColor: COLORS.RED_REMOVED,
    alignH: "left",
    fontColor: "FFFFFFFF",
    borderTop: false,
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=proje_personel.xlsx",
  );
  await workbook.xlsx.write(res);
  res.end();
});

app.listen(4000, () => console.log("Sunucu: http://localhost:4000"));
