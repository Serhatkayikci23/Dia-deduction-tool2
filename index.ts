import { Client } from "pg";
import express from "express";
import ExcelJS from "exceljs";

const app = express();
app.use(express.json());

const baseUrl = "https://diademo.ws.dia.com.tr/api/v3";
const userName = "ws";
const password = "ws";
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

const fmt = new Intl.NumberFormat("tr-TR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

await login();
await db.connect();

await db.query(`DROP TABLE IF EXISTS personel`);
await db.query(`
    CREATE TABLE IF NOT EXISTS personel (
    id SERIAL PRIMARY KEY,
    tckimlikno TEXT,
    personeladisoyadi TEXT,
    persdepartmanaciklama TEXT,
    argefaaliyetgunsayisi TEXT,
    aylikbrutkazanc TEXT
    )
    `);

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
            "argefaaliyetgunsayisi",
            "aylikbrutkazanc",
          ],
        },
      },
    }),
  });

  const data = (await response.json()).result;

  for (const p of data) {
    await db.query(
      `INSERT INTO personel (tckimlikno, personeladisoyadi, persdepartmanaciklama, argefaaliyetgunsayisi,aylikbrutkazanc) VALUES ($1,$2,$3,$4,$5)`,
      [
        p.tckimlikno,
        p.personeladisoyadi,
        p.persdepartmanaciklama,
        p.argefaaliyetgunsayisi,
        p.aylikbrutkazanc,
      ],
    );
  }
  console.log(`${data.length} personel aktarıldı`);
} catch (error) {
  console.error("Personel çekilirken hata", error);
}

app.get("/api/personel", async (req, res) => {
  const result = await db.query(
    `SELECT tckimlikno AS "TC KİMLİK NO", personeladisoyadi AS "AD SOYAD", persdepartmanaciklama AS "Departman Adı", argefaaliyetgunsayisi AS "ARGE MERKEZİNDE ÇALIŞILAN GÜN SAYISI", aylikbrutkazanc AS "Aylık Brüt Kazanç" FROM personel ORDER BY personeladisoyadi`,
  );
  res.json(result.rows);
});

app.get("/api/personel/excel", async (req, res) => {
  const result = await db.query(
    `SELECT tckimlikno, personeladisoyadi, persdepartmanaciklama, argefaaliyetgunsayisi, aylikbrutkazanc FROM personel ORDER BY personeladisoyadi`,
  );

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Personel");

  const YELLOW = "FFFFFF00";
  const GREEN = "FF00B050";
  const BLUE = "FF00B0F0";
  const ORANGE = "FFFFC000";
  const NO_COLOR = "FFFFFFFF";
  const GRAY = "FF808080";
  const BROWN = "FF843C0C";
  const NAVY = "FF1F3864";

  sheet.columns = [
    { header: "S/NO", key: "sno", width: 6 },
    { header: "TC KİMLİK NO", key: "tckimlikno", width: 16 },
    { header: "AD", key: "ad", width: 22 },
    { header: "DEPARTMAN", key: "departman", width: 45 },

    {
      header: "ARGE MERKEZİNDE ÇALIŞILAN GÜN SAYISI",
      key: "argegun",
      width: 14,
    },
    {
      header: "DİĞER FAALİYETLERDE ÇALIŞMA GÜN SAYISI",
      key: "digergun",
      width: 14,
    },
    { header: "TOPLAM ÇALIŞMA GÜN SAYISI", key: "toplamgun", width: 12 },
    { header: "BRÜT TEMEL ÜCRET", key: "bruttemel", width: 16 },
    { header: "FAZLA MESAİ - EK ÜCRET", key: "fazlamesai", width: 16 },
    { header: "Aylık Üst Sınır", key: "aylikust", width: 14 },
    { header: "Günlük Üst Sınır", key: "gunlukust", width: 14 },
    { header: "AR-GE Aylık Üst Sınır", key: "argeaylikust", width: 14 },
    { header: "5510 Aylık Üst Sınır", key: "s5510aylik", width: 14 },
    { header: "TOPLAM BRÜT ÜCRET", key: "toplambrut", width: 16 },
    {
      header: "5510 SGK MATRAHI (DİĞER FAAL, İKRAMİYE PRİM, MESAİ)",
      key: "sgkmatrah",
      width: 20,
    },
    { header: "SGK İşçi Payı (0,14)", key: "sgkisci", width: 14 },
    { header: "SGK İşçi İşsizlik Payı (0,02)", key: "sgkissizlik", width: 14 },
    { header: "SGK İŞV PAYI (21,75)", key: "sgkisv", width: 14 },
    { header: "SGK İşveren İşsizlik Payı (0,02)", key: "sgkisvisz", width: 14 },
    {
      header:
        "DİĞER FAALİYETLER KAPSAMINDA HESAPLANAN SGK İNDİRİMİ %5 (EK MESAİ, PRİM, İKRAMİYE DAHİL)",
      key: "sgkindirim",
      width: 22,
    },
    {
      header: "ARGE MERKEZİNDEKİ ÜCRET (SİGORTA MATRAHI)",
      key: "argesigorta",
      width: 20,
    },
    { header: "SGK İŞV PAYI (21,75)", key: "sgkisvarge", width: 14 },
    {
      header: "SGK İşveren İşsizlik Payı (0,02)",
      key: "sgkisvisz2",
      width: 14,
    },
    {
      header: "ARGE MATRAHINDAN 5510 SGK İNDİRİMİ %5",
      key: "argesgk5",
      width: 16,
    },
    { header: "5746 SGK İNDİRİMİ %50", key: "sgk5746", width: 14 },
    { header: "ARGE MERKEZİNDEKİ ÜCRET", key: "argeucret", width: 20 },
    { header: "SGK İşçi Payı (0,14)", key: "sgkisci2", width: 14 },
    { header: "SGK İşçi İşsizlik Payı (0,01)", key: "sgkissizlik2", width: 14 },
    { header: "AR-GE GELİR VERGİSİ MATRAHI", key: "argegvmat", width: 18 },
    { header: "GELİR VERGİSİ ORANI", key: "gvorani", width: 14 },
    { header: "GELİR VERGİSİ TUTARI", key: "gvtutari", width: 14 },
    { header: "AGİ", key: "agi", width: 10 },
    {
      header: "AGİ MAHSUBU SONRASI GELİR VERGİSİ TUTARI",
      key: "agimahsup",
      width: 18,
    },
    { header: "ARGE İSTİSNA ORANI", key: "argeorani", width: 14 },
    {
      header: "TERKİN EDİLECEK GELİR VERGİSİ TUTARI",
      key: "terkingv",
      width: 18,
    },
    { header: "ÖDENECEK GV STOPAJ", key: "odenecekgv", width: 16 },
    {
      header:
        "5746 SAYILI KANUN KAPSAMINDA DAMGA TERKİN EDİLECEK DAMGA VERGİSİ",
      key: "damgaterkin",
      width: 22,
    },
    {
      header: "TOPLAM TEŞVİK TUTARI (SGK,GV,DV)",
      key: "toplamtesvik",
      width: 18,
    },
    { header: "ARGE İŞV MALİYETİ", key: "argemaliyet", width: 18 },
  ];

  // ── ADIM 2: Sütun başlık satırını formatla (şu an row 1'de) ──────────────
  const colColors: Record<string, string> = {
    sno: GRAY,
    tckimlikno: GRAY,
    ad: GRAY,
    departman: GRAY,
    argegun: GRAY,
    digergun: GRAY,
    toplamgun: GRAY,
    bruttemel: GRAY,
    fazlamesai: GRAY,
    aylikust: YELLOW,
    gunlukust: YELLOW,
    argeaylikust: YELLOW,
    s5510aylik: YELLOW,
    toplambrut: GRAY,
    sgkmatrah: GREEN,
    sgkisci: GREEN,
    sgkissizlik: GREEN,
    sgkisv: GREEN,
    sgkisvisz: GREEN,
    sgkindirim: GREEN,
    argesigorta: BROWN,
    sgkisvarge: BROWN,
    sgkisvisz2: BROWN,
    argesgk5: BROWN,
    sgk5746: BROWN,
    argeucret: GRAY,
    sgkisci2: GRAY,
    sgkissizlik2: GRAY,
    argegvmat: GRAY,
    gvorani: GRAY,
    gvtutari: GRAY,
    agi: GRAY,
    agimahsup: GRAY,
    argeorani: GRAY,
    terkingv: GRAY,
    odenecekgv: GRAY,
    damgaterkin: NAVY,
    toplamtesvik: GRAY,
    argemaliyet: GRAY,
  };

  sheet.getRow(1).height = 60;
  sheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
    if (colNumber > sheet.columns.length) return;
    const key = sheet.columns[colNumber - 1].key as string;
    const bg = colColors[key] ?? NO_COLOR;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    cell.font = {
      bold: true,
      color: { argb: "FF000000" },
      name: "Arial",
      size: 9,
    };
    cell.alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
  });

  // ── ADIM 3: Grup başlık satırını en üste ekle (row 1'i aşağı iter → row 2) ─
  sheet.insertRow(1, []);
  const groupRow = sheet.getRow(1);
  groupRow.height = 30;

  const groups = [
    {
      label: "SGK Aylık Ve Günlük Üst Sınır İşlemleri",
      startCol: 10,
      endCol: 13,
      color: GRAY,
    },
    {
      label: "5510 SAYILI KANUN KAPSAMINDA MATRAH VE %5'LİK İNDİRİM",
      startCol: 15,
      endCol: 20,
      color: GREEN,
    },
    {
      label: "5746 SAYILI KANUN KAPSAMINDA SGK İŞVEREN PAYI HESAPLAMA",
      startCol: 21,
      endCol: 25,
      color: BROWN,
    },
    {
      label: "5746 SAYILI KANUN KAPSAMINDA GELİR VERGİSİ STOPAJ HESAPLAMA",
      startCol: 26,
      endCol: 36,
      color: GRAY,
    },
  ];

  groups.forEach(({ label, startCol, endCol, color }) => {
    const cell = groupRow.getCell(startCol);
    cell.value = label;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };

    cell.font = {
      bold: true,
      color: { argb: "FF000000" },
      name: "Arial",
      size: 9,
    };
    cell.alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    sheet.mergeCells(1, startCol, 1, endCol);
  });

  const thickBorder = {
    top: { style: "medium" as const },
    bottom: { style: "medium" as const },
    left: { style: "medium" as const },
    right: { style: "medium" as const },
  };
  const thinBorder = {
    top: { style: "thin" as const },
    bottom: { style: "thin" as const },
    left: { style: "thin" as const },
    right: { style: "thin" as const },
  };

  const cellBorder = {
    top: { style: "thin" as const },
    bottom: { style: "medium" as const },
    left: { style: "thin" as const },
    right: { style: "medium" as const },
  };

  [1, 2].forEach((rowNum) => {
    for (let col = 1; col <= sheet.columns.length; col++) {
      sheet.getRow(rowNum).getCell(col).border = thickBorder;
    }
  });

  let rowNum = 3;

  result.rows.forEach((p, index) => {
    const row = sheet.addRow({
      sno: index + 1,
      tckimlikno: p.tckimlikno,
      ad: p.personeladisoyadi,
      departman: p.persdepartmanaciklama,
    });

    row.getCell("gunlukust").value = { formula: `J${rowNum}/30` };
    row.getCell("argeaylikust").value = { formula: `E${rowNum}*K${rowNum}` };
    row.getCell("s5510aylik").value = { formula: `F${rowNum}*K${rowNum}` };
    row.getCell("toplambrut").value = { formula: `H${rowNum}+I${rowNum}` };
    row.getCell("sgkmatrah").value = {
      formula: `IF(G${rowNum}=0,0,(H${rowNum}/G${rowNum}*F${rowNum})+I${rowNum})`,
    };
    row.getCell("sgkisci").value = { formula: `O${rowNum}*0.14` };
    row.getCell("sgkissizlik").value = { formula: `O${rowNum}*0.02` };
    row.getCell("sgkisv").value = { formula: `O${rowNum}*0.2175` };
    row.getCell("sgkisvisz").value = { formula: `O${rowNum}*0.02` };
    row.getCell("sgkisvisz2").value = { formula: `O${rowNum}*0.02` };
    row.getCell("sgkindirim").value = { formula: `O${rowNum}*0.05` };
    row.getCell("sgkisvarge").value = { formula: `U${rowNum}*0.2175` };
    row.getCell("argesigorta").value = {
      formula: `IF(G${rowNum}=0,0,E${rowNum}/G${rowNum}*H${rowNum})`,
    };
    row.getCell("argesgk5").value = { formula: `U${rowNum}*0.05` };
    row.getCell("sgk5746").value = { formula: `U${rowNum}*0.1675/2` };
    row.getCell("argeucret").value = {
      formula: `IF(G${rowNum}=0,0,E${rowNum}/G${rowNum}*H${rowNum})`,
    };
    row.getCell("sgkisci2").value = { formula: `Z${rowNum}*0.14` };
    row.getCell("sgkissizlik2").value = { formula: `Z${rowNum}*0.01` };
    row.getCell("argegvmat").value = {
      formula: `Z${rowNum}-AA${rowNum}-AB${rowNum}`,
    };
    row.getCell("agimahsup").value = { formula: `AE${rowNum}-AF${rowNum}` };
    row.getCell("terkingv").value = { formula: `AG${rowNum}*AH${rowNum}` };
    row.getCell("odenecekgv").value = { formula: `AG${rowNum}-AI${rowNum}` };
    row.getCell("toplamtesvik").value = {
      formula: `Y${rowNum}+AI${rowNum}+AK${rowNum}`,
    };

    for (let col = 1; col <= sheet.columns.length; col++) {
      const cell = row.getCell(col);

      cell.border = thinBorder;

      if (index % 2 === 0) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF2F2F2" },
        };
      }
    }

    rowNum++;
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader("Content-Disposition", "attachment; filename=personel.xlsx");

  await workbook.xlsx.write(res);
  res.end();
});

app.listen(4000, () => console.log("Sunucu: https:localhost:4000"));
