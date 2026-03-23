import { Router } from "express";
import ExcelJS from "exceljs";
import { db } from "../Db.ts";

export const bordroRouter = Router();

export async function setupBordro(data: any[]) {
  await db.query(`DROP TABLE IF EXISTS bordro_personel`);
  await db.query(`
    CREATE TABLE bordro_personel (
      id SERIAL PRIMARY KEY,
      tckimlikno TEXT UNIQUE,
      personeladisoyadi TEXT,
      persdepartmanaciklama TEXT,
      gorevi TEXT
    )
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS bordro_detay (
      id SERIAL PRIMARY KEY,
      tckimlikno TEXT UNIQUE,
      argegun NUMERIC DEFAULT 0,
      digergun NUMERIC DEFAULT 0,
      toplamgun NUMERIC DEFAULT 0,
      bruttemel NUMERIC DEFAULT 0,
      fazlamesai NUMERIC DEFAULT 0,
      aylikust NUMERIC DEFAULT 0,
      argeorani NUMERIC DEFAULT 0,
      agi NUMERIC DEFAULT 0,
      gvorani NUMERIC DEFAULT 0,
      damgaterkin NUMERIC DEFAULT 0
    )
  `);
  await db.query(`
  ALTER TABLE bordro_detay 
  ADD COLUMN IF NOT EXISTS damgaterkin NUMERIC DEFAULT 0
`);

  for (const p of data) {
    await db.query(
      `INSERT INTO bordro_personel (tckimlikno, personeladisoyadi, persdepartmanaciklama, gorevi)
       VALUES ($1,$2,$3,$4) ON CONFLICT (tckimlikno) DO UPDATE SET
         personeladisoyadi=EXCLUDED.personeladisoyadi,
         persdepartmanaciklama=EXCLUDED.persdepartmanaciklama,
         gorevi=EXCLUDED.gorevi`,
      [
        p.tckimlikno,
        p.personeladisoyadi,
        p.persdepartmanaciklama,
        p.gorevi ?? "",
      ],
    );
  }
}

bordroRouter.get("/personel", async (req, res) => {
  const result = await db.query(`
    SELECT b.*, d.argegun, d.digergun, d.toplamgun, d.bruttemel,
           d.fazlamesai, d.aylikust, d.argeorani, d.agi, d.gvorani
    FROM bordro_personel b
    LEFT JOIN bordro_detay d ON b.tckimlikno = d.tckimlikno
    ORDER BY b.personeladisoyadi
  `);
  res.json(result.rows);
});

bordroRouter.put("/personel/:tckimlikno", async (req, res) => {
  const { tckimlikno } = req.params;
  const {
    argegun,
    digergun,
    toplamgun,
    bruttemel,
    fazlamesai,
    aylikust,
    argeorani,
    agi,
    gvorani,
    damgaterkin,
  } = req.body;
  await db.query(
    `INSERT INTO bordro_detay (tckimlikno, argegun, digergun, toplamgun, bruttemel, fazlamesai, aylikust, argeorani, agi, gvorani, damgaterkin)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
     ON CONFLICT (tckimlikno) DO UPDATE SET
       argegun=$2, digergun=$3, toplamgun=$4, bruttemel=$5, fazlamesai=$6,
       aylikust=$7, argeorani=$8, agi=$9, gvorani=$10, damgaterkin=$11`,
    [
      tckimlikno,
      argegun,
      digergun,
      toplamgun,
      bruttemel,
      fazlamesai,
      aylikust,
      argeorani,
      agi,
      gvorani,
      damgaterkin,
    ],
  );
  res.json({ ok: true });
});

bordroRouter.get("/excel", async (req, res) => {
  const result = await db.query(`
    SELECT b.*, d.argegun, d.digergun, d.toplamgun, d.bruttemel,
           d.fazlamesai, d.aylikust, d.argeorani, d.agi, d.gvorani, d.damgaterkin
    FROM bordro_personel b
    LEFT JOIN bordro_detay d ON b.tckimlikno = d.tckimlikno
    ORDER BY b.personeladisoyadi
  `);

  const YELLOW = "FFFFFF00";
  const GREEN = "FF00B050";
  const BROWN = "FF843C0C";
  const NAVY = "FF1F3864";
  const GRAY = "FF808080";
  const headerLabels: Record<string, string> = {
    sno: "S/NO",
    tckimlikno: "TC KİMLİK NO",
    ad: "AD SOYAD",
    departman: "DEPARTMAN",
    gorevi: "GÖREVİ",
    argegun: "ARGE MERKEZİNDE ÇALIŞILAN GÜN SAYISI",
    digergun: "DİĞER FAALİYETLERDE ÇALIŞMA GÜN SAYISI",
    toplamgun: "TOPLAM ÇALIŞMA GÜN SAYISI",
    bruttemel: "BRÜT TEMEL ÜCRET",
    fazlamesai: "FAZLA MESAİ - EK ÜCRET",
    aylikust: "Aylık Üst Sınır",
    gunlukust: "Günlük Üst Sınır",
    argeaylikust: "AR-GE Aylık Üst Sınır",
    s5510aylik: "5510 Aylık Üst Sınır",
    toplambrut: "TOPLAM BRÜT ÜCRET",
    sgkmatrah: "5510 SGK MATRAHI",
    sgkisci: "SGK İşçi Payı (0,14)",
    sgkissizlik: "SGK İşçi İşsizlik Payı (0,01)",
    sgkisv: "SGK İŞV PAYI (21,75)",
    sgkisvisz: "SGK İşveren İşsizlik Payı (0,02)",
    sgkindirim: "SGK İNDİRİMİ %5",
    argesigorta: "ARGE MERKEZİNDEKİ ÜCRET (SİGORTA MATRAHI)",
    sgkisvarge: "SGK İŞV PAYI (21,75)",
    sgkisvisz2: "SGK İşveren İşsizlik Payı (0,02)",
    argesgk5: "ARGE MATRAHINDAN 5510 SGK İNDİRİMİ %5",
    sgk5746: "5746 SGK İNDİRİMİ %50",
    argeucret: "ARGE MERKEZİNDEKİ ÜCRET",
    sgkisci2: "SGK İşçi Payı (0,14)",
    sgkissizlik2: "SGK İşçi İşsizlik Payı (0,01)",
    argegvmat: "AR-GE GELİR VERGİSİ MATRAHI",
    gvorani: "GELİR VERGİSİ ORANI",
    gvtutari: "GELİR VERGİSİ TUTARI",
    agi: "AGİ",
    agimahsup: "AGİ MAHSUBU SONRASI GELİR VERGİSİ TUTARI",
    argeorani: "ARGE İSTİSNA ORANI",
    terkingv: "TERKİN EDİLECEK GELİR VERGİSİ TUTARI",
    odenecekgv: "ÖDENECEK GV STOPAJ",
    damgaterkin:
      "5746 SAYILI KANUN KAPSAMINDA DAMGA TERKİN EDİLECEK DAMGA VERGİSİ",
    toplamtesvik: "TOPLAM TEŞVİK TUTARI (SGK,GV,DV)",
    argemaliyet: "ARGE İŞV MALİYETİ",
  };

  const colDefs = [
    { key: "sno", width: 6 },
    { key: "tckimlikno", width: 16 },
    { key: "ad", width: 22 },
    { key: "departman", width: 30 },
    { key: "gorevi", width: 20 },
    { key: "argegun", width: 14 },
    { key: "digergun", width: 14 },
    { key: "toplamgun", width: 12 },
    { key: "bruttemel", width: 16 },
    { key: "fazlamesai", width: 16 },
    { key: "aylikust", width: 14 },
    { key: "gunlukust", width: 14 },
    { key: "argeaylikust", width: 14 },
    { key: "s5510aylik", width: 14 },
    { key: "toplambrut", width: 16 },
    { key: "sgkmatrah", width: 20 },
    { key: "sgkisci", width: 14 },
    { key: "sgkissizlik", width: 14 },
    { key: "sgkisv", width: 14 },
    { key: "sgkisvisz", width: 14 },
    { key: "sgkindirim", width: 22 },
    { key: "argesigorta", width: 20 },
    { key: "sgkisvarge", width: 14 },
    { key: "sgkisvisz2", width: 14 },
    { key: "argesgk5", width: 16 },
    { key: "sgk5746", width: 14 },
    { key: "argeucret", width: 20 },
    { key: "sgkisci2", width: 14 },
    { key: "sgkissizlik2", width: 14 },
    { key: "argegvmat", width: 18 },
    { key: "gvorani", width: 14 },
    { key: "gvtutari", width: 14 },
    { key: "agi", width: 10 },
    { key: "agimahsup", width: 18 },
    { key: "argeorani", width: 14 },
    { key: "terkingv", width: 18 },
    { key: "odenecekgv", width: 16 },
    { key: "damgaterkin", width: 22 },
    { key: "toplamtesvik", width: 18 },
    { key: "argemaliyet", width: 18 },
  ];

  const colColors: Record<string, string> = {
    sno: GRAY,
    tckimlikno: GRAY,
    ad: GRAY,
    departman: GRAY,
    gorevi: GRAY,
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

  const thinBorder = {
    top: { style: "thin" as const },
    bottom: { style: "thin" as const },
    left: { style: "thin" as const },
    right: { style: "thin" as const },
  };

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Aylık Bordro");
  sheet.columns = colDefs.map((c) => ({ key: c.key, width: c.width }));

  sheet.getRow(1).height = 30;
  const groups = [
    {
      label: "SGK Aylık Ve Günlük Üst Sınır İşlemleri",
      startCol: 11,
      endCol: 14,
      color: YELLOW,
    },
    {
      label: "5510 SAYILI KANUN KAPSAMINDA MATRAH VE %5'LİK İNDİRİM",
      startCol: 16,
      endCol: 21,
      color: GREEN,
    },
    {
      label: "5746 SAYILI KANUN KAPSAMINDA SGK İŞVEREN PAYI HESAPLAMA",
      startCol: 22,
      endCol: 26,
      color: BROWN,
    },
    {
      label: "5746 SAYILI KANUN KAPSAMINDA GELİR VERGİSİ STOPAJ HESAPLAMA",
      startCol: 27,
      endCol: 37,
      color: GRAY,
    },
  ];
  groups.forEach(({ label, startCol, endCol, color }) => {
    const cell = sheet.getRow(1).getCell(startCol);
    cell.value = label;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
    cell.font = {
      bold: true,
      color: { argb: color === YELLOW ? "FFFF0000" : "FFFFFFFF" },
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

  sheet.getRow(2).height = 60;
  colDefs.forEach((col, i) => {
    const cell = sheet.getRow(2).getCell(i + 1);
    cell.value = headerLabels[col.key] ?? col.key;

    const bg = colColors[col.key] ?? "FFFFFFFF";
    cell.value = headerLabels[col.key] ?? col.key;

    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    cell.font = {
      bold: true,
      color: { argb: bg === YELLOW ? "FFFF0000" : "FFFFFFFF" },
      name: "Arial",
      size: 9,
    };
    cell.alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    cell.border = thinBorder;
  });

  result.rows.forEach((p, index) => {
    const R = 3 + index;
    const row = sheet.getRow(R);
    row.height = 18;
    row.getCell("sno").value = index + 1;
    row.getCell("tckimlikno").value = p.tckimlikno;
    row.getCell("ad").value = p.personeladisoyadi;
    row.getCell("departman").value = p.persdepartmanaciklama || "";
    row.getCell("gorevi").value = p.gorevi || "";
    row.getCell("argegun").value = p.argegun || 0;
    row.getCell("digergun").value = p.digergun || 0;
    row.getCell("toplamgun").value = p.toplamgun || 0;
    row.getCell("bruttemel").value = p.bruttemel || 0;
    row.getCell("fazlamesai").value = p.fazlamesai || 0;
    row.getCell("aylikust").value = p.aylikust || 0;
    row.getCell("gvorani").value = p.gvorani || 0;
    row.getCell("agi").value = p.agi || 0;
    row.getCell("argeorani").value = p.argeorani || 0;
    row.getCell("gunlukust").value = { formula: `K${R}/30` };
    row.getCell("argeaylikust").value = { formula: `F${R}*L${R}` };
    row.getCell("s5510aylik").value = { formula: `G${R}*L${R}` };
    row.getCell("toplambrut").value = { formula: `I${R}+J${R}` };
    row.getCell("sgkmatrah").value = {
      formula: `IFERROR(IF(H${R}=0,0,(I${R}/H${R}*G${R})+J${R}),0)`,
    };

    row.getCell("sgkisci").value = { formula: `P${R}*0.14` };
    row.getCell("sgkissizlik").value = { formula: `P${R}*0.01` };
    row.getCell("sgkisv").value = { formula: `P${R}*0.2175` };
    row.getCell("sgkisvisz").value = { formula: `P${R}*0.02` };
    row.getCell("sgkindirim").value = { formula: `P${R}*0.05` };
    row.getCell("argesigorta").value = {
      formula: `IFERROR(IF(H${R}=0,0,F${R}/H${R}*I${R}),0)`,
    };

    row.getCell("sgkisvarge").value = { formula: `V${R}*0.2175` };
    row.getCell("sgkisvisz2").value = { formula: `V${R}*0.02` };
    row.getCell("argesgk5").value = { formula: `V${R}*0.05` };
    row.getCell("sgk5746").value = { formula: `V${R}*0.1675/2` };

    row.getCell("argeucret").value = {
      formula: `IFERROR(IF(H${R}=0,0,F${R}/H${R}*I${R}),0)`,
    };
    row.getCell("sgkisci2").value = { formula: `AA${R}*0.14` };
    row.getCell("sgkissizlik2").value = { formula: `AA${R}*0.01` };
    row.getCell("argegvmat").value = { formula: `AA${R}-AB${R}-AC${R}` };
    row.getCell("gvtutari").value = { formula: `AD${R}*AE${R}` };
    row.getCell("agimahsup").value = { formula: `AF${R}-AG${R}` };
    row.getCell("terkingv").value = { formula: `AH${R}*AI${R}` };
    row.getCell("odenecekgv").value = { formula: `AH${R}-AJ${R}` };

    row.getCell("damgaterkin").value = {
      formula: `IFERROR(MAX(((I${R}*0.00759)-250.7)/H${R}*F${R},0),0)`,
    };
    row.getCell("toplamtesvik").value = { formula: `Z${R}+AJ${R}+AL${R}` };

    row.getCell("argemaliyet").value = {
      formula: `I${R}+W${R}+X${R}-Y${R}-Z${R}-AJ${R}-AL${R}`,
    };

    for (let col = 1; col <= colDefs.length; col++) {
      const cell = row.getCell(col);
      cell.border = thinBorder;
      const key = colDefs[col - 1].key;
      if (colColors[key] === YELLOW)
        cell.font = { color: { argb: "FFFF0000" }, name: "Arial", size: 9 };
      if (index % 2 === 0)
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF2F2F2" },
        };
    }
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=aylik_bordro.xlsx",
  );
  await workbook.xlsx.write(res);
  res.end();
});
