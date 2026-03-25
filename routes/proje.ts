import { Router } from "express";
import ExcelJS from "exceljs";
import { db } from "../Db.ts";

export const projeRouter = Router();

export async function setupProje(data: any[]) {
  await db.query(`DROP TABLE IF EXISTS proje_personel`);
  await db.query(`
    CREATE TABLE proje_personel (
      id SERIAL PRIMARY KEY,
      tckimlikno TEXT UNIQUE,
      personeladisoyadi TEXT,
      persdepartmanaciklama TEXT,
      gorevi TEXT
    )
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS projeler (
      id SERIAL PRIMARY KEY,
      proje_adi TEXT,
      proje_kodu TEXT UNIQUE,
      baslangic_tarihi TEXT,
      bitis_tarihi TEXT,
      renk TEXT DEFAULT 'MAVI',
      personel_gideri_carpani NUMERIC DEFAULT 0,
      proje_suresi NUMERIC DEFAULT 0
    )
  `);
  await db.query(
    `ALTER TABLE projeler ADD COLUMN IF NOT EXISTS proje_suresi NUMERIC DEFAULT 0`,
  );
  await db.query(`
    CREATE TABLE IF NOT EXISTS personel_proje_oran (
      id SERIAL PRIMARY KEY,
      tckimlikno TEXT,
      proje_kodu TEXT,
      oran NUMERIC,
      UNIQUE(tckimlikno, proje_kodu)
    )
  `);

  for (const p of data) {
    await db.query(
      `INSERT INTO proje_personel (tckimlikno, personeladisoyadi, persdepartmanaciklama, gorevi)
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

projeRouter.get("/personel", async (req, res) => {
  const result = await db.query(
    `SELECT * FROM proje_personel ORDER BY personeladisoyadi`,
  );
  res.json(result.rows);
});

projeRouter.get("/projeler", async (req, res) => {
  const result = await db.query(`SELECT * FROM projeler ORDER BY id`);
  res.json(result.rows);
});

projeRouter.post("/projeler", async (req, res) => {
  const {
    proje_adi,
    proje_kodu,
    baslangic_tarihi,
    bitis_tarihi,
    renk,
    personel_gideri_carpani,
    proje_suresi,
  } = req.body;
  const result = await db.query(
    `INSERT INTO projeler (proje_adi, proje_kodu, baslangic_tarihi, bitis_tarihi, renk, personel_gideri_carpani, proje_suresi)
     VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING *`,
    [
      proje_adi,
      proje_kodu,
      baslangic_tarihi,
      bitis_tarihi,
      renk ?? "MAVI",
      personel_gideri_carpani ?? 0,
      proje_suresi ?? 0,
    ],
  );
  res.json(result.rows[0]);
});

projeRouter.put("/projeler/:id", async (req, res) => {
  const {
    proje_adi,
    proje_kodu,
    baslangic_tarihi,
    bitis_tarihi,
    renk,
    personel_gideri_carpani,
    proje_suresi,
  } = req.body;
  const result = await db.query(
    `UPDATE projeler SET proje_adi=$1, proje_kodu=$2, baslangic_tarihi=$3,
     bitis_tarihi=$4, renk=$5, personel_gideri_carpani=$6, proje_suresi=$7
     WHERE id=$8 RETURNING *`,
    [
      proje_adi,
      proje_kodu,
      baslangic_tarihi,
      bitis_tarihi,
      renk,
      personel_gideri_carpani,
      proje_suresi ?? 0,
      req.params.id,
    ],
  );
  res.json(result.rows[0]);
});

projeRouter.delete("/projeler/:id", async (req, res) => {
  await db.query(`DELETE FROM projeler WHERE id=$1`, [req.params.id]);
  res.json({ ok: true });
});

projeRouter.get("/oranlar", async (req, res) => {
  const result = await db.query(`SELECT * FROM personel_proje_oran`);
  res.json(result.rows);
});

projeRouter.post("/oranlar", async (req, res) => {
  const { tckimlikno, proje_kodu, oran } = req.body;
  const result = await db.query(
    `INSERT INTO personel_proje_oran (tckimlikno, proje_kodu, oran)
     VALUES ($1,$2,$3) ON CONFLICT (tckimlikno, proje_kodu) DO UPDATE SET oran=EXCLUDED.oran
     RETURNING *`,
    [tckimlikno, proje_kodu, oran],
  );
  res.json(result.rows[0]);
});

projeRouter.get("/excel", async (req, res) => {
  const personelRes = await db.query(
    `SELECT * FROM proje_personel ORDER BY personeladisoyadi`,
  );
  const projelerRes = await db.query(`SELECT * FROM projeler ORDER BY id`);
  const oranlarRes = await db.query(`SELECT * FROM personel_proje_oran`);

  const personelList = personelRes.rows;
  const projeler = projelerRes.rows;
  const oranMap: Record<string, Record<string, number>> = {};
  for (const o of oranlarRes.rows) {
    if (!oranMap[o.tckimlikno]) oranMap[o.tckimlikno] = {};
    oranMap[o.tckimlikno][o.proje_kodu] = parseFloat(o.oran);
  }

  const COLORS: Record<string, string> = {
    MOR: "FF7030A0",
    TURUNCU: "FFFF6600",
    MAVI: "FF00B0F0",
    SARI: "FFFFFF00",
    YESIL: "FF28B43C",
    KIRMIZI: "FFFF0000",
    GRAY: "FFD9D9D9",
    WHITE: "FFFFFFFF",
    LIGHTYELLOW: "FFFFFFCC",
    PURPLE: "7030a0",
  };
  const getColor = (renk: string) => COLORS[renk?.toUpperCase()] ?? "FFD9D9D9";
  const thin = { style: "thin" as const };

  function styleCell(cell: ExcelJS.Cell, opts: any = {}) {
    const {
      bold = false,
      fontSize = 9,
      bgColor,
      alignH = "center",
      fontColor = "FF000000",
      border = true,
    } = opts;
    cell.font = {
      bold,
      size: fontSize,
      name: "Arial",
      color: { argb: fontColor },
    };
    cell.alignment = { horizontal: alignH, vertical: "middle", wrapText: true };
    if (bgColor)
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: bgColor },
      };
    if (border)
      cell.border = { left: thin, right: thin, top: thin, bottom: thin };
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Personel Dağılım");
  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 22;
  sheet.getColumn(3).width = 30;
  sheet.getColumn(4).width = 18;
  const projStartCol = 5;
  projeler.forEach((_, i) => {
    sheet.getColumn(projStartCol + i).width = 14;
  });
  const totalCol = projStartCol + projeler.length;
  sheet.getColumn(totalCol).width = 10;

  sheet.getRow(1).height = 50;
  projeler.forEach((p, i) => {
    const cell = sheet.getRow(1).getCell(projStartCol + i);
    cell.value = p.proje_adi;
    styleCell(cell, { bold: true, bgColor: getColor(p.renk) });
  });
  sheet.getRow(2).height = 20;
  projeler.forEach((p, i) => {
    const cell = sheet.getRow(2).getCell(projStartCol + i);
    cell.value = `Proje Kodu:\n${p.proje_kodu}`;
    styleCell(cell, { bold: true, bgColor: getColor(p.renk) });
  });
  sheet.getRow(3).height = 55;
  projeler.forEach((p, i) => {
    const cell = sheet.getRow(3).getCell(projStartCol + i);
    cell.value = `Başlangıç: ${p.baslangic_tarihi}\nBitiş: ${p.bitis_tarihi}`;
    styleCell(cell, { bgColor: getColor(p.renk), fontSize: 8 });
  });
  sheet.getRow(4).height = 25;
  ["TC KİMLİK NO", "PERSONEL ADI SOYADI", "DEPARTMAN", "GÖREVİ"].forEach(
    (h, i) => {
      const cell = sheet.getRow(4).getCell(i + 1);
      cell.value = h;
      styleCell(cell, {
        bold: true,
        bgColor: COLORS.PURPLE,
        fontColor: "FFFFFFFF",
      });
    },
  );
  projeler.forEach((p, i) => {
    styleCell(sheet.getRow(4).getCell(projStartCol + i), {
      bgColor: getColor(p.renk),
    });
  });
  sheet.getRow(4).getCell(totalCol).value = "TOPLAM";
  styleCell(sheet.getRow(4).getCell(totalCol), {
    bold: true,
    bgColor: COLORS.PURPLE,
    fontColor: "FFFFFFFF",
  });

  personelList.forEach((p, idx) => {
    const rowNum = 5 + idx;
    sheet.getRow(rowNum).height = 18;
    const bg = idx % 2 === 0 ? COLORS.WHITE : COLORS.LIGHTYELLOW;
    [
      p.tckimlikno,
      p.personeladisoyadi,
      p.persdepartmanaciklama,
      p.gorevi ?? "",
    ].forEach((val, ci) => {
      const cell = sheet.getRow(rowNum).getCell(ci + 1);
      cell.value = val;
      styleCell(cell, { bgColor: bg, alignH: "left" });
    });
    const oranlar = oranMap[p.tckimlikno] ?? {};
    const toplamParts: string[] = [];
    projeler.forEach((proj, pi) => {
      const col = projStartCol + pi;
      const colLetter = sheet.getColumn(col).letter;
      const cell = sheet.getRow(rowNum).getCell(col);
      const oran = oranlar[proj.proje_kodu];
      cell.value = oran ?? null;
      if (oran !== undefined) cell.numFmt = "0.00";
      styleCell(cell, { bgColor: bg });
      toplamParts.push(`${colLetter}${rowNum}`);
    });
    const totalCell = sheet.getRow(rowNum).getCell(totalCol);
    totalCell.value = { formula: `IFERROR(SUM(${toplamParts.join(",")}),0)` };
    styleCell(totalCell, { bgColor: bg });
  });

  const lastDataRow = 4 + personelList.length;
  const sumStart = lastDataRow + 3;
  const totalColLetter = sheet.getColumn(totalCol).letter;
  const genelToplamFormula = `SUM(${totalColLetter}5:${totalColLetter}${lastDataRow})`;

  const adamAyRow = sumStart;
  sheet.getRow(adamAyRow).getCell(4).value = "PROJE ADAM/AY";
  styleCell(sheet.getRow(adamAyRow).getCell(4), {
    bold: true,
    alignH: "left",
    bgColor: COLORS.LIGHTYELLOW,
  });
  projeler.forEach((_, i) => {
    const col = projStartCol + i;
    const colLetter = sheet.getColumn(col).letter;
    const cell = sheet.getRow(adamAyRow).getCell(col);
    cell.value = { formula: `SUM(${colLetter}5:${colLetter}${lastDataRow})` };
    styleCell(cell, { bold: true, bgColor: COLORS.LIGHTYELLOW });
  });

  const carpanRow = sumStart + 1;
  sheet.getRow(carpanRow).getCell(4).value = "Personel gideri çarpanı";
  styleCell(sheet.getRow(carpanRow).getCell(4), { alignH: "left" });
  projeler.forEach((_, i) => {
    const col = projStartCol + i;
    const colLetter = sheet.getColumn(col).letter;
    const cell = sheet.getRow(carpanRow).getCell(col);
    cell.value = {
      formula: `IFERROR(${colLetter}${adamAyRow}/${genelToplamFormula},0)`,
    };
    cell.numFmt = "0.00";
    styleCell(cell, { bgColor: COLORS.LIGHTYELLOW });
  });

  const suresiRow = sumStart + 2;
  sheet.getRow(suresiRow).getCell(4).value = "PROJE SÜRESİ (AY)";
  styleCell(sheet.getRow(suresiRow).getCell(4), { alignH: "left" });
  projeler.forEach((p, i) => {
    const cell = sheet.getRow(suresiRow).getCell(projStartCol + i);
    cell.value = parseFloat(p.proje_suresi) || 0;
    cell.numFmt = "0";
    styleCell(cell, { bgColor: COLORS.LIGHTYELLOW });
  });

  const gorevliRow = sumStart + 3;
  sheet.getRow(gorevliRow).getCell(4).value = "PROJE GÖREVLİ PERSONEL SAYISI";
  styleCell(sheet.getRow(gorevliRow).getCell(4), { alignH: "left" });
  projeler.forEach((_, i) => {
    const col = projStartCol + i;
    const colLetter = sheet.getColumn(col).letter;
    const cell = sheet.getRow(gorevliRow).getCell(col);
    cell.value = {
      formula: `COUNTA(${colLetter}5:${colLetter}${lastDataRow})`,
    };
    styleCell(cell, {});
  });

  const toplamRow = sumStart + 4;
  sheet.getRow(toplamRow).getCell(4).value = "Proje Toplam Adam/ay";
  styleCell(sheet.getRow(toplamRow).getCell(4), { alignH: "left" });
  projeler.forEach((_, i) => {
    const col = projStartCol + i;
    const colLetter = sheet.getColumn(col).letter;
    const cell = sheet.getRow(toplamRow).getCell(col);
    cell.value = {
      formula: `${colLetter}${carpanRow}*${colLetter}${suresiRow}`,
    };
    styleCell(cell, { bgColor: COLORS.LIGHTYELLOW });
  });

  const leg1 = sheet.getRow(toplamRow + 2);
  leg1.getCell(1).value = "Yeni eklenen";
  styleCell(leg1.getCell(1), { bgColor: COLORS.YESIL, alignH: "left" });
  const leg2 = sheet.getRow(toplamRow + 3);
  leg2.getCell(1).value = "Ayrılan";
  styleCell(leg2.getCell(1), {
    bgColor: COLORS.KIRMIZI,
    fontColor: "FFFFFFFF",
    alignH: "left",
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
