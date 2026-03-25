import { Router } from "express";
import ExcelJS from "exceljs";
import { db } from "../db.js";
export const personelRouter = Router();
export async function setupPersonel(data) {
    await db.query(`DROP TABLE IF EXISTS personel`);
    await db.query(`
    CREATE TABLE personel (
      id SERIAL PRIMARY KEY,
      tckimlikno TEXT UNIQUE,
      personeladisoyadi TEXT,
      persdepartmanaciklama TEXT,
      gorevi TEXT
    )
  `);
    await db.query(`
    CREATE TABLE IF NOT EXISTS personel_proje_detay (
      id SERIAL PRIMARY KEY,
      tckimlikno TEXT,
      proje_kodu TEXT,
      argegun NUMERIC DEFAULT 0,
      digergun NUMERIC DEFAULT 0,
      toplamgun NUMERIC DEFAULT 0,
      bruttemel NUMERIC DEFAULT 0,
      fazlamesai NUMERIC DEFAULT 0,
      aylikust NUMERIC DEFAULT 0,
      argeorani NUMERIC DEFAULT 0,
      agi NUMERIC DEFAULT 0,
      gvorani NUMERIC DEFAULT 0,
      UNIQUE(tckimlikno, proje_kodu)
    )
  `);
    const limitliData = data.slice(0, 3);
    for (const p of limitliData) {
        await db.query(`INSERT INTO personel (tckimlikno, personeladisoyadi, persdepartmanaciklama, gorevi)
       VALUES ($1,$2,$3,$4) ON CONFLICT (tckimlikno) DO UPDATE SET
         personeladisoyadi=EXCLUDED.personeladisoyadi,
         persdepartmanaciklama=EXCLUDED.persdepartmanaciklama,
         gorevi=EXCLUDED.gorevi`, [
            p.tckimlikno,
            p.personeladisoyadi,
            p.persdepartmanaciklama,
            p.gorevi ?? "",
        ]);
    }
}
personelRouter.get("/", async (req, res) => {
    const result = await db.query(`SELECT p.*, 
            COALESCE(b.argegun, 0) as argegun,
            COALESCE(b.digergun, 0) as digergun,
            COALESCE(b.toplamgun, 0) as toplamgun,
            COALESCE(b.bruttemel, 0) as bruttemel,
     COALESCE(b.fazlamesai, 0) as fazlamesai,
          COALESCE(b.damgaterkin, 0) as damgaterkin
     FROM personel p
     LEFT JOIN bordro_detay b ON p.tckimlikno = b.tckimlikno
     ORDER BY p.personeladisoyadi`);
    res.json(result.rows);
});
personelRouter.get("/projeler", async (req, res) => {
    const result = await db.query(`SELECT * FROM projeler ORDER BY id`);
    res.json(result.rows);
});
personelRouter.get("/detay/:proje_kodu", async (req, res) => {
    const result = await db.query(`SELECT * FROM personel_proje_detay WHERE proje_kodu=$1`, [req.params.proje_kodu]);
    res.json(result.rows);
});
personelRouter.put("/detay/:proje_kodu/:tckimlikno", async (req, res) => {
    const { proje_kodu, tckimlikno } = req.params;
    const { aylikust, argeorani, agi, gvorani } = req.body;
    await db.query(`INSERT INTO personel_proje_detay
       (tckimlikno, proje_kodu, aylikust, argeorani, agi, gvorani)
     VALUES ($1,$2,$3,$4,$5,$6)
     ON CONFLICT (tckimlikno, proje_kodu) DO UPDATE SET
       aylikust=$3, argeorani=$4, agi=$5, gvorani=$6`, [tckimlikno, proje_kodu, aylikust, argeorani, agi, gvorani]);
    res.json({ ok: true });
});
personelRouter.get("/excel", async (req, res) => {
    const personelRes = await db.query(`SELECT * FROM personel ORDER BY personeladisoyadi`);
    const projelerRes = await db.query(`SELECT * FROM projeler ORDER BY id`);
    const bordroRes = await db.query(`SELECT * FROM bordro_detay`);
    const bordroMap = {};
    for (const b of bordroRes.rows)
        bordroMap[b.tckimlikno] = b;
    const oranlarRes = await db.query(`SELECT * FROM personel_proje_oran`);
    const oranMap = {};
    for (const o of oranlarRes.rows) {
        if (!oranMap[o.tckimlikno])
            oranMap[o.tckimlikno] = {};
        oranMap[o.tckimlikno][o.proje_kodu] = parseFloat(o.oran) || 0;
    }
    const r_isci = parseFloat(req.query.sgk_isci) || 0.14;
    const r_issizlik = parseFloat(req.query.sgk_issizlik) || 0.01;
    const r_isv = parseFloat(req.query.sgk_isv) || 0.2175;
    const r_isvisz = parseFloat(req.query.sgk_isvisz) || 0.02;
    const pct_isci = Math.round(r_isci * 10000);
    const pct_issizlik = Math.round(r_issizlik * 10000);
    const pct_isv = Math.round(r_isv * 10000);
    const pct_isvisz = Math.round(r_isvisz * 10000);
    const YELLOW = "FFFFFF00";
    const GREEN = "FF00B050";
    const BROWN = "FF843C0C";
    const NAVY = "FF1F3864";
    const GRAY = "FF808080";
    const PURPLE = "7030a0";
    const colColors = {
        sno: PURPLE,
        tckimlikno: PURPLE,
        ad: PURPLE,
        departman: PURPLE,
        argegun: PURPLE,
        digergun: PURPLE,
        toplamgun: PURPLE,
        bruttemel: PURPLE,
        fazlamesai: PURPLE,
        aylikust: YELLOW,
        gunlukust: YELLOW,
        argeaylikust: YELLOW,
        s5510aylik: YELLOW,
        toplambrut: PURPLE,
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
        argeucret: PURPLE,
        sgkisci2: PURPLE,
        sgkissizlik2: PURPLE,
        argegvmat: PURPLE,
        gvorani: PURPLE,
        gvtutari: PURPLE,
        agi: PURPLE,
        agimahsup: PURPLE,
        argeorani: PURPLE,
        terkingv: PURPLE,
        odenecekgv: PURPLE,
        damgaterkin: NAVY,
        toplamtesvik: PURPLE,
        argemaliyet: PURPLE,
    };
    const headerLabels = {
        sno: "S/NO",
        tckimlikno: "TC KİMLİK NO",
        ad: "AD SOYAD",
        departman: "DEPARTMAN",
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
        sgkmatrah: "5510 SGK MATRAHI (DİĞER FAAL, İKRAMİYE PRİM, MESAİ)",
        sgkisci: `SGK İşçi Payı (${r_isci})`,
        sgkissizlik: `SGK İşçi İşsizlik Payı (${r_issizlik})`,
        sgkisv: `SGK İŞV PAYI (${r_isv})`,
        sgkisvisz: `SGK İşveren İşsizlik Payı (${r_isvisz})`,
        sgkindirim: "DİĞER FAALİYETLER KAPSAMINDA HESAPLANAN SGK İNDİRİMİ %5",
        argesigorta: "ARGE MERKEZİNDEKİ ÜCRET (SİGORTA MATRAHI)",
        sgkisvarge: `SGK İŞV PAYI (${r_isv})`,
        sgkisvisz2: `SGK İşveren İşsizlik Payı (${r_isvisz})`,
        argesgk5: "ARGE MATRAHINDAN 5510 SGK İNDİRİMİ %5",
        sgk5746: "5746 SGK İNDİRİMİ %50",
        argeucret: "ARGE MERKEZİNDEKİ ÜCRET",
        sgkisci2: `SGK İşçi Payı (${r_isci})`,
        sgkissizlik2: `SGK İşçi İşsizlik Payı (${r_issizlik})`,
        argegvmat: "AR-GE GELİR VERGİSİ MATRAHI",
        gvorani: "GELİR VERGİSİ ORANI",
        gvtutari: "GELİR VERGİSİ TUTARI",
        agi: "AGİ",
        agimahsup: "AGİ MAHSUBU SONRASI GELİR VERGİSİ TUTARI",
        argeorani: "ARGE İSTİSNA ORANI",
        terkingv: "TERKİN EDİLECEK GELİR VERGİSİ TUTARI",
        odenecekgv: "ÖDENECEK GV STOPAJ",
        damgaterkin: "5746 SAYILI KANUN KAPSAMINDA TERKİN EDİLECEK DAMGA VERGİSİ",
        toplamtesvik: "TOPLAM TEŞVİK TUTARI (SGK,GV,DV)",
        argemaliyet: "ARGE İŞV MALİYETİ",
    };
    const colDefs = [
        { key: "sno", width: 6 },
        { key: "tckimlikno", width: 16 },
        { key: "ad", width: 22 },
        { key: "departman", width: 45 },
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
        { key: "sgkindirim", width: 14 },
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
    const thinBorder = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
    };
    const workbook = new ExcelJS.Workbook();
    for (const proje of projelerRes.rows) {
        const sheetName = proje.proje_kodu
            .replace(/[\/\\*?\[\]:]/g, "_")
            .substring(0, 31);
        const sheet = workbook.addWorksheet(sheetName);
        sheet.columns = colDefs.map((c) => ({ key: c.key, width: c.width }));
        let rowNum = 1;
        const detayRes = await db.query(`SELECT * FROM personel_proje_detay WHERE proje_kodu=$1`, [proje.proje_kodu]);
        const detayMap = {};
        for (const d of detayRes.rows)
            detayMap[d.tckimlikno] = d;
        const projeRow = sheet.getRow(rowNum);
        projeRow.height = 22;
        projeRow.getCell(1).value = `${proje.proje_kodu} — ${proje.proje_adi}`;
        projeRow.getCell(1).font = {
            bold: true,
            size: 12,
            name: "Arial",
            color: { argb: "FFFFFFFF" },
        };
        projeRow.getCell(1).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF1F3864" },
        };
        sheet.mergeCells(rowNum, 1, rowNum, colDefs.length);
        rowNum++;
        const groupRow = sheet.getRow(rowNum);
        groupRow.height = 30;
        const groups = [
            {
                label: "SGK Aylık Ve Günlük Üst Sınır İşlemleri",
                startCol: 10,
                endCol: 13,
                color: YELLOW,
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
                color: PURPLE,
            },
        ];
        groups.forEach(({ label, startCol, endCol, color }) => {
            const cell = groupRow.getCell(startCol);
            cell.value = label;
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: color },
            };
            cell.font = {
                bold: true,
                color: { argb: color === YELLOW ? "0529a1" : "FFFFFFFF" },
                name: "Arial",
                size: 9,
            };
            cell.alignment = {
                horizontal: "center",
                vertical: "middle",
                wrapText: true,
            };
            sheet.mergeCells(rowNum, startCol, rowNum, endCol);
        });
        rowNum++;
        const headerRow = sheet.getRow(rowNum);
        headerRow.height = 60;
        colDefs.forEach((col, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = headerLabels[col.key] ?? col.key;
            const bg = colColors[col.key] ?? "FFFFFFFF";
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
            cell.font = {
                bold: true,
                color: { argb: bg === YELLOW ? "0529a1" : "FFFFFFFF" },
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
        rowNum++;
        const dataStartRow = rowNum;
        personelRes.rows.forEach((p, index) => {
            const d = detayMap[p.tckimlikno] || {};
            const bd = bordroMap[p.tckimlikno] || {};
            const row = sheet.getRow(rowNum);
            row.height = 18;
            row.getCell("sno").value = index + 1;
            row.getCell("tckimlikno").value = p.tckimlikno;
            row.getCell("ad").value = p.personeladisoyadi;
            row.getCell("departman").value = p.persdepartmanaciklama;
            row.getCell("argegun").value = parseFloat(bd.argegun) || 0;
            row.getCell("digergun").value = parseFloat(bd.digergun) || 0;
            row.getCell("toplamgun").value = parseFloat(bd.toplamgun) || 0;
            const adamAyOrani = oranMap[p.tckimlikno]?.[proje.proje_kodu] || 0;
            row.getCell("bruttemel").value =
                (parseFloat(bd.bruttemel) || 0) * adamAyOrani;
            row.getCell("fazlamesai").value =
                (parseFloat(bd.fazlamesai) || 0) * adamAyOrani;
            row.getCell("aylikust").value = parseFloat(d.aylikust) || 0;
            row.getCell("gvorani").value = parseFloat(d.gvorani) || 0;
            row.getCell("agi").value = parseFloat(d.agi) || 0;
            row.getCell("argeorani").value = parseFloat(d.argeorani) || 0;
            const R = rowNum;
            row.getCell("gunlukust").value = { formula: `J${R}/30` };
            row.getCell("argeaylikust").value = { formula: `E${R}*K${R}` };
            row.getCell("s5510aylik").value = { formula: `F${R}*K${R}` };
            row.getCell("toplambrut").value = { formula: `H${R}+I${R}` };
            row.getCell("sgkmatrah").value = {
                formula: `IFERROR(IF(G${R}=0,0,(H${R}/G${R}*F${R})+I${R}),0)`,
            };
            row.getCell("sgkisci").value = { formula: `O${R}*${pct_isci}/10000` };
            row.getCell("sgkissizlik").value = {
                formula: `O${R}*${pct_issizlik}/10000`,
            };
            row.getCell("sgkisv").value = { formula: `O${R}*${pct_isv}/10000` };
            row.getCell("sgkisvisz").value = { formula: `O${R}*${pct_isvisz}/10000` };
            row.getCell("sgkindirim").value = { formula: `O${R}*5/100` };
            row.getCell("argesigorta").value = {
                formula: `IFERROR(IF(G${R}=0,0,E${R}/G${R}*H${R}),0)`,
            };
            row.getCell("sgkisvarge").value = { formula: `U${R}*${pct_isv}/10000` };
            row.getCell("sgkisvisz2").value = {
                formula: `U${R}*${pct_isvisz}/10000`,
            };
            row.getCell("argesgk5").value = { formula: `U${R}*5/100` };
            row.getCell("sgk5746").value = { formula: `U${R}*1675/20000` };
            row.getCell("sgkisci2").value = { formula: `Z${R}*${pct_isci}/10000` };
            row.getCell("sgkissizlik2").value = {
                formula: `Z${R}*${pct_issizlik}/10000`,
            };
            row.getCell("argeucret").value = {
                formula: `IFERROR(IF(G${R}=0,0,E${R}/G${R}*H${R}),0)`,
            };
            row.getCell("argegvmat").value = { formula: `Z${R}-AA${R}-AB${R}` };
            row.getCell("gvtutari").value = { formula: `AC${R}*AD${R}` };
            row.getCell("agimahsup").value = { formula: `AE${R}-AF${R}` };
            row.getCell("terkingv").value = { formula: `AG${R}*AH${R}` };
            row.getCell("odenecekgv").value = { formula: `AG${R}-AI${R}` };
            const damgaAylik = parseFloat(bd.damgaterkin) || 0;
            row.getCell("damgaterkin").value = damgaAylik * adamAyOrani;
            row.getCell("toplamtesvik").value = { formula: `Y${R}+AI${R}+AK${R}` };
            row.getCell("argemaliyet").value = {
                formula: `H${R}+V${R}+W${R}-X${R}-Y${R}-AI${R}-AK${R}`,
            };
            for (let col = 1; col <= colDefs.length; col++) {
                const cell = row.getCell(col);
                cell.border = thinBorder;
                const key = colDefs[col - 1].key;
                if (colColors[key] === YELLOW)
                    cell.font = { color: { argb: "0529a1" }, name: "Arial", size: 9 };
                if (index % 2 === 0)
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FFF2F2F2" },
                    };
            }
            rowNum++;
        });
        const lastDataRow = rowNum - 1;
        rowNum++;
        const totalRow = sheet.getRow(rowNum);
        [
            ["argegun", "E"],
            ["digergun", "F"],
            ["toplamgun", "G"],
            ["bruttemel", "H"],
            ["fazlamesai", "I"],
            ["sgkmatrah", "O"],
            ["sgkisci", "P"],
            ["sgkissizlik", "Q"],
            ["sgkisv", "R"],
            ["sgkindirim", "T"],
            ["argesigorta", "U"],
            ["sgk5746", "Y"],
            ["toplamtesvik", "AL"],
        ].forEach(([key, col]) => {
            totalRow.getCell(key).value = {
                formula: `SUM(${col}${dataStartRow}:${col}${lastDataRow})`,
            };
        });
        rowNum += 3;
    }
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=personel.xlsx");
    await workbook.xlsx.write(res);
    res.end();
});
