import "dotenv/config";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { db, baseUrl, firmCode, periodCode, setSessionId } from "./db.js";
import { personelRouter, setupPersonel } from "./routes/personel.js";
import { projeRouter, setupProje } from "./routes/proje.js";
import { bordroRouter, setupBordro } from "./routes/bordro.js";
import { authRouter, authMiddleware } from "./routes/auth.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());
app.use("/image", express.static("image"));

app.use(express.static(path.join(__dirname, "public-ana")));

app.use("/api/auth", authRouter);
app.use("/api/personel", authMiddleware, personelRouter);
app.use("/api/proje", authMiddleware, projeRouter);
app.use("/api/bordro", authMiddleware, bordroRouter);

async function login() {
  const response = await fetch(baseUrl + "/sis/json", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
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
  setSessionId(data.msg);
  console.log("Login başarılı, session:", data.msg);
  return data.msg;
}

async function diaPersonelCek(sessionId: string, cols: string[]) {
  const response = await fetch(baseUrl + "/per/json", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      per_personel_puantaj_listele: {
        session_id: sessionId,
        firma_kodu: firmCode,
        donem_kodu: periodCode,
        params: { selectedcolumns: cols },
      },
    }),
  });
  const json = await response.json();
  return json.result;
}

async function main() {
  await db.connect();
  console.log("PostgreSQL bağlantısı kuruldu");

  const sid = await login();

  const data3 = await diaPersonelCek(sid, [
    "tckimlikno",
    "personeladisoyadi",
    "persdepartmanaciklama",
  ]);
  const data4 = await diaPersonelCek(sid, [
    "tckimlikno",
    "personeladisoyadi",
    "persdepartmanaciklama",
    "gorevi",
  ]);
  const limitliData = data4.slice(0, 3);

  await setupPersonel(limitliData);
  console.log("Personel setup OK");
  await setupProje(limitliData);
  console.log("Proje setup OK");
  await setupBordro(limitliData);
  console.log("Bordro setup OK");

  console.log(`${limitliData.length} personel aktarıldı`);

  app.listen(5000, () => console.log("Sunucu: http://localhost:5000"));
}

main().catch(console.error);
