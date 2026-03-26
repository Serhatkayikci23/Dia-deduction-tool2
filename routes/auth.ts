import express from "express";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { db } from "../db.js";
import type { User } from "../types/user.ts";

export const authRouter = express.Router();

const JWT_SECRET = process.env.JWT_SECRET || "arge_gizli_anahtar_degistir";

authRouter.post("/login", async (req, res) => {
  const { email, password } = req.body;
  console.log("Gelen email:", email);
  console.log("Gelen password:", password);

  if (!email || !password)
    return res.status(400).json({ error: "E-posta ve şifre zorunludur." });

  try {
    const result = await db.query<User>(
      "SELECT * FROM kullanicilar WHERE email = $1",
      [email.toLowerCase().trim()],
    );

    const user = result.rows[0];
    if (!user)
      return res.status(401).json({ error: "E-posta veya şifre hatalı." });

    console.log("Bulunan user:", user);

    const isMatch = await bcrypt.compare(password, user.sifre_hash);
    console.log("isMatch:", isMatch);

    if (!isMatch)
      return res.status(401).json({ error: "E-posta veya şifre hatalı." });

    const token = jwt.sign({ id: user.id, email: user.email }, JWT_SECRET, {
      expiresIn: "1800000",
    });

    res.status(200).json({
      token,
      kullanici: { id: user.id, ad: user.ad, email: user.email },
    });
  } catch (err) {
    console.error("Login hatası:", err);
    res.status(500).json({ error: "Sunucu hatası." });
  }
});

authRouter.get("/kullanicilar", async (req, res) => {
  const result = await db.query<User>(
    "SELECT id, ad, email, olusturma FROM kullanicilar ORDER BY id",
  );
  res.json(result.rows);
});

authRouter.get("/kullanicilar/:id", async (req, res) => {
  const { id } = req.params;

  try {
    const result = await db.query<User>(
      "select * from kullanicilar where id = $1",
      [id],
    );

    if (result.rows.length === 0)
      return res.status(404).json({ error: "Kullanıcı bulunamadı." });

    const user = result.rows[0];

    return res.status(200).json(user);
  } catch (err) {
    console.error("login getirilemedi: ", err);
    res.status(500).json({ error: "Sunucu hatası." });
  }
});

authRouter.post("/kullanicilar/ekle", async (req, res) => {
  const { ad, email, password } = req.body;
  if (!ad || !email || !password)
    return res.status(400).json({ error: "Tüm alanlar zorunludur." });
  const hash = await bcrypt.hash(password, 10);
  await db.query(
    "INSERT INTO kullanicilar (ad, email, sifre_hash) VALUES ($1, $2, $3)",
    [ad, email.toLowerCase().trim(), hash],
  );
  res.json({ basarili: true });
});

authRouter.post("/kullanicilar/sil", async (req, res) => {
  const { id } = req.body;
  await db.query("DELETE FROM kullanicilar WHERE id = $1", [id]);
  res.json({ basarili: true });
});

authRouter.post("/kullanicilar/sifre", async (req, res) => {
  const { id, password } = req.body;
  const hash = await bcrypt.hash(password, 10);
  await db.query("UPDATE kullanicilar SET sifre_hash = $1 WHERE id = $2", [
    hash,
    id,
  ]);
  res.json({ basarili: true });
});

export function authMiddleware(req: any, res: any, next: any) {
  const token = req.headers.authorization?.split(" ")[1];
  if (!token) return res.status(401).json({ error: "Token gerekli!" });
  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET || "gizli");
    req.user = decoded;
    next();
  } catch {
    return res.status(401).json({ error: "geçersiz token" });
  }
}
