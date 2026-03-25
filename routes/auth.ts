import express from "express";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { db } from "../Db.ts";

export const authRouter = express.Router();

const JWT_SECRET = process.env.JWT_SECRET || "arge_gizli_anahtar_degistir";

authRouter.post("/login", async (req, res) => {
  const { email, password } = req.body;
  console.log("Gelen email:", email);
  console.log("Gelen password:", password);

  if (!email || !password)
    return res.status(400).json({ error: "E-posta ve şifre zorunludur." });

  try {
    const result = await db.query(
      "SELECT * FROM kullanicilar WHERE email = $1",
      [email.toLowerCase().trim()],
    );

    if (result.rows.length === 0)
      return res.status(401).json({ error: "E-posta veya şifre hatalı." });

    const user = result.rows[0];
    console.log("Bulunan user:", user);

    const isMatch = await bcrypt.compare(password, user.sifre_hash);
    console.log("isMatch:", isMatch);

    if (!isMatch)
      return res.status(401).json({ error: "E-posta veya şifre hatalı." });

    const token = jwt.sign({ id: user.id, email: user.email }, JWT_SECRET, {
      expiresIn: "8h",
    });

    res.json({
      token,
      kullanici: { id: user.id, ad: user.ad, email: user.email },
    });
  } catch (err) {
    console.error("Login hatası:", err);
    res.status(500).json({ error: "Sunucu hatası." });
  }
});
