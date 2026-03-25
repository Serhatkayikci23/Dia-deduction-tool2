CREATE TABLE IF NOT EXISTS kullanicilar (
  id         SERIAL PRIMARY KEY,
  ad         VARCHAR(100) NOT NULL,
  email      VARCHAR(150) UNIQUE NOT NULL,
  sifre_hash TEXT NOT NULL,
  olusturma  TIMESTAMP DEFAULT NOW()
);

-- admin@panu.com / Admin1234
INSERT INTO kullanicilar (ad, email, sifre_hash) VALUES (
  'Admin',
  'admin@panu.com',
  '$2a$10$N9qo8uLOickgx2ZMRZoMyeIjZAgcfl7p92ldGxad68LJZdL17lhWy'
) ON CONFLICT (email) DO NOTHING;