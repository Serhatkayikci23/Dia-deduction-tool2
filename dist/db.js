import pg from "pg";
export const db = new pg.Pool({
    host: "127.0.0.1",
    port: 5432,
    database: "dia-deduction-tool2",
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
});
export const baseUrl = "https://diademo.ws.dia.com.tr/api/v3";
export const firmCode = 1;
export const periodCode = 1;
export let sessionId = null;
export function setSessionId(id) {
    sessionId = id;
}
