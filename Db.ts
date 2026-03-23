import pg from "pg";

export const db = new pg.Pool({
  host: "127.0.0.1",
  port: 5432,
  database: "dia-deduction-tool2",
  user: "postgres",
  password: "123",
});

export const baseUrl = "https://diademo.ws.dia.com.tr/api/v3";
export const firmCode = 1;
export const periodCode = 1;

export let sessionId: string | null = null;

export function setSessionId(id: string) {
  sessionId = id;
}
