// hb-invite: redeem a one-time code issued by a 園 admin.
//   kind 'join'  -> create a new account with the chosen password and add it to the 園
//   kind 'reset' -> set a new password for an existing member
// The caller is not signed in, so this function validates the code itself (verify_jwt = false).
// The service role key only lives in this function's environment; it is never sent to the browser.
import { createClient } from "npm:@supabase/supabase-js@2.49.4";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (status: number, body: unknown) =>
  new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json" } });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json(405, { ok: false, error: "method" });
  let body: Record<string, unknown>;
  try { body = await req.json(); } catch { return json(400, { ok: false, error: "bad_json", message: "送信内容を読めませんでした" }); }
  const email = String(body.email ?? "").trim().toLowerCase();
  const code = String(body.code ?? "");
  const password = String(body.password ?? "");
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || email.length > 254) return json(400, { ok: false, error: "email", message: "メールアドレスの形が正しくありません" });
  if (code.replace(/[^0-9A-Za-z]/g, "").length !== 12) return json(400, { ok: false, error: "code", message: "コードは12文字です（ハイフンはあってもなくても構いません）" });
  if (password.length < 8 || password.length > 72) return json(400, { ok: false, error: "password", message: "パスワードは8文字以上にしてください" });

  const url = Deno.env.get("SUPABASE_URL")!;
  const key = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(url, key, { auth: { persistSession: false, autoRefreshToken: false } });

  const { data: peek, error: pe } = await admin.rpc("hb_invite_peek", { p_email: email, p_code: code });
  if (pe) return json(500, { ok: false, error: "server", message: "確認できませんでした。時間をおいてもう一度お試しください" });
  if (!peek?.ok) return json(400, { ok: false, error: "invalid", message: "コードが正しくないか、期限が切れています。園の管理者に新しいコードを出してもらってください" });

  let userId: string | null = peek.user_id ?? null;
  if (peek.kind === "join") {
    if (userId) return json(409, { ok: false, error: "exists", message: "このメールアドレスはすでに登録されています。ログインしてから「設定 → 招待コードで別の園に参加」でコードを入れてください" });
    const { data: cu, error: ce } = await admin.auth.admin.createUser({ email, password, email_confirm: true });
    if (ce || !cu?.user) return json(400, { ok: false, error: "create", message: "アカウントを作れませんでした（" + (ce?.message ?? "") + "）" });
    userId = cu.user.id;
  } else {
    if (!userId) return json(400, { ok: false, error: "invalid", message: "コードが正しくないか、期限が切れています" });
    const { error: ue } = await admin.auth.admin.updateUserById(userId, { password });
    if (ue) return json(400, { ok: false, error: "update", message: "パスワードを変えられませんでした（" + ue.message + "）" });
  }
  const { data: done, error: de } = await admin.rpc("hb_invite_complete", { p_invite: peek.invite_id, p_user: userId });
  if (de || !done?.ok) return json(409, { ok: false, error: "used", message: "このコードはすでに使われています" });
  return json(200, { ok: true, kind: peek.kind });
});
