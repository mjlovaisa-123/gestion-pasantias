import { supabase } from "./supabaseClient";

async function get(key) {
  const { data, error } = await supabase
    .from("app_data")
    .select("value")
    .eq("key", key)
    .maybeSingle();

  if (error) throw error;
  if (!data) throw new Error("Clave no encontrada: " + key);
  return { key, value: JSON.stringify(data.value), shared: true };
}

async function set(key, value) {
  const parsed = typeof value === "string" ? JSON.parse(value) : value;
  const { error } = await supabase
    .from("app_data")
    .upsert({ key, value: parsed, updated_at: new Date().toISOString() });

  if (error) throw error;
  return { key, value, shared: true };
}

async function del(key) {
  const { error } = await supabase.from("app_data").delete().eq("key", key);
  if (error) throw error;
  return { key, deleted: true, shared: true };
}

async function list() {
  const { data, error } = await supabase.from("app_data").select("key");
  if (error) throw error;
  return { keys: (data || []).map((r) => r.key), shared: true };
}

export function installStorageShim() {
  window.storage = { get, set, delete: del, list };
}
