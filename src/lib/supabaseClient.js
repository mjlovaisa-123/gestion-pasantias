import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://hcszrsmfoowstapqglld.supabase.co";
const SUPABASE_KEY = "sb_publishable_BchKIhFoeSMPECGb3MaX1A_YuBrnvRf";

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
