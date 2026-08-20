// Supabase 연결 설정 (foreon2 프로젝트)
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://knugcmsyezmbmywmmgom.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_3K_yQbWRQXhtFP9kIB5xsg_wPtml1Oo";

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);
