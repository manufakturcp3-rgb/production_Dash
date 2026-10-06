import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.PUBLIC_SUPABASE_URL || 'https://awnjjzyyyxflvhpsnwby.supabase.co';
const supabaseKey = import.meta.env.PUBLIC_SUPABASE_ANON_KEY || 'sb_publishable_F2QptrbgTXqFo8pZvtNgZg_mAfhuUSS';

export const supabase = createClient(supabaseUrl, supabaseKey);

export type SyncLog = {
  id: string;
  timestamp: string;
  source1_status: string;
  source1_rows: number;
  source2_status: string;
  source2_rows: number;
};
