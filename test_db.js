process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://rjjarptqfsnainsftxli.supabase.co';
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJqamFycHRxZnNuYWluc2Z0eGxpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE5NDMyMTMsImV4cCI6MjA5NzUxOTIxM30._RzWXDbptuYYgncspEXzEp_QEMgF-2c78V7uyJFC7FM';

const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function check() {
  console.log("Fetching Rent transactions...");
  const { data, error } = await supabase
    .from('transactions')
    .select('*')
    .eq('category', 'Rent');
  
  if (error) {
    console.error("Error:", error);
  } else {
    console.log("Rent transactions:", data);
  }

  console.log("\nFetching ALL transactions in the room...");
  const { data: allTxs, error: errAll } = await supabase
    .from('transactions')
    .select('*')
    .eq('is_shared', true);

  if (errAll) {
    console.error("Error:", errAll);
  } else {
    console.log("All shared transactions:", allTxs);
  }
}

check();
