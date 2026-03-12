// Supabase Configuration
window.SUPABASE_CONFIG = {
  url: 'https://gpiujhtcuagyopinszbe.supabase.co',
  anonKey: 'sb_publishable_eFDWMkF_h-UqdJulydi8pQ_IqotB83e'
};

// Initialize Supabase Client
if (window.supabase) {
  window.supabaseClient = window.supabase.createClient(window.SUPABASE_CONFIG.url, window.SUPABASE_CONFIG.anonKey);
} else {
  console.error('Supabase library not loaded. Make sure the CDN script is included.');
}
