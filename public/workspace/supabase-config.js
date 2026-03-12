// Supabase Configuration
window.SUPABASE_CONFIG = {
  url: 'https://gpiujhtcuagyopinszbe.supabase.co',
  anonKey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdwaXVqaHRjdWFneW9waW5zemJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMzMjI3NDAsImV4cCI6MjA4ODg5ODc0MH0.5VFWVFKuFvTbFbBJlpSl7c8U8cueb_uBWgIV5fSjrP4'
};

// Initialize Supabase Client
if (window.supabase) {
  window.supabaseClient = window.supabase.createClient(window.SUPABASE_CONFIG.url, window.SUPABASE_CONFIG.anonKey);
} else {
  console.error('Supabase library not loaded. Make sure the CDN script is included.');
}
