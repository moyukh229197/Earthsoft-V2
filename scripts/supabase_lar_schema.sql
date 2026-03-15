-- ==================================================================
-- Earthsoft LAR (Lowest Accepted Rates) Database Schema
-- Run this SQL in Supabase Dashboard → SQL Editor
-- ==================================================================

-- 1. LAR Sources table - tracks each uploaded LAR database/PDF
CREATE TABLE IF NOT EXISTS lar_sources (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  source_key TEXT NOT NULL UNIQUE,         -- e.g. "er_lar_2023", "nr_lar_2024"
  display_name TEXT NOT NULL,              -- e.g. "Eastern Railway LAR 2023"
  description TEXT,                        -- optional notes (e.g. Zone, Division)
  item_count INTEGER DEFAULT 0,
  uploaded_at TIMESTAMPTZ DEFAULT NOW(),
  uploaded_by TEXT,                        -- username
  is_active BOOLEAN DEFAULT TRUE
);

-- 2. LAR Items table - all individual items across all LAR sources
CREATE TABLE IF NOT EXISTS lar_items (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  source_id UUID NOT NULL REFERENCES lar_sources(id) ON DELETE CASCADE,
  item_code TEXT NOT NULL,                 -- e.g. "011010"
  description TEXT NOT NULL,
  unit TEXT NOT NULL,                      -- e.g. "Cum", "Rkm"
  rate NUMERIC(15, 2) NOT NULL DEFAULT 0,
  contract_ref TEXT,                       -- unique to LAR: Contract/CA Number
  loi_date DATE,                          -- Date of Letter of Intent
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Indexes for fast searching
CREATE INDEX IF NOT EXISTS idx_lar_items_source ON lar_items(source_id);
CREATE INDEX IF NOT EXISTS idx_lar_items_code ON lar_items(item_code);
CREATE INDEX IF NOT EXISTS idx_lar_items_search ON lar_items USING gin(to_tsvector('english', description));

-- 4. Enable Row Level Security (allow public read, authenticated write)
ALTER TABLE lar_sources ENABLE ROW LEVEL SECURITY;
ALTER TABLE lar_items ENABLE ROW LEVEL SECURITY;

-- Public read access
CREATE POLICY "Public can read lar_sources" ON lar_sources FOR SELECT USING (true);
CREATE POLICY "Public can read lar_items" ON lar_items FOR SELECT USING (true);

-- Authenticated can insert/update/delete
CREATE POLICY "Authenticated can manage lar_sources" ON lar_sources
  FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "Authenticated can manage lar_items" ON lar_items
  FOR ALL USING (true) WITH CHECK (true);
