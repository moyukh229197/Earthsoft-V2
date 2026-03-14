-- ==================================================================
-- Earthsoft SOR Database Schema
-- Run this SQL in Supabase Dashboard → SQL Editor
-- ==================================================================

-- 1. SOR Sources table - tracks each uploaded SOR database
CREATE TABLE IF NOT EXISTS sor_sources (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  source_key TEXT NOT NULL UNIQUE,         -- e.g. "irussor_2021", "dsr_2023"
  display_name TEXT NOT NULL,              -- e.g. "IRUSSOR 2021"
  description TEXT,                        -- optional notes
  item_count INTEGER DEFAULT 0,
  uploaded_at TIMESTAMPTZ DEFAULT NOW(),
  uploaded_by TEXT,                        -- username
  is_active BOOLEAN DEFAULT TRUE
);

-- 2. SOR Items table - all individual items across all SOR sources
CREATE TABLE IF NOT EXISTS sor_items (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  source_id UUID NOT NULL REFERENCES sor_sources(id) ON DELETE CASCADE,
  item_code TEXT NOT NULL,                 -- e.g. "011010"
  description TEXT NOT NULL,
  unit TEXT NOT NULL,                      -- e.g. "Cum", "100 Sqm"
  rate NUMERIC(15, 2) NOT NULL DEFAULT 0,
  chapter TEXT,                            -- optional: chapter grouping
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Indexes for fast searching
CREATE INDEX IF NOT EXISTS idx_sor_items_source ON sor_items(source_id);
CREATE INDEX IF NOT EXISTS idx_sor_items_code ON sor_items(item_code);
CREATE INDEX IF NOT EXISTS idx_sor_items_search ON sor_items USING gin(to_tsvector('english', description));

-- 4. Enable Row Level Security (allow public read, authenticated write)
ALTER TABLE sor_sources ENABLE ROW LEVEL SECURITY;
ALTER TABLE sor_items ENABLE ROW LEVEL SECURITY;

-- Public read access
CREATE POLICY "Public can read sor_sources" ON sor_sources FOR SELECT USING (true);
CREATE POLICY "Public can read sor_items" ON sor_items FOR SELECT USING (true);

-- Authenticated can insert/update/delete
CREATE POLICY "Authenticated can manage sor_sources" ON sor_sources
  FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "Authenticated can manage sor_items" ON sor_items
  FOR ALL USING (true) WITH CHECK (true);
