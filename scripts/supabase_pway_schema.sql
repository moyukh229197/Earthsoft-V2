-- ==================================================================
-- Earthsoft P-Way (Permanent Way) Database Schema
-- Run this SQL in Supabase Dashboard → SQL Editor
-- ==================================================================

-- 1. P-Way Records table - tracks infrastructure for each station/segment
CREATE TABLE IF NOT EXISTS pway_records (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  project_id UUID,                         -- link to project if applicable
  station_name TEXT NOT NULL,
  station_type TEXT,                       -- e.g. "B-Class", "Junction"
  chainage_start NUMERIC(15, 3),           -- optional chainage tracking
  chainage_end NUMERIC(15, 3),
  
  -- Track Lengths (meters)
  mainline_len NUMERIC(15, 3) DEFAULT 0,
  passenger_loop_len NUMERIC(15, 3) DEFAULT 0,
  goods_loop_len NUMERIC(15, 3) DEFAULT 0,
  crossover_len NUMERIC(15, 3) DEFAULT 0,
  connecting_track_len NUMERIC(15, 3) DEFAULT 0,
  sidings_len NUMERIC(15, 3) DEFAULT 0,
  osl_len NUMERIC(15, 3) DEFAULT 0,
  
  -- Turnouts (Quantities)
  turnout_8_5_qty INTEGER DEFAULT 0,
  turnout_12_qty INTEGER DEFAULT 0,
  turnout_16_qty INTEGER DEFAULT 0,
  derailing_switch_qty INTEGER DEFAULT 0,
  
  -- Ancillary (Quantities)
  sej_qty INTEGER DEFAULT 0,
  buffer_stop_qty INTEGER DEFAULT 0,
  sand_hump_qty INTEGER DEFAULT 0,
  
  -- Dismantling (Tracking removal)
  dismantle_track_len NUMERIC(15, 3) DEFAULT 0,
  dismantle_turnout_qty INTEGER DEFAULT 0,
  
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Indexes for fast searching
CREATE INDEX IF NOT EXISTS idx_pway_station ON pway_records(station_name);

-- 3. Enable Row Level Security
ALTER TABLE pway_records ENABLE ROW LEVEL SECURITY;

-- Public read access
CREATE POLICY "Public can read pway_records" ON pway_records FOR SELECT USING (true);

-- Authenticated can insert/update/delete
CREATE POLICY "Authenticated can manage pway_records" ON pway_records
  FOR ALL USING (true) WITH CHECK (true);
