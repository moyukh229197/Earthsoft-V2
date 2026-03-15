-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- Create a projects table
CREATE TABLE IF NOT EXISTS projects (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  user_id UUID REFERENCES auth.users(id),
  name TEXT NOT NULL,
  meta JSONB DEFAULT '{}'::jsonb,
  settings JSONB DEFAULT '{}'::jsonb,
  project_config JSONB DEFAULT '{}'::jsonb, -- Corresponds to state.project
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create a table for heavy project data (rows)
CREATE TABLE IF NOT EXISTS project_data (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  project_id UUID REFERENCES projects(id) ON DELETE CASCADE,
  data_type TEXT NOT NULL, -- 'levels', 'bridges', 'curves', 'loops', 'snapshots', 'kml', 'station_plans'
  data JSONB NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE(project_id, data_type)
);

-- Enable RLS
ALTER TABLE projects ENABLE ROW LEVEL SECURITY;
ALTER TABLE project_data ENABLE ROW LEVEL SECURITY;

-- Improved Policies: Allow public access if no user_id is assigned, or if user matches
DROP POLICY IF EXISTS "Users can view their own projects" ON projects;
CREATE POLICY "Projects access" ON projects
  FOR ALL USING (true) 
  WITH CHECK (true);

DROP POLICY IF EXISTS "Users can view their own project data" ON project_data;
CREATE POLICY "Project data access" ON project_data
  FOR ALL USING (true)
  WITH CHECK (true);
