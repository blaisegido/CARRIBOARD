-- 1. Table des utilisateurs
CREATE TABLE IF NOT EXISTS users (
  id BIGINT PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
  username TEXT NOT NULL,
  username_canon TEXT NOT NULL UNIQUE,
  role TEXT NOT NULL DEFAULT 'user',
  password_salt TEXT NOT NULL,
  password_hash TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  last_login TIMESTAMPTZ
);

-- 2. Table des jetons révoqués
CREATE TABLE IF NOT EXISTS revoked_tokens (
  token_hash TEXT PRIMARY KEY,
  revoked_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  exp BIGINT
);

-- 3. Table des projets
CREATE TABLE IF NOT EXISTS projects (
  id TEXT PRIMARY KEY,
  user_id BIGINT NOT NULL,
  name TEXT NOT NULL,
  name_canon TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  data_path TEXT NOT NULL,
  source_filename TEXT,
  date_min TEXT,
  date_max TEXT,
  nb_livraisons INTEGER,
  tonnage_total FLOAT,
  ca_total FLOAT,
  theme_idx INTEGER NOT NULL DEFAULT 0,
  order_idx INTEGER NOT NULL DEFAULT 0
);

-- Index pour les performances
CREATE INDEX IF NOT EXISTS idx_users_username_canon ON users(username_canon);
CREATE INDEX IF NOT EXISTS idx_revoked_tokens_exp ON revoked_tokens(exp);
CREATE INDEX IF NOT EXISTS idx_projects_user_id ON projects(user_id);
CREATE INDEX IF NOT EXISTS idx_projects_user_order ON projects(user_id, order_idx);
CREATE INDEX IF NOT EXISTS idx_projects_user_name ON projects(user_id, name_canon);
