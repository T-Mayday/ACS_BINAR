-- 1. �������� ���� ������
CREATE DATABASE system_access
    WITH OWNER = postgres
    ENCODING = 'UTF8'
    LC_COLLATE = 'ru_RU.UTF-8'
    LC_CTYPE = 'ru_RU.UTF-8'
    TEMPLATE template0;

-- =========================================
-- ���������� ������� ����������� ������ ��������� ��.
-- � psql ��� ������ ������ �����:
-- \c system_access
-- =========================================

-- 2. �������� ����� (���������)
CREATE ROLE admin WITH LOGIN PASSWORD 'your password';
CREATE ROLE moderator WITH LOGIN PASSWORD 'your password';
CREATE ROLE regular_user WITH LOGIN PASSWORD 'your password';

-- 3. ������ ���� �� ����������� � ��
GRANT CONNECT ON DATABASE system_access TO admin, moderator, regular_user;

-- 4. ��������� ����������
GRANT ALL PRIVILEGES ON DATABASE system_access TO admin;
/*
-- ��� ����� ���� SUPERUSER:
ALTER USER admin WITH SUPERUSER;
*/
GRANT SELECT, INSERT, UPDATE, DELETE ON ALL TABLES IN SCHEMA public TO moderator;
ALTER DEFAULT PRIVILEGES IN SCHEMA public
    GRANT SELECT, INSERT, UPDATE, DELETE ON TABLES TO moderator;

GRANT SELECT ON ALL TABLES IN SCHEMA public TO regular_user;
ALTER DEFAULT PRIVILEGES IN SCHEMA public
    GRANT SELECT ON TABLES TO regular_user;

-- =========================================
-- 5. �������� ������ departments, positions, systems, access_rights
-- =========================================
CREATE TABLE departments (
    id SERIAL PRIMARY KEY,
    name VARCHAR(255) UNIQUE NOT NULL
);

CREATE TABLE positions (
    id SERIAL PRIMARY KEY,
    name VARCHAR(255) UNIQUE NOT NULL
);

CREATE TABLE systems (
    id SERIAL PRIMARY KEY,
    name VARCHAR(255) UNIQUE NOT NULL
);

CREATE TABLE access_rights (
    id SERIAL PRIMARY KEY,
    department_id INT REFERENCES departments(id) ON DELETE CASCADE,
    position_id INT REFERENCES positions(id) ON DELETE CASCADE,
    system_id INT REFERENCES systems(id) ON DELETE CASCADE,
    user_type VARCHAR(20) CHECK (user_type IN ('���', '���������+�������')) NOT NULL
);

-- =========================================
-- 6. �������� ������ system_attributes � access_rights_attr
-- =========================================
CREATE TABLE system_attributes (
    id SERIAL PRIMARY KEY,
    system_id INT REFERENCES systems(id) ON DELETE CASCADE,
    name VARCHAR(255),
    value TEXT
);

CREATE TABLE access_rights_attr (
    id SERIAL PRIMARY KEY,
    access_rights_id INT REFERENCES access_rights(id) ON DELETE CASCADE,
    system_attribute_id INT REFERENCES system_attributes(id) ON DELETE CASCADE
);

-- =========================================
-- 7. ������� ��� ��������������� ���������� access_rights_attr ��� ������� � system_attributes
-- =========================================
CREATE OR REPLACE FUNCTION trigger_auto_fill_access_rights_attr()
RETURNS TRIGGER AS $$
BEGIN
    INSERT INTO access_rights_attr (access_rights_id, system_attribute_id)
    SELECT ar.id, NEW.id
      FROM access_rights ar
     WHERE ar.system_id = NEW.system_id;
    
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trigger_auto_fill_access_rights_attr
AFTER INSERT ON system_attributes
FOR EACH ROW
EXECUTE FUNCTION trigger_auto_fill_access_rights_attr();

-- =========================================
-- 8. ������� ��� ���������� ����� � access_rights_attr ��� ���������� system_attributes
-- =========================================
CREATE OR REPLACE FUNCTION update_access_rights_attr()
RETURNS TRIGGER AS $$
BEGIN
    UPDATE access_rights_attr 
       SET system_attribute_id = NEW.id
     WHERE system_attribute_id = OLD.id;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER update_access_rights_attr
AFTER UPDATE ON system_attributes
FOR EACH ROW
EXECUTE FUNCTION update_access_rights_attr();

-- =========================================
-- 9. �������� ������� users
-- =========================================
CREATE TABLE users (
    id BIGSERIAL PRIMARY KEY,
    encrypted_inn VARCHAR(64) UNIQUE NOT NULL,
    full_name VARCHAR(255),
    mail VARCHAR(255),
    phone VARCHAR(20),
    department_id INT REFERENCES departments(id) ON DELETE SET NULL,
    position_id INT REFERENCES positions(id) ON DELETE SET NULL,
    active BOOLEAN DEFAULT FALSE,
    tg_id BIGINT UNIQUE,
    tg_status BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- ������� ��� ��������������� ���������� updated_at
CREATE OR REPLACE FUNCTION update_timestamp()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER set_timestamp
BEFORE UPDATE ON users
FOR EACH ROW
EXECUTE FUNCTION update_timestamp();

-- =========================================
-- 10. �������� ������� open_in_system
-- =========================================
CREATE TABLE open_in_system(
    id BIGSERIAL PRIMARY KEY,
    user_id INT REFERENCES users(id) ON DELETE CASCADE,
    username VARCHAR(255) NOT NULL,
    password VARCHAR(128) NOT NULL,
    system_id INT REFERENCES systems(id) ON DELETE CASCADE,
    status BOOLEAN DEFAULT FALSE
);

-- ������� ��� ��������������� ���������� ������� (status = False),
-- ���� � users.active ������������ ���������� ����������
CREATE OR REPLACE FUNCTION deactivate_user_access()
RETURNS TRIGGER AS $$
BEGIN
    IF NEW.active = FALSE THEN
        UPDATE open_in_system
           SET status = FALSE
         WHERE user_id = NEW.id;
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trigger_deactivate_access
AFTER UPDATE OF active ON users
FOR EACH ROW
WHEN (OLD.active = TRUE AND NEW.active = FALSE)
EXECUTE FUNCTION deactivate_user_access();

-- =========================================
-- 11. �������� ������ INTERFACE_logs, MODUL_logs � ��������� logs
-- =========================================

CREATE TABLE INTERFACE_logs (
    id SERIAL PRIMARY KEY,
    table_name VARCHAR(255) NOT NULL,
    record_id INT NOT NULL
);

CREATE TABLE MODUL_logs (
    id SERIAL PRIMARY KEY,
    data JSONB NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

/*
  �������� ��������, ��� � PostgreSQL 
  ���� ������� �� ����� �������� ��������� �� ��� ������ �������. 
  ����� � ��� ��� ������� ����� �� record_id � INTERFACE_logs � MODUL_logs ������������.
  ������������ ���������� ��� �� ����� �������� (���� ������� � ��� �����).
  ������ �������� ���������� ������ (��������, interface_id � modul_id) 
  ��� ������� CHECK + ����������.

  �� ���������� ������ ��������� "��� ����", ��������� �� ��� ������ �������� ������.
*/

CREATE TABLE logs (
    id SERIAL PRIMARY KEY,
    location VARCHAR(10) CHECK (location IN ('INTERFACE', 'MODUL')),
    record_id INT NOT NULL,
    action VARCHAR(10) CHECK (action IN ('CREATE', 'UPDATE', 'DELETE')),
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    old_value TEXT,
    new_value TEXT,

    CONSTRAINT fk_logs_interface
        FOREIGN KEY (record_id) REFERENCES INTERFACE_logs(id) ON DELETE CASCADE,
    CONSTRAINT fk_logs_modul
        FOREIGN KEY (record_id) REFERENCES MODUL_logs(id) ON DELETE CASCADE
);

-- =========================================
-- 12. �������� ������� queue
-- =========================================
CREATE TABLE queue (
    id SERIAL PRIMARY KEY,
    data JSONB NOT NULL,
    status VARCHAR(20) CHECK (status IN ('pending', 'processing', 'completed', 'failed')) NOT NULL DEFAULT 'pending',
    attempts INT DEFAULT 0 NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
    last_attempt TIMESTAMP
);

-- =========================================
-- 13. ������� ��� ��������������� ���������� ������� ���������
-- =========================================
CREATE OR REPLACE FUNCTION update_attempts_and_status()
RETURNS TRIGGER AS $$
BEGIN
    -- ����������� ���������� �������
    NEW.attempts = OLD.attempts + 1;
    NEW.last_attempt = CURRENT_TIMESTAMP;

    -- ���� ������� ������ 5, ��������� ������ � 'failed'
    IF NEW.attempts > 5 THEN
        NEW.status = 'failed';
    END IF;

    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trigger_update_attempts
BEFORE UPDATE ON queue
FOR EACH ROW
WHEN (OLD.status = 'pending' OR OLD.status = 'processing')
EXECUTE FUNCTION update_attempts_and_status();

-- =========================================
-- 14. ������� ��� �������� ����������� ����� (completed) ������ 30 ����
-- =========================================
CREATE OR REPLACE FUNCTION delete_old_completed_tasks()
RETURNS VOID AS $$
BEGIN
    DELETE FROM queue
     WHERE status = 'completed'
       AND created_at < NOW() - INTERVAL '30 days';
END;
$$ LANGUAGE plpgsql;

-- ��������, � ����� ����� ������������ �������� (����� job/cron).
