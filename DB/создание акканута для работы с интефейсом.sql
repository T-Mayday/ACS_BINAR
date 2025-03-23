-- 1. Добавляем access_rights
WITH ar AS (
    INSERT INTO access_rights (department_id, position_id, system_id, user_type)
    VALUES (
        (SELECT id FROM departments WHERE name = 'ИТ ЛАБОРАТОРИЯ'),
        (SELECT id FROM positions WHERE name = 'Техник-программист'),
        (SELECT id FROM systems WHERE name = 'DB'),
        'ФИО'
    )
    RETURNING id, department_id, position_id, system_id
),

-- 2. Добавляем system_attributes и возвращаем их id
attrs AS (
    INSERT INTO system_attributes (system_id, name, value)
    VALUES
        ((SELECT system_id FROM ar), 'role', 'admin'),
        ((SELECT system_id FROM ar), 'role', 'moderator'),
        ((SELECT system_id FROM ar), 'role', 'user')
    RETURNING id, value
),

-- 3. Привязываем нужный атрибут (например, "admin") к правам
ar_attr AS (
    INSERT INTO access_rights_attr (access_rights_id, system_attribute_id)
    VALUES (
        (SELECT id FROM ar),
        (SELECT id FROM attrs WHERE value = 'admin')
    )
    RETURNING id
),

-- 4. Добавляем пользователя
usr AS (
    INSERT INTO users (encrypted_inn, full_name, mail, phone, department_id, position_id, active, tg_id, tg_status, created_at, updated_at)
    VALUES (
        'encrypted_value',
        'Admin User',
        'admin@example.com',
        '1234567890',
        (SELECT department_id FROM ar),
        (SELECT position_id FROM ar),
        TRUE,
        NULL,
        NULL,
        NOW(),
        NOW()
    )
    RETURNING id
)

-- 5. Привязываем пользователя к системе
INSERT INTO open_in_system (user_id, username, password, system_id, status)
VALUES (
    (SELECT id FROM usr),
    'admin',
    '09f8985ae60f3dfd30402c9c48da8809:c54051a540a528c6069fa19adad83f96',
    (SELECT system_id FROM ar),
    TRUE
)
RETURNING id;
