INSERT INTO access_rights (department_id, position_id, system_id, user_type)
VALUES 
    ((SELECT id FROM departments WHERE name = 'ИТ ЛАБОРАТОРИЯ'),
     (SELECT id FROM positions WHERE name = 'Техник-программист'),
     8,
     'ФИО')
RETURNING id;

INSERT INTO access_rights_attr (access_rights_id, system_attribute_id)
VALUES (
    (SELECT id FROM access_rights 
     WHERE department_id = (SELECT id FROM departments WHERE name = 'ИТ ЛАБОРАТОРИЯ')
     AND position_id = (SELECT id FROM positions WHERE name = 'Техник-программист')
     AND system_id = 8
     ORDER BY id DESC LIMIT 1),
    1
)
RETURNING id;


INSERT INTO system_attributes (system_id, name, value)
VALUES 
    (8, 'role', 'admin'),
    (8, 'role', 'moderator'),
    (8, 'role', 'user')
RETURNING id;


INSERT INTO users (encrypted_inn, full_name, mail, phone, department_id, position_id, active, tg_id, tg_status, created_at, updated_at)
VALUES 
('encrypted_value', 'Admin User', 'admin@example.com', '1234567890', 6, 18, TRUE, NULL, NULL, NOW(), NOW())
RETURNING id;


INSERT INTO open_in_system (user_id, username, password, system_id, status)
VALUES (
    (SELECT id FROM users WHERE department_id = 6 AND position_id = 18 ORDER BY created_at DESC LIMIT 1),
    'admin',
    '09f8985ae60f3dfd30402c9c48da8809:c54051a540a528c6069fa19adad83f96',
    8,
    TRUE
)
RETURNING id;
