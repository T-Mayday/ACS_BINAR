
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
