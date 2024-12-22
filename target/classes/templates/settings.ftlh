<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Настройки - Xcelify</title>
    <link rel="stylesheet" type="text/css" href="/css/main.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
    <style>
        .container {
            max-width: 1200px;
            margin: 20px auto;
            padding: 20px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        hr {
            border: 1px solid #2E8B57;
            margin: 20px 0;
        }

        .settings-box {
            padding: 20px;
            background: #fafafa;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }

        .settings-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }

        .settings-title {
            font-size: 1.6em;
            color: #333;
        }

        .edit-button {
            background-color: #2E8B57;
            color: white;
            border: none;
            padding: 10px 20px;
            font-size: 1em;
            border-radius: 5px;
            cursor: pointer;
            transition: background-color 0.3s ease;
        }

        .edit-button:hover {
            background-color: #218838;
        }

        .user-info p {
            font-size: 1.2em;
            margin: 10px 0;
            color: #555;
        }

        .password-toggle {
            cursor: pointer;
            font-size: 1.4em;
            color: #888;
            margin-left: 10px;
        }

        .form-group {
            margin-bottom: 15px;
        }

        .form-group label {
            display: block;
            font-size: 1.1em;
            margin-bottom: 5px;
            color: #333;
        }

        .form-group input {
            width: 95%;
            padding: 10px;
            font-size: 1em;
            border: 1px solid #ddd;
            border-radius: 5px;
            margin-bottom: 10px;
            display: block;
        }

        .save-button, .cancel-button {
            padding: 0.6em 1.2em;
            font-size: 1em;
            font-weight: bold;
            border-radius: 5px;
            cursor: pointer;
            transition: transform 0.2s ease, background-color 0.3s ease;
            border: none;
            width: 100%;
            max-width: 300px;
            margin: 0 auto;
            display: block;
        }

        .save-button {
            background-color: #2E8B57;
            color: white;
        }

        .save-button:hover {
            background-color: #218838;
            transform: scale(1.05);
        }

        .cancel-button {
            background-color: #ff2400;
            color: white;
        }

        .cancel-button:hover {
            background-color: #c82333;
            transform: scale(1.05);
        }

        .hidden {
            display: none;
        }

        .show {
            animation: fadeIn 0.3s ease-in-out;
        }

        @keyframes fadeIn {
            0% { opacity: 0; }
            100% { opacity: 1; }
        }

    </style>

</head>
<body>

<!-- Навигационное меню -->
<div class="navbar">
    <a href="/" class="logo">Xcelify</a>
    <a href="/">Главная</a>
    <a href="/reports">Мои отчёты</a>
    <a href="/products">Мои товары</a>
    <a href="/settings">Настройки</a>
    <a href="/documentation">Документация</a>
</div>

<!-- Основной контент -->
<div class="container">
    <h1>Настройки пользователя</h1>
    <hr>
    <div class="settings-box">
        <div class="settings-header">
            <span class="settings-title">Ваши данные</span>
            <button class="edit-button" onclick="showEditForm()">Изменить</button>
        </div>
        <div class="user-info">
            <p><strong>Пароль:</strong> <span id="display-password">••••••••</span>
                <span class="password-toggle" onclick="togglePassword()">👁️</span></p>
        </div>

        <div id="edit-form" class="hidden">
            <form action="/settings/changepassword" method="post">
                <input type="hidden" name="_csrf" value="${_csrf.token}" />
                <div class="form-group">
                    <label for="new-password">Новый пароль:</label>
                    <input type="text" id="new-password" name="newPassword" required />
                </div>

                <button type="submit" class="save-button">Сохранить</button>
                <button type="button" class="cancel-button" onclick="cancelEdit()">Отмена</button>
            </form>
        </div>
    </div>
</div>

<script>
    // Получаем пароль из модели
    const currentPassword = '${user.password}'; // Используем значение пароля текущего пользователя из модели

    function togglePassword() {
        const password = document.getElementById('display-password');
        const currentText = password.textContent;
        password.textContent = currentText === '••••••••' ? currentPassword : '••••••••';
        const eyeIcon = document.querySelector('.password-toggle');
        eyeIcon.textContent = currentText === '••••••••' ? '🔒' : '👁️'; // Изменение иконки
    }

    function togglePasswordVisibility() {
        const passwordField = document.getElementById('new-password');
        passwordField.type = passwordField.type === 'password' ? 'text' : 'password';
        const eyeIcon = document.querySelector('.password-toggle');
        eyeIcon.textContent = passwordField.type === 'password' ? '👁️' : '🔒'; // Изменение иконки
    }

    function showEditForm() {
        document.getElementById('edit-form').classList.remove('hidden'); // Показываем форму
        document.querySelector('.edit-button').classList.add('hidden'); // Скрываем кнопку "Изменить"
    }

    function cancelEdit() {
        document.getElementById('edit-form').classList.add('hidden'); // Скрываем форму
        document.querySelector('.edit-button').classList.remove('hidden'); // Показываем кнопку "Изменить"
    }
</script>

</body>
</html>
