<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Xcelify - Документация</title>
    <link rel="stylesheet" type="text/css" href="/css/main.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap" rel="stylesheet">
    <style>

        .section-title {
            font-size: 1.2em;
            font-weight: bold;
            color: #2E8B57;
            cursor: pointer;
            padding: 15px 20px;
            background: linear-gradient(145deg, #f0f0f0, #e8e8e8);
            border: 1px solid #dcdcdc;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: all 0.3s ease;
        }

        .section-title:hover {
            background: linear-gradient(145deg, #e8f5e8, #f0f0f0);
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.15);
            transform: translateY(-2px);
        }

        .section-title::after {
            content: '▼';
            font-size: 0.8em;
            color: #2E8B57;
            transition: transform 0.3s ease;
        }

        .section-title.active::after {
            transform: rotate(-180deg);
        }

        .section-content {
            display: none;
            padding: 20px;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            background: #ffffff;
            box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.05);
            margin-top: 10px;
            animation: fadeIn 0.5s ease;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .section-content p {
            margin-bottom: 15px;
            text-align: justify;
        }

        .section-content ul, .section-content ol {
            padding-left: 20px;
            margin-bottom: 15px;
        }

        .section-content li {
            margin-bottom: 8px;
        }

        .section-content strong {
            color: #2E8B57;
            font-weight: 700;
        }

        .important-note {
            background-color: #fff8e1;
            border-left: 4px solid #ffc107;
            padding: 10px 15px;
            margin-bottom: 15px;
            border-radius: 4px;
        }

        .action-button {
            display: inline-block;
            background-color: #2E8B57;
            color: white;
            padding: 8px 15px;
            border-radius: 4px;
            text-decoration: none;
            transition: background-color 0.3s ease;
        }

        .action-button:hover {
            background-color: #3cb371;
        }

        hr {
            border: 1px solid #6bb067; /* Толщина и цвет линии */
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2); /* Тень для объема */
            margin: 20px 0; /* Отступы сверху и снизу */
        }

        li {
            list-style-type: disc;
            margin-left: 20px;
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

<!-- Основной контейнер -->
<div class="container">
    <h1>Документация пользователя</h1>
    <hr>

    <!-- Пункт 1 -->
    <div class="section">
        <div class="section-title">
            1. Фильтрация еженедельных и месячных отчётов
        </div>
        <div class="section-content">
            <ol>
                <li><h2>Еженедельные отчёты:</h2>
                    <ul>
                        <li>Загрузите два файла отчетов:
                            <ul>
                                <li>Отчет по зарубежным продажам.</li>
                                <li>Отчет по продажам в России.</li>
                            </ul>
                        </li>
                        <li>После загрузки файлов внесите себестоимость для <strong>каждого продукта</strong>.</li>
                        <li class="important-note"><strong>Важно!</strong> Нельзя оставлять себестоимость пустой.</li>
                        <li>Нажмите кнопку <a href="#" class="action-button">Сохранить себестоимости</a>.</li>
                        <li>После сохранения нажмите кнопку <a href="#" class="action-button">Сгенерировать отчет</a>.</li>
                    </ul>
                </li>
                <li><h2>Месячные отчёты:</h2>
                    <ul>
                        <li>Введите месяц, за который нужно сгенерировать отчет.</li>
                        <li>Загрузите четыре еженедельных отчета для выбранного месяца:</li>
                        <ul>
                            <li>Неделя 1</li>
                            <li>Неделя 2</li>
                            <li>Неделя 3</li>
                            <li>Неделя 4</li>
                        </ul>
                        <li>Нажмите кнопку <a href="#" class="action-button">Сгенерировать месячный отчет</a>.</li>
                    </ul>
                </li>
            </ol>
        </div>
    </div>

    <!-- Пункт 2 -->
    <div class="section">
        <div class="section-title">
            2. Мои продукты
        </div>
        <div class="section-content">
            <p>На странице "Мои продукты" хранятся все товары, которые были обнаружены в загруженных отчетах.</p>
            <p>Возможные действия с продуктами:</p>
            <ul>
                <li><strong>Обновить себестоимость товара</strong> - позволяет изменить текущую себестоимость продукта.</li>
                <li><strong>Удалить товар из списка</strong> - удаляет продукт из вашей базы данных.</li>
            </ul>
        </div>
    </div>

    <!-- Пункт 3 -->
    <div class="section">
        <div class="section-title">
            3. Мои отчеты
        </div>
        <div class="section-content">
            <p>На странице "Мои отчеты" можно просмотреть все сгенерированные отчеты:</p>
            <ul>
                <li><strong>Месячный отчет</strong> - содержит данные о продажах за последний месяц.</li>
                <li><strong>Еженедельный отчет</strong> - предоставляет информацию о продажах за последнюю неделю.</li>
            </ul>
            <p>Доступные действия:</p>
            <ul>
                <li><strong>Скачать отчет</strong> - позволяет сохранить отчет на ваше устройство.</li>
                <li><strong>Удалить отчет</strong> - удаляет отчет из системы.</li>
            </ul>
            <p class="important-note"><strong>Важно!</strong> После удаления отчет будет удален с сервера, и доступ к нему восстановить невозможно.</p>
        </div>
    </div>

    <!-- Пункт 4 -->
    <div class="section">
        <div class="section-title">
            4. Настройки
        </div>
        <div class="section-content">
            <p>В разделе "Настройки" пользователь может изменить свои личные данные:</p>
            <ul>
                <li><strong>Изменить логин</strong> - позволяет обновить ваше имя пользователя.</li>
                <li><strong>Изменить пароль</strong> - дает возможность установить новый пароль для вашей учетной записи.</li>
            </ul>
            <p class="important-note"><strong>Совет:</strong> Регулярно обновляйте ваш пароль для обеспечения безопасности вашего аккаунта.</p>
        </div>
    </div>
</div>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        const sections = document.querySelectorAll('.section');

        sections.forEach(section => {
            const title = section.querySelector('.section-title');
            const content = section.querySelector('.section-content');

            title.addEventListener('click', () => {
                content.style.display = content.style.display === 'block' ? 'none' : 'block';
                title.classList.toggle('active');
            });
        });
    });
</script>

</body>
</html>
