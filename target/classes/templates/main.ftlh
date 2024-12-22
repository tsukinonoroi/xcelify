<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Xcelify - Фильтрация отчетов</title>
    <link rel="stylesheet" type="text/css" href="/css/main.css">
</head>
<body>
<div class="navbar">
    <a href="/" class="logo">Xcelify</a>
    <a href="/">Главная</a>
    <a href="/reports">Мои отчёты</a>
    <a href="/products">Мои товары</a>
    <a href="/settings">Настройки</a>
    <a href="/documentation">Документация</a>
</div>

<div class="container">
    <h1>Фильтрация отчетов Excel</h1>
    <hr>
    <style>
        hr {
            border: 1px solid #6bb067; /* Толщина и цвет линии */
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2); /* Тень для объема */
            margin: 20px 0; /* Отступы сверху и снизу */
        }

        h1 {

        }
    </style>
    <!-- Секция для загрузки еженедельных отчетов -->
    <div class="section">
        <h2>Генерация еженедельного отчёта</h2>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}" />

            <p>Загрузите еженедельный отчет (Продажи по России):</p>
            <input id="file-upload-1" type="file" name="fileRussia" accept=".xlsx, .xls" required />

            <p>Загрузите еженедельный отчет (Продажи зарубежные):</p>
            <input id="file-upload-2" type="file" name="fileInternational" accept=".xlsx, .xls" required />

            <input type="submit" value="Отфильтровать отчёт" />
        </form>
    </div>

    <!-- Секция для создания месячного отчета -->
    <div class="section">
        <h2>Генерация месячного отчёта</h2>
        <form action="/monthly-report" method="post" enctype="multipart/form-data">
            <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}" />

            <p>Введите название месяца:</p>
            <input type="text" name="month" placeholder="Например: Август 2024" required />

            <p>Загрузите еженедельные отчеты:</p>
            <p>Отчёт 1:</p>
            <input id="file-weekly-1" type="file" name="fileWeekly1" accept=".xlsx, .xls" required />

            <p>Отчёт 2:</p>
            <input id="file-weekly-2" type="file" name="fileWeekly2" accept=".xlsx, .xls" required />

            <p>Отчёт 3:</p>
            <input id="file-weekly-3" type="file" name="fileWeekly3" accept=".xlsx, .xls" required />

            <p>Отчёт 4:</p>
            <input id="file-weekly-4" type="file" name="fileWeekly4" accept=".xlsx, .xls" required />

            <input type="submit" value="Отфильтровать отчёт" />
        </form>
    </div>
</div>


</body>
</html>