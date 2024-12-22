<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Мои отчеты</title>
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
<main class="reports-container">
    <h1>Мои отчёты</h1>
    <hr>
    <style>
        hr {
            border: 1px solid #6bb067;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            margin: 20px 0;
        }
    </style>
    <#list reports as report>
        <article class="report-item">

            <img src="/img/excel.png" alt="Excel" class="report-icon">

            <div class="report-details">
                <p>Название отчета: ${report.filename}</p>
                <p>Дата создания: ${report.formattedDate}</p>
            </div>

            <a href="/report/download/${report.id?replace(' ', '')}" class="download-button">Скачать</a>
            <form action="/report/delete/${report.filename}" method="post" style="display:inline;">
                <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}" />
                <button type="submit" class="delete-button">Удалить</button>
            </form>
        </article>
        <#else> <h1>Пусто!</h1>
    </#list>
</main>
</body>
</html>
