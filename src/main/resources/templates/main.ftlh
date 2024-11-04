
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
    <a href="/upload">Загрузить Отчёт</a>
    <a href="/history">История</a>
    <a href="/settings">Настройки</a>
</div>

<div class="container">
    <h1>Фильтрация отчетов Excel</h1>
    <p>Загрузите еженедельный отчет для получения отфильтрованного отчета.</p>

    <div class="upload-area" id="upload-area">
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}"/>
            <label for="file-upload" class="file-upload-label">
                Выберите файл
            </label>
            <input id="file-upload" type="file" name="file" accept=".xlsx, .xls" required>
            <input type="submit" value="Отправить">
        </form>
    </div>
</div>

</body>
</html>
