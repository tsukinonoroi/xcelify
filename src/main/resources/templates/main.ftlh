
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
            <label for="file-upload" class="file-upload-label">
                Перетащите файл сюда или нажмите для выбора
            </label>
            <input id="file-upload" type="file" name="report" accept=".xlsx, .xls" required>
            <input type="submit" value="Отправить">
        </form>
    </div>
</div>

</body>
</html>
