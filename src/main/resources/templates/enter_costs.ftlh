<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Введите себестоимость</title>
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
    <h1>Введите себестоимость для продуктов</h1>
    <form action="/submit-costs" method="post">
        <table>
            <thead>
            <tr>
                <th>Продукт</th>
                <th>Себестоимость</th>
            </tr>
            </thead>
            <tbody>
            <#list uniqueProducts as product>
                <tr>
                    <td>${product}</td>
                    <td>
                        <input type="number" name="costs[${product}]" step="0.01" required>
                    </td>
                </tr>
            </#list>
            </tbody>
        </table>
        <input type="submit" value="Сохранить себестоимость">
    </form>
</div>

</body>
</html>
