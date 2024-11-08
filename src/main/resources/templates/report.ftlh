<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Список отчетов</title>
</head>
<body>
<h1>Все доступные отчеты</h1>
<ul>
    <#list reports as report>
        <li>
            <a href="/downloadReport?file=${report.name}">${report.name}</a>
        </li>
    </#list>
</ul>
</body>
</html>