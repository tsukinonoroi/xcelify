<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Генерация отчета</title>
    <link rel="stylesheet" type="text/css" href="/css/enter_costs.css">
    <style>
        /* Стиль для анимации */
        .loader {
            border: 8px solid #f3f3f3;
            border-top: 8px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 2s linear infinite;
            margin: 20px auto;
        }

        /* Анимация вращения */
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .hidden {
            display: none;
        }

        .content {
            text-align: center;
            margin-top: 20px;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Генерация отчета</h1>

    <div class="content">
        <!-- Показываем анимацию до тех пор, пока отчет не будет готов -->
        <div id="loader" class="loader"></div>

        <!-- Это сообщение будет показываться после завершения генерации -->
        <div id="completedMessage" class="hidden">
            <p>Отчет успешно сгенерирован!</p>
            <a href="/some-link">Перейти к отчету</a>
        </div>
    </div>

</div>

<script>
    // Функция для отображения анимации и скрытия содержимого
    function showLoader() {
        // Показываем анимацию
        document.getElementById('loader').classList.remove('hidden');
        // Скрываем сообщение о завершении генерации
        document.getElementById('completedMessage').classList.add('hidden');
    }

    // Эта функция будет вызываться, когда отчет сгенерирован
    function reportGenerated() {
        // Скрываем анимацию
        document.getElementById('loader').classList.add('hidden');
        // Показываем сообщение о завершении
        document.getElementById('completedMessage').classList.remove('hidden');
    }

    // Симуляция генерации отчета через 5 секунд для демонстрации
    setTimeout(reportGenerated, 5000);  // Здесь можно поставить реальную логику
</script>

</body>
</html>