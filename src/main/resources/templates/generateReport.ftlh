<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Генерация отчета</title>
    <link rel="stylesheet" type="text/css" href="/css/main.css">
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f9;
            text-align: center;
            padding: 50px;
        }

        h1 {
            color: #333;
        }

        .button-container {
            margin-top: 50px;
        }

        .button {
            display: inline-block;
            padding: 15px 30px;
            font-size: 18px;
            font-weight: bold;
            color: #fff;
            background: linear-gradient(45deg, #6a11cb, #2575fc);
            border: none;
            border-radius: 30px;
            text-decoration: none;
            cursor: pointer;
            transition: all 0.3s ease-in-out;
        }

        .button:hover {
            transform: scale(1.05);
        }

        .message {
            margin-top: 20px;
            font-size: 16px;
            color: #777;
        }

        /* Пульсация кнопки */
        @keyframes pulse {
            0% { box-shadow: 0 0 10px rgba(106, 17, 203, 0.8); }
            50% { box-shadow: 0 0 20px rgba(106, 17, 203, 0.5); }
            100% { box-shadow: 0 0 10px rgba(106, 17, 203, 0.8); }
        }

        .active-button {
            animation: pulse 1.5s infinite;
        }
    </style>
</head>
<body>
<h1>Генерация отчета завершена!</h1>
<div class="button-container">
    <a href="/reports" class="button active-button">Перейти к отчету</a>
    <p class="message">Ваш отчет готов. Нажмите на кнопку, чтобы посмотреть его.</p>
</div>

<script>

</script>
</body>
</html>
