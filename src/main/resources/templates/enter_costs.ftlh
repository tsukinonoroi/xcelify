<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Введите себестоимость</title>
    <link rel="stylesheet" type="text/css" href="/css/main.css">
</head>
<body>
<div class="container">
    <h1>Введите себестоимость для продуктов</h1>
    <style>
        button[type="submit"] {
            padding: 12px 20px;
            font-size: 1.1em;
            font-weight: bold;
            color: #ffffff;
            background: linear-gradient(135deg, #2E8B57, #3CB371); /* Градиент для стильного эффекта */
            border: none;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.15); /* Тень для объема */
        }

        button[type="submit"]:hover {
            background: linear-gradient(135deg, #3CB371, #2E8B57); /* Инверсия градиента при наведении */
            transform: translateY(-3px); /* Легкий подъем */
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.25); /* Увеличенная тень */
        }

        button[type="submit"]:active {
            transform: translateY(0); /* Возврат в исходное положение при нажатии */
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2); /* Уменьшенная тень */
            @keyframes gradient-animation {
                0% {
                    background-position: 0% 50%;
                }
                50% {
                    background-position: 100% 50%;
                }
                100% {
                    background-position: 0% 50%;
                }
            }
        }
        label[for="reportName"] {
            display: inline-block; /* Для правильного отображения градиента */
            font-size: 1.2em;
            font-weight: bold;
            background: linear-gradient(45deg, #2E8B57, #3CB371, #20B2AA, #8FBC8F);
            background-size: 200% 200%; /* Увеличенный размер фона для плавного эффекта */
            -webkit-background-clip: text; /* Обрезка по тексту */
            -webkit-text-fill-color: transparent; /* Прозрачный текст */
            animation: gradient-animation 3s infinite; /* Анимация с бесконечным циклом */
            transition: transform 0.3s ease;
        }

        label[for="reportName"]:hover {
            transform: scale(1.1); /* Увеличение текста при наведении */
        }
    </style>
    <#if products?has_content>
        <form action="/updateCosts" method="post">
            <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}"/>
            <table>
                <tr>
                    <th>Артикул</th>
                    <th>Название</th>
                    <th>Себестоимость</th>
                </tr>
                <#list products as product>
                    <tr>
                        <td>${product.articul}</td>
                        <td>${product.name}</td>
                        <td>
                            <input type="text" name="costs[${product.id}]"
                                   value="${product.cost?if_exists}" />
                        </td>
                    </tr>
                </#list>
            </table>
            <input type="submit" value="Сохранить себестоимости">
        </form>

        <form action="/generateReport" method="post" style="margin-top: 20px;">
            <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}"/>
            <label for="reportName">Как назовем отчет?</label>
            <input type="text" id="reportName" name="reportName" required/>
            <button type="submit">Сгенерировать отчет</button>
        </form>
    </#if>
</div>
</body>
</html>