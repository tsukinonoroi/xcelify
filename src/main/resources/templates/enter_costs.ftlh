<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Введите себестоимость</title>
    <link rel="stylesheet" type="text/css" href="/css/enter_costs.css">
</head>
<body>
<div class="container">
    <h1>Введите себестоимость для продуктов</h1>

    <#if products?has_content>
        <form action="/updateCosts" method="post">
            <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}"/>
            <table>
                <thead>
                <tr>
                    <th>Название</th>
                    <th>Артикул поставщика</th>
                    <th>Себестоимость</th>
                </tr>
                </thead>
                <tbody>
                <#list products as product>
                <tr>
                    <td>${product.name}</td>
                    <td>${product.articul}</td>
                    <td>
                        <input type="hidden" name="productId" value="${product.id}"> <!-- Добавлено скрытое поле для productId -->
                        <input type="number" name="costs[${product.id}]" step="0.01" value="${product.costs?default('')}" required>
                    </td>
                </tr>
                </</#list>
                </tbody>
            </table>
            <input type="submit" value="Сохранить себестоимости">
        </form>
    </#if>
</div>
</body>
</html>