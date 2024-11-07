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
                <tr>
                    <th>ID</th>
                    <th>Название</th>
                    <th>Себестоимость</th>
                </tr>
                <#list products as product>
                    <tr>
                        <td>${product.id}</td>
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
    </#if>
</div>
</body>
</html>
