<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Мои продукты</title>
    <link rel="stylesheet" type="text/css" href="/css/main.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css">
    <style>
        hr {
            border: 1px solid #6bb067;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            margin: 20px 0;
        }

        .submit-container {
            display: flex;
            justify-content: center;
            align-items: center;
            margin-bottom: 10px;
        }

        .delete-button {
            display: flex;
            margin-right: 35px;
        }
        .fa-trash-alt {
            margin-right: 10px;
        }

    </style>
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

<div class="container">
    <h1>Себестоимости товаров</h1>

    <hr>

    <form action="/products/update" method="post" id="updateForm">
        <input type="hidden" name="${_csrf.parameterName}" value="${_csrf.token}"/>
        <table>
            <tr>
                <th>Артикул</th>
                <th>Название</th>
                <th>Себестоимость</th>
                <th>Действия</th>
            </tr>
            <#list products as product>
                <tr>
                    <td>${product.articul}</td>
                    <td>${product.name}</td>
                    <td>
                        <input type="text" name="costs[${product.id}]" value="${product.cost?if_exists}">
                    </td>
                    <td>
                        <button type="button" class="delete-button" data-id="${product.id}">
                            <i class="fas fa-trash-alt"></i> Удалить
                        </button>
                    </td>
                </tr>
            </#list>
        </table>
        <div class="submit-container">
            <input type="submit" value="Сохранить себестоимости">
        </div>
    </form>
</div>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        const deleteButtons = document.querySelectorAll('.delete-button');
        const csrfToken = document.querySelector('input[name="${_csrf.parameterName}"]').value;

        deleteButtons.forEach(button => {
            button.addEventListener('click', function() {
                const productId = this.getAttribute('data-id');
                if (confirm('Вы уверены, что хотите удалить этот продукт?')) {
                    fetch('/products/delete', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/x-www-form-urlencoded',
                            '${_csrf.headerName}': csrfToken
                        },
                        body: 'id=' + productId
                    })
                        .then(response => {
                            if (response.ok) {
                                this.closest('tr').remove();
                            } else {
                                alert('Произошла ошибка при удалении продукта');
                            }
                        })
                        .catch(error => {
                            console.error('Error:', error);
                            alert('Произошла ошибка при удалении продукта');
                        });
                }
            });
        });
    });
</script>
</body>
</html>

