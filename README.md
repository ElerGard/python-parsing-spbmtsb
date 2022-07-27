# python-parsing-spbmtsb

Это приложение сделано для скачивания файлов с сайта: https://spimex.com/markets/oil_products/trades/results/ сохранения данных из скачанных файлов в БД MsAccess,
а также для экспорта конкретных данных из БД по выбранным критериям. 

На текущий момент, работает только с файлами после 19.10.2015

## Интерфейс
<details>
<summary>Поиск элементов</summary>
Для поиска элемента введите необходимый текст в поле ввода

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/search.gif)
</details>

<details>
<summary>Добавление/удаление элементов</summary>
Выбор и удаление выбранного элемента происходит двойным щелчком лкм

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/add_del.gif)
</details>

<details>
<summary>Выбор даты</summary>

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/date.gif)
</details>

<details>
<summary>Кнопки</summary>
<details>
<summary>Кнопка сброс</summary>
Кнопка сброс удаляет все выбранные элементы и обнуляет строку поиска

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/reset.gif)
</details>

<details>
<summary>Кнопка Обновить базу данных</summary>
Кнопка обновить базу данных скачивает новые файлы с сайта и добавляет их в БД

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/db.gif)
</details>

<details>
<summary>Обновить ресурсы</summary>
Кнопка обновить ресурсы добавляет в БД новые ресурсы в из файла recources.txt

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/res.gif)
</details>
<details>
<summary>Кнопка экспорт</summary>
Делается выборка из БД по выбранным критериям и данные записываются в эксель файл с текущей датой и временем в названии

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/export.gif)
</details>
</details>

## Структура файла бд

Название таблицы: Главная

Структура таблицы:

![table](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/Table_structure.jpg)
## Библиотеки

Установка необходимых библиотек:

    pip install -r requirements.txt

## Запуск исходников

    python GUI.pyw