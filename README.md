# python-parsing-spbmtsb

Этот скрипт сделан для скачивания файлов с сайта: https://spimex.com/markets/oil_products/trades/results/ сохранения данных из файлов в БД MsAccess.
А также для экспорта конкретных данных из БД по критериям

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
Кнопка сброс удаляет все выбранные элементы и обнуляет строку поиска

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/db.gif)
</details>

<details>
<summary>Обновить ресурсы</summary>
Кнопка сброс удаляет все выбранные элементы и обнуляет строку поиска

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/res.gif)
</details>
<details>
<summary>Кнопка экспорт</summary>
Делается выборка из БД по выбранным критериям и данные записываются в эксель файл с текущей датой и временем в названии

![gif](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/export.gif)
</details>
</details>

## Структура файла бд

Название таблицы: Главная\
Структура таблицы:\
![table](https://github.com/ElerGard/python-parsing-spbmtsb/blob/master/demo/Table_structure.jpg)

## Библиотеки

Установка необходимых библиотек:

    pip install -r requirements.txt

## Запуск скрипта

    python GUI.pyw