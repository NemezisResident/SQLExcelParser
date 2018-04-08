В данном проекте реализован универсальный парсер для передачи данных с сервера MSSQL в файлы Excel и обратно. Методы не полностью автоматические и требуют на вход указания некоторых параметров. 

1) parseFromExcel (Метод загрузки даных из Excel на сервер MSSQL ) на вход подаются следующие параметры:
 - Connection - соединение с MSSQL сервером;
 - tableSQL  - название целевой таблицы на сервере;
 - query - строка вставки по шаблону INSERT INTO table (columnName) values (?);
 - fileName - путь к файлу источнику Excel;
 - typeList - массив перечисление столбцов и их типы данных;
 
 2) fillSheet (Метод загрузки данных из таблицы на сервере MSSQL в файл Excel)
 - targetPath - путь для сохранения файла Excel;
 - execQuery - запрос для выборки из таблицы;
 - sheet - лист для заполнения данными;
 - CellStyle - формат для дат;
 
 Для создания  sheet требуется создать книгу workbook, и затем вызвать метод fillSheet для заполнения листа данными, затем необходимо сохранить книгу на диске. 

