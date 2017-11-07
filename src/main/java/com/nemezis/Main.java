package com.nemezis;

import com.nemezis.com.nemezis.helper.CellType;
import com.nemezis.com.nemezis.helper.ManagerExcelSQL;
import com.nemezis.sql.ConnectionSQL;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.ArrayList;

/**
 * Created by Nemezis on 07.11.2017.
 *
 *  Работа с данными
 *  Для заливки данных с файла Excel в таблицу на сервер MSSQL
 *  необходимо заполнить параметры:
 *   - Connection - соединение с MSSQL сервером;
 *   - tableSQL  - название таблицы на сервере;
 *   - query - строка вставки по шаблону INSERT INTO table (columnName) values (?);
 *   - fileName - путь к файлу источнику;
 *   - typeList - массив типов данных перечисление столбцов и их типы данных;
 *
 *  Для заливки данных из таблицы на сервере MSSQL в файла Excel
 *  - targetPath - путь для сохранения файла Excel
 *  - execQuery - запрос для выборки из таблицы
 *  - sheet - лист для заполнения данными
 *  - CellStyle - формат для дат
 *
 */


public class Main {
    public static void main(String[] args) throws SQLException, IOException {

        // Создаем подключение к Базе данных
        Connection con = ConnectionSQL.createConnection();

        // From Excel to MSSQL ------------------------------------------------------------------
        // Целевая таблица на сервере
        String tableSQL = "lusers.dbo.excel_parse_data"; // target table

        // Строка для вставки данных в целевую  таблицу
        String query = "INSERT INTO " + tableSQL + " (id, txt, dat, value) values (?,?,?,?)";

        // Файл источник данных для сервера SQL
        String fileName = "src/main/resources/test.xlsx";

        // Параметры Excel файла
        ArrayList<CellType>  typeList = new ArrayList<CellType>();
        typeList.add(new CellType(0, "int"));
        typeList.add(new CellType(1, "String"));
        typeList.add(new CellType(2, "Date"));
        typeList.add(new CellType(3, "Float"));

        // Выполняем запись данных  / Execute
        try {
            ManagerExcelSQL.parseFromExcel(con, tableSQL, query, fileName, typeList);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        //-------------------------------------------------------------------------------------------


        // From MSSQL to Excel  ------------------------------------------------------------------
        // Путь для сохранения файла с данными
        String targetPath = "C:/Users/Nemezis/Documents/result.xlsx";

        // Текст запроса для выгрузки данных из таблицы источинка
        String  execQuery = "SELECT [id] " +
                "      ,[txt] " +
                "      ,[dat] " +
                "      ,[value] " +
                "  FROM " + tableSQL;

        // Новый файл excel
        SXSSFWorkbook workbookR = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        Sheet sheet = workbookR.createSheet("Результат");  // Название листа Excel

        // Формат для Даты
        XSSFDataFormat df = (XSSFDataFormat) workbookR.createDataFormat();
        CellStyle cs = workbookR.createCellStyle(); cs.setDataFormat(df.getFormat("dd.MM.yyyy"));

        // Заполняем новый файл данными из полученного результата запроса
        ManagerExcelSQL.fillSheet(con, sheet, execQuery, cs);

        // Сохраняем файл
        FileOutputStream out = new FileOutputStream(targetPath);
        workbookR.write(out);
        out.close();

        // Очищаем временные файлы
        workbookR.dispose();

        //-------------------------------------------------------------------------------------------
        con.close();
    }
}