package com.nemezis.com.nemezis.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;

/**
 * Created by Nemezis on 07.11.2017.
 * <p>
 * Универсальные методы для работы с Excel и БД
 */


public abstract class ManagerExcelSQL {


    // Метод для заливки данных из Excel на сервер MSSQL
    // From Excel to MSSQL
    public static void parseFromExcel(Connection con, String tableSQL, String query, String fileName, List<CellType> mas) throws ParseException {

        // Номер строки для обхода
        int rowNum = 1;

        try {
            // Ограничение пакета в 50000 строк
            boolean is50000 = false;

            // Строки excel файла
            Row row;

            // Очищаем таблицу в которую будем заливать данные
            Statement st = con.createStatement();
            st.executeUpdate(" DELETE " + tableSQL);
            st.close();

            // Формируем динамический запрос
            PreparedStatement ps = con.prepareStatement(query);

            // Открываем Excel файл
            FileInputStream fis = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            // Получаем 1 лист в книге
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Парсим 1 лист
            while (rowIterator.hasNext()) {
                is50000 = false;

                // Получаем строку
                row = rowIterator.next();

                // Приводим отдельные клетки к текстовому формату
                for (int i = 0; i < mas.size(); i++) {
                    if (mas.get(i).getType().equals("String")) {
                        try {
                            row.getCell(mas.get(i).getNumber()).setCellType(Cell.CELL_TYPE_STRING);
                        } catch (NullPointerException exc) {
                            exc.printStackTrace();
                        }
                    }
                }

                // Отбрасываем заголовок
                if (rowNum == 1) {
                } else {

                    // Обрабатываем каждую строку
                    for (int i = 0; i < mas.size(); i++) {

                        // Индекс для формирования пакета
                        int ind = i + 1;

                        // Если тип int
                        if (mas.get(i).getType().equals("int")) {
                            int value = 0;
                            try {
                                // Пробуем получить число из числовой ячейки
                                value = (int) row.getCell(mas.get(i).getNumber()).getNumericCellValue();
                            } catch (IllegalStateException ex) {
                                // Пробуем получить число из текстовой ячейки
                                String valueS = row.getCell(mas.get(i).getNumber()).getStringCellValue();
                                value = Integer.parseInt(valueS);
                            }
                            ps.setInt(ind, value);
                        }

                        // Если тип String
                        else if (mas.get(i).getType().equals("String")) {
                            ps.setString(ind, row.getCell(mas.get(i).getNumber()).getStringCellValue());
                        }

                        // Если тип Date
                        else if (mas.get(i).getType().equals("Date")) {

                            Date datA = null;
                            try {
                                datA = (Date) row.getCell(mas.get(i).getNumber()).getDateCellValue();
                            } catch (IllegalStateException exc) {

                                //Эксельные ячейки с датами не всегда воспринимаются как даты,
                                //поэтому кастим из строки
                                String str = row.getCell(mas.get(i).getNumber()).getStringCellValue();
                                DateFormat format = new SimpleDateFormat("dd.MM.yyyy", Locale.ENGLISH);

                                datA = format.parse(str);

                            } catch (Exception exc) {
                                datA = null;
                            }
                            ps.setDate(ind, new java.sql.Date(datA.getTime()));
                        }

                        // Если тип Float
                        else if (mas.get(i).getType().equals("Float")) {
                            float value = 0;
                            try {
                                value = (float) row.getCell(mas.get(i).getNumber()).getNumericCellValue();
                            } catch (IllegalStateException exc) {
                                // Пробуем получить число из текстовой ячейки
                                String valueS = row.getCell(mas.get(i).getNumber()).getStringCellValue();
                                value = Float.parseFloat(valueS);
                            }
                            ps.setFloat(ind, value);
                        }
                    }

                    // Добавляем в пакет
                    ps.addBatch();

                    // Если накопилось 50000 записей исполняем запись пакета
                    if (rowNum % 50000 == 0) {
                        // Исполняем пакет
                        ps.executeBatch();
                        ps.clearBatch();
                        is50000 = true;
                    }
                    // else
                }
                rowNum++;
                // while
            }

            // Исполняем пакет
            if (!is50000) {
                ps.executeBatch();
            }

            // Закрываем ресурсы
            ps.close();
            fis.close();

            System.out.println("Данные записанны на сервер!");
            System.out.println("The data is recorded, success!");

        } catch (Exception exc) {
            exc.printStackTrace();
        }
    }


    // Метод для динамического заполнения листа
    public static void fillSheet(Connection con, Sheet spreadsheet, String query, CellStyle cs) throws SQLException {

        // Подучаем набор
        PreparedStatement pst;
        pst = con.prepareStatement(query);
        ResultSet rs = pst.executeQuery();

        // Получаем метадату результирующего набора
        ResultSetMetaData rsmd = rs.getMetaData();
        int columnCount = rsmd.getColumnCount();

        // Информация о столбцах таблицы
        ArrayList<TableEntity> data = new ArrayList<TableEntity>();

        // Номер строки
        int row_num = 1;

        // Формируем заголовки
        Row row0 = spreadsheet.createRow(0);

        // Забираем названия столбцов и типы данных
        for (int i = 1; i <= columnCount; i++) {

            String name = rsmd.getColumnName(i);
            String type = rsmd.getColumnTypeName(i);
            int typel = rsmd.getColumnType(i);

            data.add(new TableEntity(name, type, typel));
            // Заголовки
            Cell cell = row0.createCell(i - 1);
            cell.setCellValue(name);
        }

        // Проходим ResultSet
        while (rs.next()) {
            // Новая строка
            Row rowR = spreadsheet.createRow(row_num);

            // Обходим строку по столбцам
            for (int i = 0; i < data.size(); i++) {

                // Если int или numeric
                if (data.get(i).getColumnTypeI() == 4 || data.get(i).getColumnTypeI() == 2) {
                    Cell cell = rowR.createCell(i);
                    cell.setCellValue(rs.getInt(i + 1));
                }

                // Если string или varchar
                else if (data.get(i).getColumnTypeI() == -9 || data.get(i).getColumnTypeI() == 12) {
                    Cell cell = rowR.createCell(i);
                    cell.setCellValue(rs.getString(i + 1));
                }

                // Если float
                else if (data.get(i).getColumnTypeI() == 8) {
                    Cell cell = rowR.createCell(i);
                    cell.setCellValue(rs.getFloat(i + 1));
                }

                // Если date
                else if (data.get(i).getColumnTypeI() == 91) {
                    Cell cell = rowR.createCell(i);
                    cell.setCellStyle(cs);
                    cell.setCellValue(rs.getDate(i + 1));
                }
            }
            row_num++;
        }

        // Закрываем ресурсы
        rs.close();
        pst.close();

        System.out.println("Файл сохранен!");
        System.out.println("File saved, success!");
    }
}
