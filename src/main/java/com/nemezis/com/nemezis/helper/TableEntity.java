package com.nemezis.com.nemezis.helper;

/**
 * Created by Nemezis on 07.11.2017.
 * <p>
 * Класс для представления таблицы с сервера MSSQL
 * определения количества столбцов и их типы данных
 */


public class TableEntity {

    public String columnName; // Название столбца
    public String columnType; // Тип столбца
    public int columnTypel; // Тип столбца

// Конструктор

    public TableEntity() {
        super();
    }

    public TableEntity(String columnName, String columnType, int columnTypel) {
        super();
        this.columnName = columnName;
        this.columnType = columnType;
        this.columnTypel = columnTypel;

    }

    // Методы
    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getColumnType() {
        return columnType;
    }

    public void setColumnType(String columnType) {
        this.columnType = columnType;
    }

    public int getColumnTypeI() {
        return columnTypel;
    }
}