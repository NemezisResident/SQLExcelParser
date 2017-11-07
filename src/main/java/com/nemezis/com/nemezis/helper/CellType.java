package com.nemezis.com.nemezis.helper;

/**
 * Created by Nemezis on 07.11.2017.
 *
 *  Класс для представления столбцов таблицы Excel
 *  с параметрами
 *
 *  number -  номер столбца
 *  type  - тип данных в этом столбце
 *
 */


public class CellType {

    int number;
    String type;

    public CellType(int number, String type) {
        this.number = number;
        this.type = type;
    }

    public int getNumber() {
        return number;
    }

    public void setNumber(int number) {
        this.number = number;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }
}
