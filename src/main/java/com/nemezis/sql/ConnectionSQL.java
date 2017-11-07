package com.nemezis.sql;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * Created by Nemezis on 07.11.2017.
 *
 *    Класс для создания подключения к Базе данных MSSQL
 *    необходимо указать пользовательские параметры
 *
 */

public abstract class ConnectionSQL {

    // Метод получения подключения
    public static Connection createConnection() {

        // Создаем подключение
        Connection connection = null;
        // Параметры для подключения
        String url = "jdbc:sqlserver://localhost:58072;";  //  manual url
        String user = "root";   // manual username
        String password = "XXX";  // manual passwordL

        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            connection = DriverManager.getConnection(url, user, password);

        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }

        System.out.println("Соединение установленно.");
        return connection;
    }
}
