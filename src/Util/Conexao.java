package Util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;

public class Conexao {

    Connection con;
    Statement st;

    public Statement getConexao(String ip, String diretorio, String usuario,String senha) {
        try {
            Class.forName("org.firebirdsql.jdbc.FBDriver");
            con = DriverManager.getConnection(
                    "jdbc:firebirdsql://" + ip + ":3050/" + diretorio,
                    usuario,
                    senha);
            st = con.createStatement();
            return st;
        } catch (Exception e) {
            return null;
        }
    }

}
