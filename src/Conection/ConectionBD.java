/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Conection;

import java.sql.*;
//import java.sql.ResultSet;
import java.sql.PreparedStatement;

/**
 *
 * @author Tiago Dantas
 */
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;



/**
 *
 * @author Tiago Dantas
 */
public class ConectionBD {
        // Definido variaveis de conexão com o banco de dados
    
    public Statement stm; // Variavel responsavel pela pesquisa.
    public ResultSet rs; // Variavel responsavel por guardar as pesquisas.
    private String driver = "org.postgres.Driver"; // Variavel responsavel por guarda o endereço.
    private String caminho = "jdbc:postgresql://localhost:5432/ProjetoSIONR"; // Variavel responsavel por con
    private String usuario = "postgres";
    private String senha = "123456";
    public Connection con;
    
    
    public void conexao (){   // Metodo responsavel por realizar conexao com a base de dados.
    
        try {
            
            System.setProperty("jdbc.Drivers", driver);
            con = DriverManager.getConnection(caminho, usuario, senha);
            //JOptionPane.showMessageDialog(null, "Conexão Efetuado Com Sucesso!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Na Conexão Ao Banco de Dados: \n"+ ex.getMessage());
        }
    
}
    
    public void executaSql (String sql){
        
        try {
            stm = con.createStatement(rs.TYPE_SCROLL_INSENSITIVE, rs.CONCUR_READ_ONLY);
            rs = stm.executeQuery(sql);
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null,"Erro Ao Executar SQL!\n Erro: " +ex.getMessage());
        }
        
    }
    
    public void desconecta (){ // Metodo responsavel por realizar desconexao com a base de dados.
        
        try {
            con.close();
            //JOptionPane.showMessageDialog(null, "BD Desconectado Com Sucesso!");
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Fechar a Conexão Ao Banco de Dados: \n"+ ex.getMessage());
        }
    }

}