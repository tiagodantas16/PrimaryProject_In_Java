/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package EditionDao;

import Conection.ConectionBD;
import ModeloBeans.BeansExcel;
import ModeloBeans.BeansUsuários;
import Visão.FormUsuários;
import Visão.VistaExcel;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

/**
 *
 * @author Tiago Dantas
 */
public class DaoUsuario {
        ConectionBD conex = new ConectionBD();
        ModeloBeans.BeansUsuários mod = new BeansUsuários();
        ModeloBeans.BeansExcel modexcel = new BeansExcel();
    
    public void Salvar (BeansUsuários mod){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into usuarios(nom_usuario,nom_nome,nom_senha,nom_tipo) values (?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1, mod.getUsu_Usuário());
         pst.setString(2,mod.getUsu_Nome());
         pst.setString(3,mod.getUsu_Senha());
         pst.setString(4,mod.getUsu_Tipo());
         pst.execute();
         
            JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       
       conex.desconecta();
    }
        
    public void Editar(BeansUsuários mod){
        
        conex.conexao();
        
        try {
            PreparedStatement pst = conex.con.prepareStatement("update usuarios set nom_usuario=?,nom_nome=?,nom_senha=?,nom_tipo=? where id_cod=?"); // Comando usado para editar dados no banco de dados.
            // update - é um código sql que edita dados no banco de dados.
            // where - é um código sql usado para indicar referencia de localização.
            
         pst.setString(1, mod.getUsu_Usuário());
         pst.setString(2,mod.getUsu_Nome());
         pst.setString(3,mod.getUsu_Senha());
         pst.setString(4,mod.getUsu_Tipo());
         pst.setInt(5,mod.getUsu_Id());
         pst.execute();
            
            JOptionPane.showMessageDialog(null,"Dados Editados Com Sucesso!");
                    } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null,"Erro Ao Editar Dados!\nErro: "+ex);
        }
        
        conex.desconecta();
        
    }
    public void Excluir (BeansUsuários mod){
        
        conex.conexao();
        
        try {
            PreparedStatement pst = conex.con.prepareStatement("delete from usuarios where id_cod=?");
            
            pst.setInt (1, mod.getUsu_Id());
            pst.execute();
            
            JOptionPane.showMessageDialog(null,"Dados Excluidos Com Sucesso!");
                    } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null,"Erro Ao Excluir Dados!\nErro: "+ex);
        }
        
        conex.desconecta();
        
    }
    public BeansUsuários buscaUsuario (BeansUsuários mod){
        
        conex.conexao();
        
        conex.executaSql("Select *from usuarios where nom_usuario like'%" +mod.getPesquisa() +"%'");
        
        try {
            
            conex.rs.first();
            mod.setUsu_Id(conex.rs.getInt("id_cod"));
            mod.setUsu_Usuário(conex.rs.getString("nom_usuario"));
            mod.setUsu_Nome(conex.rs.getString("nom_nome"));
            mod.setUsu_Senha(conex.rs.getString("nom_senha"));
            mod.setUsu_Tipo(conex.rs.getString("nom_tipo"));
            } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Usuário Não Existe!");
        }
        
        conex.desconecta();
        return mod;
        
    }
    

    public BeansUsuários buscaTipo (BeansUsuários mod){
        
        conex.conexao();
        
        conex.executaSql("Select *from usuarios where nom_usuario like'%" +mod.getUsu_BTipo()+"%'");
        
        try {
            
            conex.rs.first();
            mod.setUsu_Tipo(conex.rs.getString("nom_tipo"));
            
            } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Usuário Não Existe!");
        }
        
        conex.desconecta();
        return mod;
    }
    public BeansUsuários PuxeBase (BeansUsuários mod){
        
       conex.conexao();
       conex.executaSql("select *from "+mod.getUsu_Puxe_01()+" where perfil like'%" +mod.getUsu_Puxe()+"%'"); 
        try {
                    //inset into - é um código sql para inserir dados
         conex.rs.first();
            mod.setPuxe_Base(conex.rs.getString("base"));
            mod.setPuxe_Nivel(conex.rs.getString("aging_cliente"));
            mod.setPuxe_Cod_Cidade(conex.rs.getString("code_operator"));
            mod.setPuxe_Contrato(conex.rs.getString("code_net"));
            mod.setPuxe_Segmento(conex.rs.getString("segmento"));
            mod.setPuxe_Cliente(conex.rs.getString("nm_mailing"));
         
           } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Usuário Não Existe!");
        }
        
        conex.desconecta();
        return mod;
    }
    public BeansUsuários buscaRegistro (BeansUsuários modUser){
        
        conex.conexao();
        
        conex.executaSql("Select *from usuarios where nom_usuario like'%" +modUser.getUsu_BTipo()+"%'");
        
        try {
            
            conex.rs.first();
            modUser.setUsu_Tipo(conex.rs.getString("nom_tipo"));
            
            } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Usuário Não Existe!");
        }
        
        conex.desconecta();
        return mod;
    }
    
    public void InsertTr (BeansUsuários mod){
        
       conex.conexao();
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into "+mod.getUsu_Puxe_02()+"(data,hora_inicio,hora_fim,usuario,base,perfil,nv_regua,cod_cidade,contrato,segmento,cliente,cpf,tabulacao) values (?,?,?,?,?,?,?,?,?,?,?,?,?)");
            
            pst.setString(1, mod.getTrat_Data());
            pst.setString(2, mod.getTrat_HrInicio());
            pst.setString(3, mod.getTrat_HrFim());
            pst.setString(4, mod.getTrat_Usuario());
            pst.setString(5, mod.getTrat_Base());
            pst.setString(6, mod.getTrat_Perfil());
            pst.setString(7, mod.getTrat_Nivel());
            pst.setString(8, mod.getTrat_Cod_Cidade());
            pst.setString(9, mod.getTrat_Contrato());
            pst.setString(10, mod.getTrat_Segmento());
            pst.setString(11, mod.getTrat_Cliente());
            pst.setString(12, mod.getTrat_CPF());
            pst.setString(13, mod.getTrat_Tabulacao());
            pst.execute();
            
         
           } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Dados Não Existe!");
        }
        
        conex.desconecta();
    }
    
     public void Importar (BeansExcel modexcel){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into finalizacao(base,perfil,aging_cliente,code_action,code_operator,code_net,code_client,code_address,nu_ddd_1,residencial,nu_ddd_2,celular,nu_ddd_3,comercial,nu_ddd_4,ura,nu_ddd_5,net_fone,nm_mailing,nm_city,de_segment_strategy,de_extra_field_1,de_extra_field_2,segmento,alto_valor,dt_desco,dt_instalacao,tipo,dt_envio,dt_atendimento) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1,modexcel.getBase_usu());
         pst.setString(2,modexcel.getPerfil_usu());
         pst.setString(3,modexcel.getAging_cliente_usu());
         pst.setString(4,modexcel.getCode_action_usu());
         pst.setString(5,modexcel.getCode_operator_usu());
         pst.setString(6,modexcel.getCode_net_usu());
         pst.setString(7,modexcel.getCode_client_usu());
         pst.setString(8,modexcel.getCode_address_usu());
         pst.setString(9,modexcel.getNu_ddd_1_usu());
         pst.setString(10,modexcel.getResidencial_usu());
         pst.setString(11,modexcel.getNu_ddd_2_usu());
         pst.setString(12,modexcel.getCelular_usu());
         pst.setString(13,modexcel.getNu_ddd_3_usu());
         pst.setString(14,modexcel.getComercial_usu());
         pst.setString(15,modexcel.getNu_ddd_4_usu());
         pst.setString(16,modexcel.getUra_usu());
         pst.setString(17,modexcel.getNu_ddd_5_usu());
         pst.setString(18,modexcel.getNet_fone_usu());
         pst.setString(19,modexcel.getNm_mailing_usu());
         pst.setString(20,modexcel.getNm_city_usu());
         pst.setString(21,modexcel.getDe_segment_strategy_usu());
         pst.setString(22,modexcel.getDe_extra_field_1_usu());
         pst.setString(23,modexcel.getDe_extra_field_2_usu());
         pst.setString(24,modexcel.getSegmento_usu());
         pst.setString(25,modexcel.getAlto_valor_usu());
         pst.setString(26,modexcel.getDt_desco_usu());
         pst.setString(27,modexcel.getDt_instalacao_usu());
         pst.setString(28,modexcel.getTipo_usu());
         pst.setString(29,modexcel.getDt_envio_usu());
         pst.setString(30,modexcel.getDt_atendimento_usu());



         pst.execute();
         
            //JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       conex.desconecta();
    }
    
     public void Importar_01 (BeansExcel modexcel){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into auditoriaos(base,perfil,aging_cliente,code_action,code_operator,code_net,code_client,code_address,nu_ddd_1,residencial,nu_ddd_2,celular,nu_ddd_3,comercial,nu_ddd_4,ura,nu_ddd_5,net_fone,nm_mailing,nm_city,de_segment_strategy,de_extra_field_1,de_extra_field_2,segmento,alto_valor,dt_desco,dt_instalacao,tipo,dt_envio,dt_atendimento) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1,modexcel.getBase_usu());
         pst.setString(2,modexcel.getPerfil_usu());
         pst.setString(3,modexcel.getAging_cliente_usu());
         pst.setString(4,modexcel.getCode_action_usu());
         pst.setString(5,modexcel.getCode_operator_usu());
         pst.setString(6,modexcel.getCode_net_usu());
         pst.setString(7,modexcel.getCode_client_usu());
         pst.setString(8,modexcel.getCode_address_usu());
         pst.setString(9,modexcel.getNu_ddd_1_usu());
         pst.setString(10,modexcel.getResidencial_usu());
         pst.setString(11,modexcel.getNu_ddd_2_usu());
         pst.setString(12,modexcel.getCelular_usu());
         pst.setString(13,modexcel.getNu_ddd_3_usu());
         pst.setString(14,modexcel.getComercial_usu());
         pst.setString(15,modexcel.getNu_ddd_4_usu());
         pst.setString(16,modexcel.getUra_usu());
         pst.setString(17,modexcel.getNu_ddd_5_usu());
         pst.setString(18,modexcel.getNet_fone_usu());
         pst.setString(19,modexcel.getNm_mailing_usu());
         pst.setString(20,modexcel.getNm_city_usu());
         pst.setString(21,modexcel.getDe_segment_strategy_usu());
         pst.setString(22,modexcel.getDe_extra_field_1_usu());
         pst.setString(23,modexcel.getDe_extra_field_2_usu());
         pst.setString(24,modexcel.getSegmento_usu());
         pst.setString(25,modexcel.getAlto_valor_usu());
         pst.setString(26,modexcel.getDt_desco_usu());
         pst.setString(27,modexcel.getDt_instalacao_usu());
         pst.setString(28,modexcel.getTipo_usu());
         pst.setString(29,modexcel.getDt_envio_usu());
         pst.setString(30,modexcel.getDt_atendimento_usu());



         pst.execute();
         
            //JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       conex.desconecta();
    }
     
     public void Importar_02 (BeansExcel modexcel){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into auditoriarevdesc(base,perfil,aging_cliente,code_action,code_operator,code_net,code_client,code_address,nu_ddd_1,residencial,nu_ddd_2,celular,nu_ddd_3,comercial,nu_ddd_4,ura,nu_ddd_5,net_fone,nm_mailing,nm_city,de_segment_strategy,de_extra_field_1,de_extra_field_2,segmento,alto_valor,dt_desco,dt_instalacao,tipo,dt_envio,dt_atendimento) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1,modexcel.getBase_usu());
         pst.setString(2,modexcel.getPerfil_usu());
         pst.setString(3,modexcel.getAging_cliente_usu());
         pst.setString(4,modexcel.getCode_action_usu());
         pst.setString(5,modexcel.getCode_operator_usu());
         pst.setString(6,modexcel.getCode_net_usu());
         pst.setString(7,modexcel.getCode_client_usu());
         pst.setString(8,modexcel.getCode_address_usu());
         pst.setString(9,modexcel.getNu_ddd_1_usu());
         pst.setString(10,modexcel.getResidencial_usu());
         pst.setString(11,modexcel.getNu_ddd_2_usu());
         pst.setString(12,modexcel.getCelular_usu());
         pst.setString(13,modexcel.getNu_ddd_3_usu());
         pst.setString(14,modexcel.getComercial_usu());
         pst.setString(15,modexcel.getNu_ddd_4_usu());
         pst.setString(16,modexcel.getUra_usu());
         pst.setString(17,modexcel.getNu_ddd_5_usu());
         pst.setString(18,modexcel.getNet_fone_usu());
         pst.setString(19,modexcel.getNm_mailing_usu());
         pst.setString(20,modexcel.getNm_city_usu());
         pst.setString(21,modexcel.getDe_segment_strategy_usu());
         pst.setString(22,modexcel.getDe_extra_field_1_usu());
         pst.setString(23,modexcel.getDe_extra_field_2_usu());
         pst.setString(24,modexcel.getSegmento_usu());
         pst.setString(25,modexcel.getAlto_valor_usu());
         pst.setString(26,modexcel.getDt_desco_usu());
         pst.setString(27,modexcel.getDt_instalacao_usu());
         pst.setString(28,modexcel.getTipo_usu());
         pst.setString(29,modexcel.getDt_envio_usu());
         pst.setString(30,modexcel.getDt_atendimento_usu());



         pst.execute();
         
            //JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       conex.desconecta();
    }
     
         public void Importar_03 (BeansExcel modexcel){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into auditoriasemlead(base,perfil,aging_cliente,code_action,code_operator,code_net,code_client,code_address,nu_ddd_1,residencial,nu_ddd_2,celular,nu_ddd_3,comercial,nu_ddd_4,ura,nu_ddd_5,net_fone,nm_mailing,nm_city,de_segment_strategy,de_extra_field_1,de_extra_field_2,segmento,alto_valor,dt_desco,dt_instalacao,tipo,dt_envio,dt_atendimento) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1,modexcel.getBase_usu());
         pst.setString(2,modexcel.getPerfil_usu());
         pst.setString(3,modexcel.getAging_cliente_usu());
         pst.setString(4,modexcel.getCode_action_usu());
         pst.setString(5,modexcel.getCode_operator_usu());
         pst.setString(6,modexcel.getCode_net_usu());
         pst.setString(7,modexcel.getCode_client_usu());
         pst.setString(8,modexcel.getCode_address_usu());
         pst.setString(9,modexcel.getNu_ddd_1_usu());
         pst.setString(10,modexcel.getResidencial_usu());
         pst.setString(11,modexcel.getNu_ddd_2_usu());
         pst.setString(12,modexcel.getCelular_usu());
         pst.setString(13,modexcel.getNu_ddd_3_usu());
         pst.setString(14,modexcel.getComercial_usu());
         pst.setString(15,modexcel.getNu_ddd_4_usu());
         pst.setString(16,modexcel.getUra_usu());
         pst.setString(17,modexcel.getNu_ddd_5_usu());
         pst.setString(18,modexcel.getNet_fone_usu());
         pst.setString(19,modexcel.getNm_mailing_usu());
         pst.setString(20,modexcel.getNm_city_usu());
         pst.setString(21,modexcel.getDe_segment_strategy_usu());
         pst.setString(22,modexcel.getDe_extra_field_1_usu());
         pst.setString(23,modexcel.getDe_extra_field_2_usu());
         pst.setString(24,modexcel.getSegmento_usu());
         pst.setString(25,modexcel.getAlto_valor_usu());
         pst.setString(26,modexcel.getDt_desco_usu());
         pst.setString(27,modexcel.getDt_instalacao_usu());
         pst.setString(28,modexcel.getTipo_usu());
         pst.setString(29,modexcel.getDt_envio_usu());
         pst.setString(30,modexcel.getDt_atendimento_usu());



         pst.execute();
         
            //JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       conex.desconecta();
    }
    
    public void Importar_04 (BeansExcel modexcel){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into adequacao(base,perfil,aging_cliente,code_action,code_operator,code_net,code_client,code_address,nu_ddd_1,residencial,nu_ddd_2,celular,nu_ddd_3,comercial,nu_ddd_4,ura,nu_ddd_5,net_fone,nm_mailing,nm_city,de_segment_strategy,de_extra_field_1,de_extra_field_2,segmento,alto_valor,dt_desco,dt_instalacao,tipo,dt_envio,dt_atendimento) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1,modexcel.getBase_usu());
         pst.setString(2,modexcel.getPerfil_usu());
         pst.setString(3,modexcel.getAging_cliente_usu());
         pst.setString(4,modexcel.getCode_action_usu());
         pst.setString(5,modexcel.getCode_operator_usu());
         pst.setString(6,modexcel.getCode_net_usu());
         pst.setString(7,modexcel.getCode_client_usu());
         pst.setString(8,modexcel.getCode_address_usu());
         pst.setString(9,modexcel.getNu_ddd_1_usu());
         pst.setString(10,modexcel.getResidencial_usu());
         pst.setString(11,modexcel.getNu_ddd_2_usu());
         pst.setString(12,modexcel.getCelular_usu());
         pst.setString(13,modexcel.getNu_ddd_3_usu());
         pst.setString(14,modexcel.getComercial_usu());
         pst.setString(15,modexcel.getNu_ddd_4_usu());
         pst.setString(16,modexcel.getUra_usu());
         pst.setString(17,modexcel.getNu_ddd_5_usu());
         pst.setString(18,modexcel.getNet_fone_usu());
         pst.setString(19,modexcel.getNm_mailing_usu());
         pst.setString(20,modexcel.getNm_city_usu());
         pst.setString(21,modexcel.getDe_segment_strategy_usu());
         pst.setString(22,modexcel.getDe_extra_field_1_usu());
         pst.setString(23,modexcel.getDe_extra_field_2_usu());
         pst.setString(24,modexcel.getSegmento_usu());
         pst.setString(25,modexcel.getAlto_valor_usu());
         pst.setString(26,modexcel.getDt_desco_usu());
         pst.setString(27,modexcel.getDt_instalacao_usu());
         pst.setString(28,modexcel.getTipo_usu());
         pst.setString(29,modexcel.getDt_envio_usu());
         pst.setString(30,modexcel.getDt_atendimento_usu());



         pst.execute();
         
            //JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       conex.desconecta();
    }
    
    public void Importar_05 (BeansExcel modexcel){
        
       conex.conexao();
       
        try {
            PreparedStatement pst = conex.con.prepareStatement("insert into portabilidade(base,perfil,aging_cliente,code_action,code_operator,code_net,code_client,code_address,nu_ddd_1,residencial,nu_ddd_2,celular,nu_ddd_3,comercial,nu_ddd_4,ura,nu_ddd_5,net_fone,nm_mailing,nm_city,de_segment_strategy,de_extra_field_1,de_extra_field_2,segmento,alto_valor,dt_desco,dt_instalacao,tipo,dt_envio,dt_atendimento) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); // Comando insere os dados na tabela do banco de dados
        // inset into - é um código sql para inserir dados
        
         pst.setString(1,modexcel.getBase_usu());
         pst.setString(2,modexcel.getPerfil_usu());
         pst.setString(3,modexcel.getAging_cliente_usu());
         pst.setString(4,modexcel.getCode_action_usu());
         pst.setString(5,modexcel.getCode_operator_usu());
         pst.setString(6,modexcel.getCode_net_usu());
         pst.setString(7,modexcel.getCode_client_usu());
         pst.setString(8,modexcel.getCode_address_usu());
         pst.setString(9,modexcel.getNu_ddd_1_usu());
         pst.setString(10,modexcel.getResidencial_usu());
         pst.setString(11,modexcel.getNu_ddd_2_usu());
         pst.setString(12,modexcel.getCelular_usu());
         pst.setString(13,modexcel.getNu_ddd_3_usu());
         pst.setString(14,modexcel.getComercial_usu());
         pst.setString(15,modexcel.getNu_ddd_4_usu());
         pst.setString(16,modexcel.getUra_usu());
         pst.setString(17,modexcel.getNu_ddd_5_usu());
         pst.setString(18,modexcel.getNet_fone_usu());
         pst.setString(19,modexcel.getNm_mailing_usu());
         pst.setString(20,modexcel.getNm_city_usu());
         pst.setString(21,modexcel.getDe_segment_strategy_usu());
         pst.setString(22,modexcel.getDe_extra_field_1_usu());
         pst.setString(23,modexcel.getDe_extra_field_2_usu());
         pst.setString(24,modexcel.getSegmento_usu());
         pst.setString(25,modexcel.getAlto_valor_usu());
         pst.setString(26,modexcel.getDt_desco_usu());
         pst.setString(27,modexcel.getDt_instalacao_usu());
         pst.setString(28,modexcel.getTipo_usu());
         pst.setString(29,modexcel.getDt_envio_usu());
         pst.setString(30,modexcel.getDt_atendimento_usu());
         pst.execute();
         
            //JOptionPane.showMessageDialog(null, "Usuário Cadastrado Com Sucessao!");
            
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Erro Ao Cadastrar Usuário!\nErro: "+ex);
        }
        
       conex.desconecta();
    }
    public void ExcluirF (BeansUsuários mod){
        
        conex.conexao();
        
        try {
            PreparedStatement pst = conex.con.prepareStatement("delete from "+mod.getUsu_Puxe_01()+" where nm_mailing=?");
            
            pst.setString(1, mod.getPuxe_Cliente());
            pst.execute();
                    } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null,"Erro Ao Excluir Dados!\nErro: "+ex);
        }
    }
}
