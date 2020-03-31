/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Controle;

import EditionDao.DaoUsuario;
import ModeloBeans.BeansExcel;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import Visão.VistaExcel;
import ModeloBeans.ModeloExcel;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import jxl.Sheet;
import jxl.Workbook;
import sun.awt.windows.ThemeReader;

/**
 *
 * @author Tiago Dantas
 */
public class ControladorExcel implements ActionListener {
    ModeloExcel modeloE = new ModeloExcel();
    VistaExcel vistaE= new VistaExcel();
    JFileChooser selecinarArquivo = new JFileChooser();
    File arquivo;
    int contador=0;
    BeansExcel modexcel = new BeansExcel();
    DaoUsuario dao = new DaoUsuario();
    int i = 0;
    
                
    public ControladorExcel(VistaExcel vistaE, ModeloExcel modeloE){
        this.vistaE= vistaE;
        this.modeloE= modeloE;
        this.vistaE.btn_Importar.addActionListener(this);
        this.vistaE.btn_Exportar.addActionListener(this);
    }
    
    public void AgregarFiltro(){
        selecinarArquivo.setFileFilter(new FileNameExtensionFilter("Excel (*.xls)", "xls"));
        selecinarArquivo.setFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));
    }
    
    @Override
    public void actionPerformed(ActionEvent e) {
        contador++;
        if(contador==1)AgregarFiltro();
        
        if(e.getSource() == vistaE.btn_Importar){
            if(selecinarArquivo.showDialog(null, "Selecione o Arquivo.")==JFileChooser.APPROVE_OPTION){
                arquivo=selecinarArquivo.getSelectedFile();
                if(arquivo.getName().endsWith("xls") || arquivo.getName().endsWith("xlsx")){
                    JOptionPane.showMessageDialog(null, modeloE.Importar(arquivo, vistaE.jtDados) + "\n Formato ."+ arquivo.getName().substring(arquivo.getName().lastIndexOf(".")+1), 
                            "IMPORTAR EXCEL", JOptionPane.INFORMATION_MESSAGE);
                    
                         
               
            for(int i=0; i<vistaE.jtDados.getRowCount(); i++){
                                        
            modexcel.setBase_usu(vistaE.jtDados.getValueAt(i , 0).toString());
            modexcel.setPerfil_usu(vistaE.jtDados.getValueAt(i , 1).toString());
            modexcel.setAging_cliente_usu(vistaE.jtDados.getValueAt(i , 2).toString());
            modexcel.setCode_action_usu(vistaE.jtDados.getValueAt(i , 3).toString());
            modexcel.setCode_operator_usu(vistaE.jtDados.getValueAt(i , 4).toString());
            modexcel.setCode_net_usu(vistaE.jtDados.getValueAt(i , 5).toString());
            modexcel.setCode_client_usu(vistaE.jtDados.getValueAt(i , 6).toString());
            modexcel.setCode_address_usu(vistaE.jtDados.getValueAt(i , 7).toString());
            modexcel.setNu_ddd_1_usu(vistaE.jtDados.getValueAt(i , 8).toString());
            modexcel.setResidencial_usu(vistaE.jtDados.getValueAt(i , 9).toString());
            modexcel.setNu_ddd_2_usu(vistaE.jtDados.getValueAt(i , 10).toString());
            modexcel.setCelular_usu(vistaE.jtDados.getValueAt(i , 11).toString());
            modexcel.setNu_ddd_3_usu(vistaE.jtDados.getValueAt(i , 12).toString());
            modexcel.setComercial_usu(vistaE.jtDados.getValueAt(i , 13).toString());
            modexcel.setNu_ddd_4_usu(vistaE.jtDados.getValueAt(i , 14).toString());
            modexcel.setUra_usu(vistaE.jtDados.getValueAt(i , 15).toString());
            modexcel.setNu_ddd_5_usu(vistaE.jtDados.getValueAt(i , 16).toString());
            modexcel.setNet_fone_usu(vistaE.jtDados.getValueAt(i , 17).toString());
            modexcel.setNm_mailing_usu(vistaE.jtDados.getValueAt(i , 18).toString());
            modexcel.setNm_city_usu(vistaE.jtDados.getValueAt(i , 19).toString());
            modexcel.setDe_segment_strategy_usu(vistaE.jtDados.getValueAt(i , 20).toString());
            modexcel.setDe_extra_field_1_usu(vistaE.jtDados.getValueAt(i , 21).toString());
            modexcel.setDe_extra_field_2_usu(vistaE.jtDados.getValueAt(i , 22).toString());
            modexcel.setSegmento_usu(vistaE.jtDados.getValueAt(i , 23).toString());
            modexcel.setAlto_valor_usu(vistaE.jtDados.getValueAt(i , 24).toString());
            modexcel.setDt_desco_usu(vistaE.jtDados.getValueAt(i , 25).toString());
            modexcel.setDt_instalacao_usu(vistaE.jtDados.getValueAt(i , 26).toString());
            modexcel.setTipo_usu(vistaE.jtDados.getValueAt(i , 27).toString());
            modexcel.setDt_envio_usu(vistaE.jtDados.getValueAt(i , 28).toString());
            modexcel.setDt_atendimento_usu(vistaE.jtDados.getValueAt(i , 29).toString());
            
            dao.Importar(modexcel);//Comando para inserir dados no banco.
            

            }
                }else{
                    JOptionPane.showMessageDialog(null, "Formato Inválido.");
                }
            }
        }
        
        if(e.getSource() == vistaE.btn_Exportar){
            if(selecinarArquivo.showDialog(null, "Exportar")==JFileChooser.APPROVE_OPTION){
                arquivo=selecinarArquivo.getSelectedFile();
                if(arquivo.getName().endsWith("xls") || arquivo.getName().endsWith("xlsx")){
                    JOptionPane.showMessageDialog(null, modeloE.Exportar(arquivo, vistaE.jtDados) + "\n Formato ."+ arquivo.getName().substring(arquivo.getName().lastIndexOf(".")+1));
                }else{
                    JOptionPane.showMessageDialog(null, "Formato Inválido.");
                }
            }
        }
    }

    }
    


