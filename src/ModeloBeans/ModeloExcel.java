/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ModeloBeans;
import java.io.*;
import java.util.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Tiago Dantas
 */
public class ModeloExcel {
    Workbook wb;
    
    public String Importar(File Arquivo, JTable TabelaD){
        String resposta = "Não Foi Possível Importação!";
        DefaultTableModel modeloT = new DefaultTableModel();
        TabelaD.setModel(modeloT);
        try {
            wb = WorkbookFactory.create(new FileInputStream(Arquivo));
            Sheet agora = wb.getSheetAt(0);
            Iterator filaIterator = agora.rowIterator();
            int indiceFila =-1;
            while (filaIterator.hasNext()){
                indiceFila++;
                Row fila = (Row) filaIterator.next();
                Iterator columnaIterator = fila.cellIterator();
                Object[] listaColumna = new Object[1000];
                int indiceColumna=-1;
                while (columnaIterator.hasNext()) {                    
                    indiceColumna++;
                    Cell celula = (Cell) columnaIterator.next();
                    if(indiceFila==0){
                        modeloT.addColumn(celula.getStringCellValue());
                    }else{
                        if(celula!=null){
                            switch(celula.getCellType()){
                                case Cell.CELL_TYPE_NUMERIC:
                                    listaColumna[indiceColumna]= (int)Math.round(celula.getNumericCellValue());
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    listaColumna[indiceColumna]= celula.getStringCellValue();
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    listaColumna[indiceColumna]= celula.getBooleanCellValue();
                                    break;
                                default:
                                    listaColumna[indiceColumna]=celula.getDateCellValue();
                                    break;
                            }
                            System.out.println("col"+indiceColumna+" Valor: true - "+celula+".");
                        }                        
                    }
                }
                if(indiceFila!=0)modeloT.addRow(listaColumna);
            }
            resposta="Importação Realizada Com Sucesso!";
        } catch (IOException | InvalidFormatException | EncryptedDocumentException e) {
            System.err.println(e.getMessage());
        }
        return resposta;
    }
    
    public String Exportar(File Arquivo, JTable TabelaD){
        String resposta="Não Foi Possível Exportação!.";
        int numFila=TabelaD.getRowCount(), numColumna=TabelaD.getColumnCount();
        if(Arquivo.getName().endsWith("xls")){
            wb = new HSSFWorkbook();
        }else{
            wb = new XSSFWorkbook();
        }
        Sheet hoja = wb.createSheet("Pruebita");
        
        try {
            for (int i = -1; i < numFila; i++) {
                Row fila = hoja.createRow(i+1);
                for (int j = 0; j < numColumna; j++) {
                    Cell celula = fila.createCell(j);
                    if(i==-1){
                        celula.setCellValue(String.valueOf(TabelaD.getColumnName(j)));
                    }else{
                        celula.setCellValue(String.valueOf(TabelaD.getValueAt(i, j)));
                    }
                    wb.write(new FileOutputStream(Arquivo));
                }
            }
            resposta="Exportação Realizada Com Sucesso!";
        } catch (Exception e) {
            System.err.println(e.getMessage());
      
        }
        return resposta;
            }
}
