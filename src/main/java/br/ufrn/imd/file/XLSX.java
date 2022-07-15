package br.ufrn.imd.file;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class XLSX {
    private FileInputStream FIS;

    private FormulaEvaluator evaluator;
    private XSSFSheet SH;
    private XSSFWorkbook WB;

    public XLSX(String path) throws IOException, FileNotFoundException {
      this.FIS =  new FileInputStream(path);
      openXLSX();
    }
    private void openXLSX() throws IOException{
        // Creating XSSF workbook
        this.WB = new XSSFWorkbook(FIS);

        // Creating sheet
        this.SH =  WB.getSheetAt(0);

        // Evaluating sheet
        this.evaluator = WB.getCreationHelper().createFormulaEvaluator();
    }

//    public int readXLSX() throws IOException {
//        int lincount = 0;
//        int cont = 0;
//
//        for(Row r: SH){
//            lincount += 1;
//            for(Cell cell : r){
//                if(cont == 0){
//                    cont += 1;
//                }
//                switch(evaluator.evaluateInCell(cell).getCellType()){
//                    case NUMERIC:
//                        System.out.print(cell.getNumericCellValue() + "\t\t");
//                        break;
//
//                    case STRING:
//                        System.out.print(cell.getStringCellValue() + "\t\t");
//                        break;
//
//                    case BOOLEAN:
//                        System.out.print(cell.getBooleanCellValue() + "\t\t");
//                        break;
//
//                    case BLANK:
//                        System.out.print("- \t\t");
//                        break;
//
//                    case ERROR:
//                    case FORMULA:
//                    case _NONE:
//                        break;
//
//                }
//            }
//            System.out.println();
//        }
//
//        WB.close();
//        return lincount;
//
//    }

    public int readXLSX(){

        List<String> nomesJaPassados = new ArrayList<String>();

        for(Row linha: SH){
            if(linha.getRowNum() > 1){
                // Nome da empresa
//            System.out.print(linha.getCell(2).toString() + "\t\t");

                String nomeCondutor = linha.getCell(10).toString();
                if (!nomesJaPassados.contains(nomeCondutor)) {
                    nomesJaPassados.add(nomeCondutor);
                    System.out.println(nomeCondutor);
                }
            }
//
        }
        return 0;
    }
}
