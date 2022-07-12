import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class App {
    public static void main(String[] args) throws IOException {
        // Fonte
        FileInputStream FIS = new FileInputStream("./src/Resource/RFCV_185494_20220711_093049.xlsx");

        // Criando workbook
        XSSFWorkbook WB = new XSSFWorkbook(FIS);

        // Criando sheet/planilha
        XSSFSheet SH =  WB.getSheetAt(0);

        // Evaluando(?) planilha
        FormulaEvaluator evaluator = WB.getCreationHelper().createFormulaEvaluator();
        for(Row coluna: SH){
            for(Cell celula: coluna){
                switch(evaluator.evaluateInCell(celula).getCellType()){
                    case NUMERIC:
                        System.out.println(celula.getNumericCellValue() + "\t\t");
                        break;

                    case STRING:
                        System.out.println(celula.getStringCellValue() + "\t\t");
                        break;

                    case BOOLEAN:
                        System.out.println(celula.getBooleanCellValue() + "\t\t");
                        break;

                    case BLANK:
                        System.out.println("- \t\t");
                        break;

                    case ERROR:

                    case FORMULA:

                    case _NONE:
                        break;

                }
            }
            System.out.println();
        }

        WB.close();
    }
}
