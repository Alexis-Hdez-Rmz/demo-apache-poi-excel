package excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class PrincipalExcel {

    public static void main(String[] args) {
        /*
         * Workbook libro = new XSSFWorkbook();  // .xlsx 2007 - en adelante
         * Workbook libro2 = new HSSFWorkbook(); // .xls  1997 - 2003
         */

        // 1. Crear un libro
        Workbook libro = new XSSFWorkbook();

        // 2. Crear hojas
        Sheet hoja = libro.createSheet("Test1");
        Sheet hoja2 = libro.createSheet("Test2");

        try {
            // Exporta los bytes a un documento con extensión .xlsx
            OutputStream output = new FileOutputStream("TestExcel.xlsx");
            libro.write(output);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
