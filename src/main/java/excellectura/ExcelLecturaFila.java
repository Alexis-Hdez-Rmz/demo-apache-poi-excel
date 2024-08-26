package excellectura;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Date;
import java.util.Iterator;

public class ExcelLecturaFila {

    public static void main(String[] args) {
        /*
         * Si el archivo se encuentra ubicado en la raíz del proyecto, solo pasar como
         * parámetro el nombre del archivo con extensión
         * Si el archivo no se encuentra ubicado en la raíz del proyecto, pasar como
         * parámetro la ruta del archivo
         */
        // 1. Crear una referencia del archivo
        File archivo = new File("Datos.xlsx");

        try {
            // Importar el archivo
            InputStream input = new FileInputStream(archivo);

            // Crear el libro
            XSSFWorkbook libro = new XSSFWorkbook(input);

            // Referenciar hoja por índice
            XSSFSheet hoja = libro.getSheetAt(1);

            // Referenciar una fila
            Row fila = hoja.getRow(0);

            // Referenciar columnas
            Iterator<Cell> columnas = fila.cellIterator();

            // Recorrer columnas
            while(columnas.hasNext()) {

                Cell celda = columnas.next();

                // Valor String
                if(celda.getCellType() == CellType.STRING) {
                    String valor = celda.getStringCellValue();
                    System.out.println(valor);
                }

                // Valor Númerico
                if(celda.getCellType() == CellType.NUMERIC) {
                    double valor = celda.getNumericCellValue();
                    System.out.println(valor);
                }

                // Valor Fecha
                if(celda.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(celda)) {
                    Date fecha = celda.getDateCellValue();
                    System.out.println(fecha);
                }
            }
            // Cerrar recursos
            input.close();
            libro.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
