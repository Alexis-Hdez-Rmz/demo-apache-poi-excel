package excellectura;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

public class ExcelLecturaColumna {

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
            XSSFSheet hoja = libro.getSheetAt(0);

            // Referenciar una fila
            //Row fila = hoja.getRow(1);

            // Referenciar filas
            Iterator<Row> filas = hoja.rowIterator();

            Cell columna = null;

            // Recorrer filas
            while(filas.hasNext()) {
                columna = filas.next().getCell(0);
                System.out.println(columna.getStringCellValue());
            }
            // Cerrar recursos
            input.close();
            libro.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
