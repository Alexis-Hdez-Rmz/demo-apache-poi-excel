package excellectura;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelLecturaTabla {

    public static void main(String[] args) {

        // 1. Crear una referencia del archivo
        File archivo = new File("Datos.xlsx");

        try {
            // Importar el archivo
            InputStream input = new FileInputStream(archivo);

            // Crear el libro
            XSSFWorkbook libro = new XSSFWorkbook(input);

            // Referenciar hoja por índice
            XSSFSheet hoja = libro.getSheetAt(2);

            // Referenciar filas y columnas
            Iterator<Row> filas = hoja.rowIterator();
            Iterator<Cell> columnas = null;

            Row filaActual = null;
            Cell columnaActual = null;

            // Recorrer filas
            while (filas.hasNext()) {

                filaActual = filas.next();
                columnas = filaActual.cellIterator();

                // Recorrer columnas
                while (columnas.hasNext()) {

                    columnaActual = columnas.next();

                    if(columnaActual.getCellType() == CellType.STRING) {
                        String valor = columnaActual.getStringCellValue();
                        System.out.println(valor);
                    }

                    if(columnaActual.getCellType() == CellType.NUMERIC) {
                        double valor = columnaActual.getNumericCellValue();
                        System.out.println(valor);
                    }

                    if(columnaActual.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(columnaActual)) {
                        SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
                        Date fecha = columnaActual.getDateCellValue();
                        System.out.println(formato.format(fecha));
                    }
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
