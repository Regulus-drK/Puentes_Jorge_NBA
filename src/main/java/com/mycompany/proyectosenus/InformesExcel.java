/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.proyectosenus;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author drank
 */
public class InformesExcel {
    public static void main(String[] args) throws IOException {
        String nombreArchivo = "informe_excel.xlsx";
        String nombreHoja = "Hoja2";
        boolean archivoExistente = verificarArchivoExistente(nombreArchivo);
        
        try (Workbook libroTrabajo = archivoExistente ? WorkbookFactory.create(new FileInputStream(nombreArchivo)) : new XSSFWorkbook()) {
            Sheet hoja = libroTrabajo.getSheet(nombreHoja);
            
            if (hoja == null) {
                hoja = libroTrabajo.createSheet(nombreHoja);
            }
            int filaNumero = hoja.getPhysicalNumberOfRows();
            
            // Variable para guardar la suma de las dos primeras columnas
            double valorNumerico = 0.0;
            for (int i = 0; i < filaNumero; i++) {
                Row filab = hoja.getRow(i);
                if (filab != null) { // Verifica que la fila no esté vacía
                        // Recorre cada celda de la fila
                        for (int cellIndex = 1; cellIndex < 2; cellIndex++) { // Lee máximo la columna 2 de las filas que haya disponibles
                            Cell celda = filab.getCell(cellIndex);
                            
                            if (celda != null) {
                                if (celda.getCellType() == CellType.NUMERIC) { // Saca el valor numerico de la celda si es del tipo numerico
                                    valorNumerico += celda.getNumericCellValue(); 
                                }
                            }
                        }
                }
            }
            
            
            Row fila = hoja.createRow(filaNumero);
            Cell celda = fila.createCell(0, CellType.STRING);
            celda.setCellValue("Suma2");
            celda = fila.createCell(1, CellType.NUMERIC);
            celda.setCellValue(valorNumerico);

            try (FileOutputStream archivoSalida = new FileOutputStream(nombreArchivo)) {
                libroTrabajo.write(archivoSalida);
                System.out.println("Datos agregados al archivo Excel correctamente.");
            }
            
        } catch (IOException e) {
                e.printStackTrace();
            }
        
        // Reabrir el archivo y leer los datos actualizados
        try (Workbook libroTrabajo = WorkbookFactory.create(new FileInputStream(nombreArchivo))) {
            Sheet hoja = libroTrabajo.getSheet(nombreHoja);
            int filaNumero = hoja.getPhysicalNumberOfRows();

            System.out.println("Datos actualizados en el archivo:");
            for (int i = 0; i < filaNumero; i++) {
                Row filab = hoja.getRow(i);
                if (filab != null) { // Verifica que la fila no esté vacía
                    for (int cellIndex = 0; cellIndex < filab.getLastCellNum(); cellIndex++) {
                        Cell celda = filab.getCell(cellIndex);

                        // Verifica si la celda contiene algún valor
                        if (celda != null) {
                            System.out.println("Valor encontrado en hoja " + hoja.getSheetName() + ", fila " + (i + 1) + ", columna " + (cellIndex + 1) + ": " + celda.toString());
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    private static boolean verificarArchivoExistente(String nombreArchivo) {
        return new File(nombreArchivo).exists();
    }
    
}
