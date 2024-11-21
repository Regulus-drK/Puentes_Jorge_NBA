/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.proyectosenus;

import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author drank
 */
public class Proyectosenus {

    public static void main(String[] args) throws IOException {
        // try {
        //    crearInformeExcel("informe_excel.xlsx");
          //  System.out.println("Informe Excel creado con éxito");
        //} catch (IOException e) {
          //  System.err.println("Error: " + e.getMessage());
        //}
        InformesExcel informesExcel = new InformesExcel();
        informesExcel.main(new String[0]);
    }
    
    private static void crearInformeExcel(String nombreArchivo) throws IOException {
        Workbook libroExcel = new XSSFWorkbook();
        Sheet hoja = libroExcel.createSheet("MiHoja");
        Row fila = hoja.createRow(0);
        fila.createCell(0).setCellValue("Nombre");
        fila.createCell(1).setCellValue("Edad");
        fila.createCell(2).setCellValue("Ciudad");
        String[][] datos = {
            {"Juan", "25", "Madrid"},
            {"María", "30", "Barcelona"},
            {"Pedro", "28", "Sevilla"}
        };
        for (int i = 0; i < datos.length; i++) {
            fila = hoja.createRow(i + 1);
            for (int j = 0; j < datos[i].length; j++) {
                fila.createCell(j).setCellValue(datos[i][j]);
            }
        }
        try (FileOutputStream archivo = new FileOutputStream(nombreArchivo)) {
            libroExcel.write(archivo);
        }
        libroExcel.close();
    }
}
