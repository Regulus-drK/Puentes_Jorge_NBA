/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.jpm_baloncesto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author drank
 */
public class EjercicioBaloncestoExcel extends javax.swing.JFrame {
    private static String nombreJugador;
    private static Integer tirosDe2Realizados, tirosDe2Metidos, triplesRealizados, triplesMetidos, 
            tirosLibresMetidos, tirosLibresRealizados, rebotes, asistencias, taponesFavor, taponesContra,
            robos, perdidas;
    /**
     * Creates new form EjercicioBaloncestoExcel
     */
    public EjercicioBaloncestoExcel() {
        
        initComponents();
        setLocationRelativeTo(null);
        this.setResizable(false);
        this.setTitle("Estadísticas Baloncesto");
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        paneles = new javax.swing.JTabbedPane();
        panelTiros = new javax.swing.JPanel();
        jLabel2 = new javax.swing.JLabel();
        campoTextoJugador = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        numTirosLibresMetidos = new javax.swing.JSpinner();
        jLabel7 = new javax.swing.JLabel();
        numTirosLibresRealizados = new javax.swing.JSpinner();
        jLabel3 = new javax.swing.JLabel();
        numTirosDe2Metidos = new javax.swing.JSpinner();
        jLabel8 = new javax.swing.JLabel();
        numTirosDe2Realizados = new javax.swing.JSpinner();
        jLabel4 = new javax.swing.JLabel();
        numTriplesMetidos = new javax.swing.JSpinner();
        jLabel5 = new javax.swing.JLabel();
        numTriplesRealizados = new javax.swing.JSpinner();
        panelOtrosDatos = new javax.swing.JPanel();
        numRebotes = new javax.swing.JSpinner();
        numTaponesFavor = new javax.swing.JSpinner();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jLabel14 = new javax.swing.JLabel();
        numRobos = new javax.swing.JSpinner();
        numAsistencias = new javax.swing.JSpinner();
        numTaponesContra = new javax.swing.JSpinner();
        numPerdidas = new javax.swing.JSpinner();
        botonInsertar = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel2.setText("Nombre Jugador:");

        campoTextoJugador.setToolTipText("Inserta el nombre");

        jLabel6.setText("Tiros libres metidos:");

        numTirosLibresMetidos.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTirosLibresMetidos.setPreferredSize(new java.awt.Dimension(90, 25));

        jLabel7.setText("Tiros libres realizados:");

        numTirosLibresRealizados.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTirosLibresRealizados.setPreferredSize(new java.awt.Dimension(90, 25));

        jLabel3.setText("Tiros de 2 metidos:");

        numTirosDe2Metidos.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTirosDe2Metidos.setPreferredSize(new java.awt.Dimension(90, 25));

        jLabel8.setText("Tiros de 2 realizados:");

        numTirosDe2Realizados.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTirosDe2Realizados.setPreferredSize(new java.awt.Dimension(90, 25));

        jLabel4.setText("Triples metidos:");

        numTriplesMetidos.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTriplesMetidos.setPreferredSize(new java.awt.Dimension(90, 25));

        jLabel5.setText("Triples realizados:");

        numTriplesRealizados.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTriplesRealizados.setPreferredSize(new java.awt.Dimension(90, 25));

        javax.swing.GroupLayout panelTirosLayout = new javax.swing.GroupLayout(panelTiros);
        panelTiros.setLayout(panelTirosLayout);
        panelTirosLayout.setHorizontalGroup(
            panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(panelTirosLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(panelTirosLayout.createSequentialGroup()
                        .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(panelTirosLayout.createSequentialGroup()
                                .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(19, 19, 19)
                                .addComponent(numTirosLibresMetidos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(panelTirosLayout.createSequentialGroup()
                                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(numTirosDe2Metidos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 76, Short.MAX_VALUE)
                        .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelTirosLayout.createSequentialGroup()
                                .addComponent(jLabel7)
                                .addGap(18, 18, 18)
                                .addComponent(numTirosLibresRealizados, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelTirosLayout.createSequentialGroup()
                                    .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGap(18, 18, 18)
                                    .addComponent(numTriplesRealizados, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGroup(panelTirosLayout.createSequentialGroup()
                                    .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGap(18, 18, 18)
                                    .addComponent(numTirosDe2Realizados, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)))))
                    .addGroup(panelTirosLayout.createSequentialGroup()
                        .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(panelTirosLayout.createSequentialGroup()
                                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(campoTextoJugador, javax.swing.GroupLayout.PREFERRED_SIZE, 159, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(panelTirosLayout.createSequentialGroup()
                                .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(19, 19, 19)
                                .addComponent(numTriplesMetidos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        panelTirosLayout.setVerticalGroup(
            panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelTirosLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(campoTextoJugador, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(numTirosLibresMetidos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel7)
                    .addComponent(numTirosLibresRealizados, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel8)
                    .addComponent(numTirosDe2Realizados, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3)
                    .addComponent(numTirosDe2Metidos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(numTriplesRealizados, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel5))
                    .addGroup(panelTirosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(numTriplesMetidos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel4)))
                .addContainerGap(55, Short.MAX_VALUE))
        );

        paneles.addTab("Tiros", panelTiros);

        numRebotes.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numRebotes.setPreferredSize(new java.awt.Dimension(90, 25));

        numTaponesFavor.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTaponesFavor.setPreferredSize(new java.awt.Dimension(90, 25));

        jLabel9.setText("Rebotes:");

        jLabel10.setText("Asistencias:");

        jLabel11.setText("Robos:");

        jLabel12.setText("Tapones a favor:");

        jLabel13.setText("Tapones en contra:");

        jLabel14.setText("Pérdidas:");

        numRobos.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numRobos.setPreferredSize(new java.awt.Dimension(90, 25));

        numAsistencias.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numAsistencias.setPreferredSize(new java.awt.Dimension(90, 25));

        numTaponesContra.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numTaponesContra.setPreferredSize(new java.awt.Dimension(90, 25));

        numPerdidas.setModel(new javax.swing.SpinnerNumberModel(0, 0, null, 1));
        numPerdidas.setPreferredSize(new java.awt.Dimension(90, 25));

        javax.swing.GroupLayout panelOtrosDatosLayout = new javax.swing.GroupLayout(panelOtrosDatos);
        panelOtrosDatos.setLayout(panelOtrosDatosLayout);
        panelOtrosDatosLayout.setHorizontalGroup(
            panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(panelOtrosDatosLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(panelOtrosDatosLayout.createSequentialGroup()
                        .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel11, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel12, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(numTaponesFavor, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(numRobos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(panelOtrosDatosLayout.createSequentialGroup()
                        .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(numRebotes, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(67, 67, 67)
                .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(panelOtrosDatosLayout.createSequentialGroup()
                        .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(numAsistencias, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(panelOtrosDatosLayout.createSequentialGroup()
                        .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(numTaponesContra, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(panelOtrosDatosLayout.createSequentialGroup()
                        .addComponent(jLabel14, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(numPerdidas, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(41, Short.MAX_VALUE))
        );
        panelOtrosDatosLayout.setVerticalGroup(
            panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelOtrosDatosLayout.createSequentialGroup()
                .addGap(28, 28, 28)
                .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(numRebotes, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel9)
                    .addComponent(jLabel10)
                    .addComponent(numAsistencias, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel12)
                    .addComponent(numTaponesFavor, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel13)
                    .addComponent(numTaponesContra, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(panelOtrosDatosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel11)
                    .addComponent(numRobos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel14)
                    .addComponent(numPerdidas, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(67, Short.MAX_VALUE))
        );

        paneles.addTab("Otros Datos", panelOtrosDatos);

        botonInsertar.setText("Insertar Datos");
        botonInsertar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                botonInsertarActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(paneles)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(botonInsertar)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(paneles, javax.swing.GroupLayout.PREFERRED_SIZE, 241, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(botonInsertar)
                .addGap(0, 10, Short.MAX_VALUE))
        );

        getContentPane().add(jPanel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 540, 280));

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void botonInsertarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_botonInsertarActionPerformed
        nombreJugador = campoTextoJugador.getText();
        tirosLibresMetidos = (Integer) numTirosLibresMetidos.getValue();
        tirosLibresRealizados = (Integer) numTirosLibresRealizados.getValue();
        tirosDe2Metidos = (Integer) numTirosDe2Metidos.getValue();
        tirosDe2Realizados = (Integer) numTirosDe2Realizados.getValue();
        triplesMetidos = (Integer) numTriplesMetidos.getValue();
        triplesRealizados = (Integer) numTriplesRealizados.getValue();
        rebotes = (Integer) numRebotes.getValue();
        asistencias = (Integer) numAsistencias.getValue();
        taponesFavor = (Integer) numTaponesFavor.getValue();
        taponesContra = (Integer) numTaponesContra.getValue();
        robos = (Integer) numRobos.getValue();
        perdidas = (Integer) numPerdidas.getValue();
        
        String nombreArchivo = "jugadores_baloncesto.xlsx";
        
        if (nombreJugador == null || nombreJugador.trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "Inserte un nombre.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        if ((tirosDe2Metidos + triplesMetidos + tirosLibresMetidos) > (triplesRealizados + tirosDe2Realizados + tirosLibresRealizados)) {
            JOptionPane.showMessageDialog(this, "La suma de los tiros metidos no puede ser mayor a los tiros realizados.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        if (archivoEstaAbierto(nombreArchivo)) {
            JOptionPane.showMessageDialog(this, "El archivo Excel se encuentra abierto o bloqueado por otro programa. Ciérrelo para escribir datos en él.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        crearInforme(nombreArchivo, nombreJugador, tirosLibresMetidos, tirosLibresRealizados, tirosDe2Metidos, tirosDe2Realizados, 
                triplesMetidos, triplesRealizados, rebotes, asistencias, taponesFavor, taponesContra, robos, perdidas);
        JOptionPane.showMessageDialog(this, "Datos insertados.", "Aviso", JOptionPane.INFORMATION_MESSAGE);
    }//GEN-LAST:event_botonInsertarActionPerformed

    public static boolean archivoEstaAbierto(String nombreArchivo) {
        try (FileOutputStream fos = new FileOutputStream(nombreArchivo, true)) {
            return false;
        } catch (IOException e) {
            return true;
        }
    }
    
    private static void crearInforme(String nombreArchivoIn, String nombreJugador, Integer tirosLibresMetidos, Integer tirosLibresRealizados, 
            Integer tirosDe2Metidos, Integer tirosDe2Realizados, 
            Integer triplesMetidos, Integer triplesRealizados, Integer rebotes, Integer asistencias, 
            Integer taponesFavor, Integer taponesContra, Integer robos, Integer perdidas) {
        String nombreArchivo = nombreArchivoIn;
        String nombreHoja = "Version2";
        boolean archivoExistente = verificarArchivoExistente(nombreArchivo);
         
        
        try (Workbook libroTrabajo = archivoExistente ? WorkbookFactory.create(new FileInputStream(nombreArchivo)) : new XSSFWorkbook()) {
            Sheet hoja = libroTrabajo.getSheet(nombreHoja);
            
            CellStyle estiloMain = libroTrabajo.createCellStyle();
            Font fuenteMain = libroTrabajo.createFont();
            fuenteMain.setBold(true);
            fuenteMain.setColor(IndexedColors.DARK_GREEN.getIndex());
            fuenteMain.setFontName("Calibri");
            estiloMain.setFont(fuenteMain);
            
            
            if (hoja == null) {
                hoja = libroTrabajo.createSheet(nombreHoja);
                Row fila = hoja.createRow(0);

                String[] encabezados = {
                    "Jugador", "Tiros libres metidos", "Tiros libres realizados",
                    "Tiros de 2 metidos", "Tiros de 2 realizados", "Triples metidos",
                    "Triples realizados", "Rebotes", "Asistencias", "Tapones a favor", 
                    "Tapones en contra", "Robos", "Perdidas", "Valoracion", "FG%", "eFG%", "TS%"
                };

                for (int i = 0; i < encabezados.length; i++) {
                    Cell celda = fila.createCell(i);
                    celda.setCellValue(encabezados[i]);
                    celda.setCellStyle(estiloMain);
                }
            }
            
            // Borrar fila Media si ya existe
            for (int i = hoja.getPhysicalNumberOfRows() - 1; i >= 0; i--) {
                Row filaExistente = hoja.getRow(i);
                if (filaExistente != null && filaExistente.getCell(0) != null && filaExistente.getCell(0).getStringCellValue().equals("Media")) {
                    hoja.removeRow(filaExistente);
                }
            }
            
            CellStyle estilo = libroTrabajo.createCellStyle();
            Font fuente = libroTrabajo.createFont();
            fuente.setBold(true);
            fuente.setColor(IndexedColors.BLUE.getIndex());
            fuente.setFontName("Arial");
            estilo.setFont(fuente);
            estilo.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            estilo.setAlignment(HorizontalAlignment.CENTER);
            estilo.setVerticalAlignment(VerticalAlignment.TOP);
            
            Double FG = ((double) (tirosDe2Metidos + triplesMetidos + tirosLibresMetidos) / (tirosDe2Realizados + triplesRealizados + tirosLibresRealizados)) * 100;
            Double eFG = (((tirosDe2Metidos + triplesMetidos + tirosLibresMetidos) + (0.5 * triplesMetidos)) / (tirosDe2Realizados + triplesRealizados + tirosLibresRealizados)) * 100;
            int puntosTotales = (2 * tirosDe2Metidos) + (3 * triplesMetidos) + tirosLibresMetidos;
            Double TS = (double) (puntosTotales / (2 * (tirosDe2Realizados + triplesRealizados + (0.44 * tirosLibresRealizados)))) * 100;
            
            Double tirosFallados = (double) (tirosDe2Realizados + triplesRealizados) - (tirosDe2Metidos + triplesMetidos);
            Double tirosLibresFallados = (double) (tirosLibresRealizados - tirosLibresMetidos);
            
            Double valoracion = (double) (puntosTotales + rebotes + asistencias + robos + taponesFavor) - (tirosFallados + tirosLibresFallados + perdidas + taponesContra);
            
            int filaNumero = hoja.getPhysicalNumberOfRows();
            Row fila = hoja.createRow(filaNumero);
            Cell celda = fila.createCell(0, CellType.STRING);
            celda.setCellValue(nombreJugador);
            celda = fila.createCell(1, CellType.NUMERIC);
            celda.setCellValue(tirosLibresMetidos);
            celda = fila.createCell(2, CellType.NUMERIC);
            celda.setCellValue(tirosLibresRealizados);
            celda = fila.createCell(3, CellType.NUMERIC);
            celda.setCellValue(tirosDe2Metidos);
            celda = fila.createCell(4, CellType.NUMERIC);
            celda.setCellValue(tirosDe2Realizados);
            celda = fila.createCell(5, CellType.NUMERIC);
            celda.setCellValue(triplesMetidos);
            celda = fila.createCell(6, CellType.NUMERIC);
            celda.setCellValue(triplesRealizados);
            celda = fila.createCell(7, CellType.NUMERIC);
            celda.setCellValue(rebotes);
            celda = fila.createCell(8, CellType.NUMERIC);
            celda.setCellValue(asistencias);
            celda = fila.createCell(9, CellType.NUMERIC);
            celda.setCellValue(taponesFavor);
            celda = fila.createCell(10, CellType.NUMERIC);
            celda.setCellValue(taponesContra);
            celda = fila.createCell(11, CellType.NUMERIC);
            celda.setCellValue(robos);
            celda = fila.createCell(12, CellType.NUMERIC);
            celda.setCellValue(perdidas);
            celda = fila.createCell(13, CellType.NUMERIC);
            celda.setCellValue(valoracion);
            celda = fila.createCell(14, CellType.NUMERIC);
            celda.setCellValue(FG);
            celda = fila.createCell(15, CellType.NUMERIC);
            celda.setCellValue(eFG);
            celda = fila.createCell(16, CellType.NUMERIC);
            celda.setCellValue(TS);
            

            // Luego, crea la nueva fila de media en la última posición
            Row filaMedia = hoja.createRow(hoja.getPhysicalNumberOfRows());
            Cell celdaMedia = filaMedia.createCell(0);
            celdaMedia.setCellValue("Media");
            celdaMedia.setCellStyle(estilo);

            // Cálculo de medias
            for (int col = 1; col <= 16; col++) {
                double suma = 0;
                int filasDatos = hoja.getPhysicalNumberOfRows() - 1;
                for (int i = 1; i <= filasDatos; i++) {
                    Row filaActual = hoja.getRow(i);
                    if (filaActual == null) {
                        filaActual = hoja.createRow(i);  // Crea la fila si no existe
                    }
                    if (filaActual != null && filaActual.getCell(col) != null) {
                        suma += filaActual.getCell(col).getNumericCellValue();
                    }
                }
                double media = suma / (filasDatos - 1); // Número total de filas con datos
                Cell celdasMedia = filaMedia.createCell(col, CellType.NUMERIC);
                celdasMedia.setCellValue(media);
                celdasMedia.setCellStyle(estilo);
            }

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
    /**
     * @param args the command line arguments
     */
    
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(EjercicioBaloncestoExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(EjercicioBaloncestoExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(EjercicioBaloncestoExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(EjercicioBaloncestoExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new EjercicioBaloncestoExcel().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton botonInsertar;
    private javax.swing.JTextField campoTextoJugador;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JSpinner numAsistencias;
    private javax.swing.JSpinner numPerdidas;
    private javax.swing.JSpinner numRebotes;
    private javax.swing.JSpinner numRobos;
    private javax.swing.JSpinner numTaponesContra;
    private javax.swing.JSpinner numTaponesFavor;
    private javax.swing.JSpinner numTirosDe2Metidos;
    private javax.swing.JSpinner numTirosDe2Realizados;
    private javax.swing.JSpinner numTirosLibresMetidos;
    private javax.swing.JSpinner numTirosLibresRealizados;
    private javax.swing.JSpinner numTriplesMetidos;
    private javax.swing.JSpinner numTriplesRealizados;
    private javax.swing.JPanel panelOtrosDatos;
    private javax.swing.JPanel panelTiros;
    private javax.swing.JTabbedPane paneles;
    // End of variables declaration//GEN-END:variables
}
