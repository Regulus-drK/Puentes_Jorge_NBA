/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.examendi_1224_jpm;

import java.awt.Color;
import java.awt.Font;
import javax.swing.JLabel;
import javax.swing.SwingConstants;

/**
 *
 * @author drank
 */
public class TextoAprobado_Suspenso extends JLabel {
        // Constructor sin parámetros
    public TextoAprobado_Suspenso() {
        super("Suspenso"); // Texto predeterminado
        configurarEstilo();
    }

    // Constructor con texto personalizado
    public TextoAprobado_Suspenso(String texto) {
        super(texto);
        configurarEstilo();
    }

    public void detectarNota(Integer parametro) {
        if (parametro == 1) {
            setText("Suspenso");
            setForeground(Color.RED);
        } else if (parametro == 2) {
            setText("Aprobado");
            setForeground(Color.BLACK);
        }
    }
    
    private void configurarEstilo() {
        setForeground(Color.RED);
        setHorizontalAlignment(SwingConstants.CENTER); // Centrar texto
        setFont(new Font("Arial", Font.BOLD, 14)); // Fuente personalizada
    }
}
