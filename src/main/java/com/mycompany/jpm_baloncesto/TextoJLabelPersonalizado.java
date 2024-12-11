/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.jpm_baloncesto;

import java.awt.Color;
import java.awt.Component;
import java.awt.Font;
import javax.swing.JLabel;
import javax.swing.SwingConstants;

/**
 *
 * @author drank
 */
public class TextoJLabelPersonalizado extends JLabel {
    // Constructor sin parámetros
    public TextoJLabelPersonalizado() {
        super("Texto por defecto"); // Texto predeterminado
        configurarEstilo();
    }

    // Constructor con texto personalizado
    public TextoJLabelPersonalizado(String texto) {
        super(texto);
        configurarEstilo();
    }

    private void configurarEstilo() {
        setForeground(Color.YELLOW); // Texto amarillo
        setBackground(Color.MAGENTA); // Fondo morado
        setOpaque(true); // Hacer visible el fondo
        setHorizontalAlignment(SwingConstants.CENTER); // Centrar texto
        setFont(new Font("Arial", Font.BOLD, 14)); // Fuente personalizada
    }
    
    public void cambiarTamanio(Integer parametro) {
        if (parametro == 1) {
            setFont(new Font("Arial", Font.BOLD, 11));
        } else if (parametro == 2) {
            setFont(new Font("Arial", Font.BOLD, 14));
        } else if (parametro == 3) {
            setFont(new Font("Arial", Font.BOLD, 18));
        } else {
            setFont(new Font("Arial", Font.BOLD, 14));
        }
    }   
}
