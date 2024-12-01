package com.jobjouney.job_journey_back.models;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor

public class FormData {
    private String nombre;
    private int edad;
    private String fechaNacimiento;
    private String genero;
    private String correo;
    private String telefono;
    private String direccion;
    private String ciudad;
    private String redes;
    private String nivelEducativo;
    private String institucion;
    private String titulo;
    private int anioGraduacion;
    private String empresa;
    private String cargo;
    private String fechaInicio;
    private String fechaFin;
    private String experiencia;
    private String responsabilidades;
}
