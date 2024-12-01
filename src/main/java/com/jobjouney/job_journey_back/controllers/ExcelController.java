package com.jobjouney.job_journey_back.controllers;

import com.jobjouney.job_journey_back.models.FormData;
import com.jobjouney.job_journey_back.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/excel")
@CrossOrigin(origins = "http://localhost:4200", originPatterns = "*")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @PostMapping("/save")
    public ResponseEntity<?> saveData(@RequestBody FormData formData) {
        try {
            excelService.saveDataToExcel(formData);
            // Cambiar la respuesta a un objeto JSON
            return ResponseEntity.ok(Map.of("message", "Datos guardados exitosamente."));
        } catch (Exception e) {
            // Devuelve un objeto JSON con detalles del error
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(Map.of("error", "Error al guardar los datos en Excel.", "detalle", e.getMessage()));
        }
    }   

    @PostMapping("/upload")
    public ResponseEntity<?> uploadExcelFile(@RequestParam("file") MultipartFile file) {
        if (!file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().body(Map.of("error", "Formato de archivo no válido."));
        }

        try {
            excelService.processExcelFile(file);
            return ResponseEntity.ok("Archivo Excel procesado correctamente.");
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(Map.of("error", "Error al procesar el archivo.", "detalle", e.getMessage()));
        }
    }

    @GetMapping("/data")
    public ResponseEntity<?> getDataFromExcel() {
        try {
            List<FormData> data = excelService.readDataFromExcel();
            return ResponseEntity.ok(data);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(Map.of("error", "Error al leer datos del Excel.", "detalle", e.getMessage()));
        }
    }
}
