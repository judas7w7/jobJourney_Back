package com.jobjouney.job_journey_back.service;

import com.jobjouney.job_journey_back.models.FormData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {
    private static final String FILE_PATH = "data.xlsx";

    // Método para verificar si el archivo existe y crearlo si no
    private void createExcelFileIfNotExists() {
        File file = new File(FILE_PATH);
        if (!file.exists()) {
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Datos");

                // Crear fila de encabezados
                Row headerRow = sheet.createRow(0);
                String[] headers = { "Nombre", "Edad", "FechaNacimiento", "Género", "Correo", "Teléfono",
                        "Dirección", "Ciudad", "Redes", "NivelEducativo", "Institución",
                        "Título", "AñoGraduación", "Empresa", "Cargo", "FechaInicio",
                        "FechaFin", "Experiencia", "Responsabilidades" };
                for (int i = 0; i < headers.length; i++) {
                    headerRow.createCell(i).setCellValue(headers[i]);
                }

                // Guardar el archivo
                try (FileOutputStream fos = new FileOutputStream(FILE_PATH)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void saveDataToExcel(FormData formData) {
        createExcelFileIfNotExists(); // Verificar y crear el archivo si no existe

        try (FileInputStream fis = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Crear nueva fila al final
            int lastRowNum = sheet.getLastRowNum();
            Row newRow = sheet.createRow(lastRowNum + 1);

            // Escribir datos en las celdas
            newRow.createCell(0).setCellValue(formData.getNombre());
            newRow.createCell(1).setCellValue(formData.getEdad());
            newRow.createCell(2).setCellValue(formData.getFechaNacimiento());
            newRow.createCell(3).setCellValue(formData.getGenero());
            newRow.createCell(4).setCellValue(formData.getCorreo());
            newRow.createCell(5).setCellValue(formData.getTelefono());
            newRow.createCell(6).setCellValue(formData.getDireccion());
            newRow.createCell(7).setCellValue(formData.getCiudad());
            newRow.createCell(8).setCellValue(formData.getRedes());
            newRow.createCell(9).setCellValue(formData.getNivelEducativo());
            newRow.createCell(10).setCellValue(formData.getInstitucion());
            newRow.createCell(11).setCellValue(formData.getTitulo());
            newRow.createCell(12).setCellValue(formData.getAnioGraduacion());
            newRow.createCell(13).setCellValue(formData.getEmpresa());
            newRow.createCell(14).setCellValue(formData.getCargo());
            newRow.createCell(15).setCellValue(formData.getFechaInicio());
            newRow.createCell(16).setCellValue(formData.getFechaFin());
            newRow.createCell(17).setCellValue(formData.getExperiencia());
            newRow.createCell(18).setCellValue(formData.getResponsabilidades());

            // Guardar los cambios
            try (FileOutputStream fos = new FileOutputStream(FILE_PATH)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public List<FormData> readDataFromExcel() {
        createExcelFileIfNotExists(); // Verificar y crear el archivo si no existe

        List<FormData> dataList = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                FormData formData = new FormData();
                
                // Verificar si las celdas son nulas antes de acceder a sus valores
                formData.setNombre(getCellStringValue(row.getCell(0)));
                formData.setEdad(getCellNumericValue(row.getCell(1)));
                formData.setFechaNacimiento(getCellStringValue(row.getCell(2)));
                formData.setGenero(getCellStringValue(row.getCell(3)));
                formData.setCorreo(getCellStringValue(row.getCell(4)));
                formData.setTelefono(getCellStringValue(row.getCell(5)));
                formData.setDireccion(getCellStringValue(row.getCell(6)));
                formData.setCiudad(getCellStringValue(row.getCell(7)));
                formData.setRedes(getCellStringValue(row.getCell(8)));
                formData.setNivelEducativo(getCellStringValue(row.getCell(9)));
                formData.setInstitucion(getCellStringValue(row.getCell(10)));
                formData.setTitulo(getCellStringValue(row.getCell(11)));
                formData.setAnioGraduacion(getCellNumericValue(row.getCell(12)));
                formData.setEmpresa(getCellStringValue(row.getCell(13)));
                formData.setCargo(getCellStringValue(row.getCell(14)));
                formData.setFechaInicio(getCellStringValue(row.getCell(15)));
                formData.setFechaFin(getCellStringValue(row.getCell(16)));
                formData.setExperiencia(getCellStringValue(row.getCell(17)));
                formData.setResponsabilidades(getCellStringValue(row.getCell(18)));

                dataList.add(formData);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataList;
    }

    // Método para obtener el valor de una celda de tipo String
    private String getCellStringValue(Cell cell) {
        if (cell != null && cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return ""; // Valor predeterminado si la celda es nula o no es de tipo String
    }

    // Método para obtener el valor de una celda de tipo numérico
    private int getCellNumericValue(Cell cell) {
        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
            return (int) cell.getNumericCellValue();
        }
        return 0; // Valor predeterminado si la celda es nula o no es de tipo numérico
    }

    public void processExcelFile(MultipartFile file) throws IOException {
        // Verifica si el archivo es válido
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("El archivo está vacío o es nulo.");
        }

        // Obtén el flujo de entrada del archivo
        InputStream inputStream = file.getInputStream();

        // Usa Apache POI para leer el archivo Excel
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0); // Obtén la primera hoja
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // Lógica para procesar cada celda del archivo Excel
                    System.out.print(cell.toString() + " ");
                }
                System.out.println(); // Salto de línea entre filas
            }
        } catch (Exception e) {
            throw new IOException("Error al procesar el archivo Excel: " + e.getMessage(), e);
        }
    }
}
