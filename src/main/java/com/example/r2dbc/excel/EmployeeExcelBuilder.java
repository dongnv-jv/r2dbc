package com.example.r2dbc.excel;

import com.example.r2dbc.entity.Employee;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

public class EmployeeExcelBuilder {
    public static Mono<ByteArrayInputStream> generateExcelFromFlux(Flux<Employee> studentFlux) {
        return Mono.create(sink -> {
            Workbook workbook = new XSSFWorkbook();
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            Sheet sheet = workbook.createSheet("Students");

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("City");

            // Use a counter to keep track of the current row index
            int[] rowIndex = {1};

            // Subscribe to the Flux and write data to the Excel sheet
            studentFlux.subscribe(student -> {
                Row row = sheet.createRow(rowIndex[0]++);
                row.createCell(0).setCellValue(student.getId());
                row.createCell(1).setCellValue(student.getName());
                row.createCell(2).setCellValue(student.getDepartment());
            }, sink::error, () -> {
                // Convert the workbook to a ByteArrayInputStream when the Flux completes
                try {
                    workbook.write(out);
                    ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(out.toByteArray());
                    sink.success(byteArrayInputStream);
                } catch (IOException e) {
                    sink.error(e);
                } finally {
                    try {
                        workbook.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            });
        });
    }
}
