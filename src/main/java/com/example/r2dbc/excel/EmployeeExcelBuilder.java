package com.example.r2dbc.excel;

import com.example.r2dbc.entity.Employee;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

public class EmployeeExcelBuilder {


//    public static Mono<ByteArrayInputStream> generateExcelFromFlux(Flux<Employee> studentFlux) {
//        return Mono.create(sink -> {
//            Workbook workbook = new XSSFWorkbook();
//            ByteArrayOutputStream out = new ByteArrayOutputStream();
//
//            SheetWrapper sheetWrapper = new SheetWrapper(workbook);
//
//            // Sử dụng một biến để theo dõi chỉ số hàng hiện tại
//            int[] rowIndex = {1};
//
//            // Subscribe to the Flux and write data to the Excel sheet
//            studentFlux.subscribe(student -> {
//                // Kiểm tra nếu số dòng vượt quá 1,000,000
//                if (rowIndex[0] > 300000) {
//                    sheetWrapper.createNewSheet(workbook);
//                    rowIndex[0] = 1; // Đặt lại chỉ số hàng
//                }
//                Row row = sheetWrapper.sheet.createRow(rowIndex[0]++);
//                row.createCell(0).setCellValue(student.getId());
//                row.createCell(1).setCellValue(student.getName());
//                row.createCell(2).setCellValue(student.getDepartment());
//            }, sink::error, () -> {
//                // Chuyển đổi workbook thành ByteArrayInputStream khi Flux hoàn thành
//                try {
//                    workbook.write(out);
//                    ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(out.toByteArray());
//                    sink.success(byteArrayInputStream);
//                } catch (IOException e) {
//                    sink.error(e);
//                } finally {
//                    try {
//                        workbook.close();
//                    } catch (IOException e) {
//                        e.printStackTrace();
//                    }
//                }
//            });
//        });
//    }
private static final int MAX_ROWS_PER_SHEET = 500000;

    public static Mono<ByteArrayInputStream> generateExcelFromFlux(Flux<Employee> students) {
        return students.collectList().flatMap(studentList -> {
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                Workbook wb = new Workbook(out, "MyApp", "1.0");
                int sheetIndex = 0;
                Worksheet ws = wb.newWorksheet("Students" + (sheetIndex + 1));

                int rowIndex = 0;

                // Write headers to the first sheet
                writeHeaders(ws, rowIndex++);

                for (Employee student : studentList) {
                    if (rowIndex == MAX_ROWS_PER_SHEET) {
                        // Move to the next sheet
                        sheetIndex++;
                        ws = wb.newWorksheet("Students" + (sheetIndex + 1));
                        rowIndex = 0;
                        // Write headers to the new sheet
                        writeHeaders(ws, rowIndex++);
                    }

                    // Write student data
                    writeStudentData(ws, rowIndex++, student);
                }

                wb.finish();
                return Mono.just(new ByteArrayInputStream(out.toByteArray()));
            } catch (IOException e) {
                return Mono.error(e);
            }
        });
    }

    private static void writeHeaders(Worksheet ws, int rowIndex) {
        ws.value(rowIndex, 0, "ID");
        ws.value(rowIndex, 1, "Name");
        ws.value(rowIndex, 2, "Age");
        // Add more headers if needed
    }

    private static void writeStudentData(Worksheet ws, int rowIndex, Employee student) {
        ws.value(rowIndex, 0, student.getId());
        ws.value(rowIndex, 1, student.getName());
        ws.value(rowIndex, 2, student.getDepartment());
        // Add more student fields if needed
    }

}
