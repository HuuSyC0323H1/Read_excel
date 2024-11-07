package com.example.springtest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.Normalizer;
import java.util.regex.Pattern;

@SpringBootApplication
public class SpringtestApplication {
    private static final Pattern DIACRITICS_PATTERN = Pattern.compile("\\p{M}");
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\Huu Sy\\Downloads\\List danh sach.xlsx";
        String excelFilePathResult = "C:\\Users\\Huu Sy\\Downloads\\Exceedra - Project NESCAFE TET - TT sampling  redemption - TD.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis);
             FileInputStream fisResult = new FileInputStream(new File(excelFilePathResult));
             Workbook workbookResult = new XSSFWorkbook(fisResult)) {

            Sheet sheet = workbook.getSheetAt(1); // Sheet chứa dữ liệu nguồn
            Sheet resultSheet = workbookResult.getSheetAt(1); // Sheet chứa kết quả

            int customerKeyIndex = -1;
            int customerInfoIndex = -1;
            int addressIndex = -1;

            // Xác định chỉ số cột CustomerKey, CustomerInfo, Address trong file nguồn
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                String cellValue = cell.getStringCellValue();
                if ("CustomerKey".equalsIgnoreCase(cellValue)) {
                    customerKeyIndex = cell.getColumnIndex();
                } else if ("CustomerInfo".equalsIgnoreCase(cellValue)) {
                    customerInfoIndex = cell.getColumnIndex();
                } else if ("Address".equalsIgnoreCase(cellValue)) {
                    addressIndex = cell.getColumnIndex();
                }
            }

            if (customerKeyIndex == -1 || customerInfoIndex == -1 || addressIndex == -1) {
                System.out.println("Không tìm thấy một trong các cột: CustomerKey, CustomerInfo, Address trong file nguồn.");
                return;
            }

            int resultRowIndex = 2; // Bắt đầu ghi từ dòng thứ 2 (bỏ qua tiêu đề)

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Duyệt qua từng dòng dữ liệu
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Lấy dữ liệu từ các cột cần thiết
                String customerKey = getCellValueAsString(row.getCell(customerKeyIndex));
                String customerInfo = getCellValueAsString(row.getCell(customerInfoIndex));
                String address = getCellValueAsStringAddress(row.getCell(addressIndex));

                // Tạo dòng mới trong sheet kết quả
                Row resultRow = resultSheet.getRow(resultRowIndex);
                if (resultRow == null) {
                    resultRow = resultSheet.createRow(resultRowIndex);
                }

                // Ghi dữ liệu vào các cột tương ứng trong sheet kết quả
                resultRow.createCell(3).setCellValue(customerKey); // Cột Customer ID (ID address)
                resultRow.createCell(4).setCellValue(customerInfo); // Cột Location name
                resultRow.createCell(5).setCellValue(address);      // Cột Location Address

                resultRowIndex++;
            }

            // Ghi workbook kết quả ra file Excel
            try (FileOutputStream fos = new FileOutputStream(new File(excelFilePathResult))) {
                workbookResult.write(fos);
            }

            System.out.println("Dữ liệu đã được sao chép thành công!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Hàm phụ để lấy giá trị ô dưới dạng chuỗi
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                String cellValue = cell.getStringCellValue();

                int firstUnderscoreIndex = cellValue.indexOf("_");
                if (firstUnderscoreIndex != -1 && cellValue.matches("^\\d.*")) {
                    cellValue = cellValue.substring(firstUnderscoreIndex + 1);
                }
                return removeDiacritics(cellValue);
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
    private static String getCellValueAsStringAddress(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                String cellValue = cell.getStringCellValue();

                int firstUnderscoreIndex = cellValue.indexOf("_");
                int secondUnderscoreIndex = cellValue.indexOf("_", firstUnderscoreIndex + 1);
                if (secondUnderscoreIndex != -1 && cellValue.matches("^\\d.*")) {
                    cellValue = cellValue.substring(secondUnderscoreIndex + 1);
                }
                return removeDiacritics(cellValue);
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return formatNumericCell(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static String formatNumericCell(double value) {
        BigDecimal bigDecimalValue = new BigDecimal(value);
        return bigDecimalValue.toPlainString();
    }

    private static String removeDiacritics(String input) {
        if (input == null) {
            return "";
        }

        String normalized = Normalizer.normalize(input, Normalizer.Form.NFD);
        String withoutDiacritics = DIACRITICS_PATTERN.matcher(normalized).replaceAll("");

        withoutDiacritics = withoutDiacritics.replace("Đ", "D").replace("đ", "d");

        return withoutDiacritics;
    }
}