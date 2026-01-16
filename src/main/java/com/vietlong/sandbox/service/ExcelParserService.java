package com.vietlong.sandbox.service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.text.Normalizer;
import java.util.*;
import java.util.regex.Pattern;

@Service
public class ExcelParserService {

    public List<Map<String, Object>> parseToJson(MultipartFile file) {
        List<Map<String, Object>> jsonDataList = new ArrayList<>();

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet firstSheet = workbook.getSheetAt(0);

            Row headerRow = firstSheet.getRow(0);
            if (headerRow == null) {
                throw new RuntimeException("File Excel rỗng hoặc không có Header!");
            }

            Map<Integer, String> columnHeaderMap = new HashMap<>();

            int lastCellIndex = headerRow.getLastCellNum();
            for (int cellIndex = 0; cellIndex < lastCellIndex; cellIndex++){
                String originalHeader = headerRow.getCell(cellIndex).getStringCellValue();
                String normalizedKey = formatHeaderKey(originalHeader);
                columnHeaderMap.put(cellIndex, normalizedKey);
            }

            int lastRowIndex = firstSheet.getLastRowNum();
            for (int rowIndex = 1; rowIndex <= lastRowIndex; rowIndex++) {
                Row dataRow = firstSheet.getRow(rowIndex);
                if (isEmptyRow(dataRow)) continue;
                Map<String, Object> jsonObject = new LinkedHashMap<>();
                for (Map.Entry<Integer, String> headerEntry : columnHeaderMap.entrySet()) {
                    int columnIndex = headerEntry.getKey();
                    String jsonKey = headerEntry.getValue();
                    Cell dataCell = dataRow.getCell(columnIndex);
                    Object cellValue = getCellValue(dataCell);
                    jsonObject.put(jsonKey, cellValue);
                }
                jsonDataList.add(jsonObject);
            }

        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Lỗi xử lý file Excel: " + e.getMessage());
        }

        return jsonDataList;
    }

    public List<Map<String, Object>> parseToJson(MultipartFile file, List<Integer> columnIndexes, List<String> customKeys) {
        List<Map<String, Object>> jsonDataList = new ArrayList<>();

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet firstSheet = workbook.getSheetAt(0);

            Row headerRow = firstSheet.getRow(0);
            if (headerRow == null) {
                throw new RuntimeException("File Excel rỗng hoặc không có Header!");
            }

            Map<Integer, String> columnHeaderMap = new HashMap<>();
            
            // Kiểm tra có dùng customKeys hay không
            boolean useCustomKeys = customKeys != null && !customKeys.isEmpty();

            for (int i = 0; i < columnIndexes.size(); i++) {
                int excelColumnIndex = columnIndexes.get(i);
                String jsonKey;
                
                if (useCustomKeys) {
                    // Dùng custom key do client cung cấp
                    jsonKey = customKeys.get(i);
                } else {
                    // Lấy từ header Excel và normalize
                    Cell headerCell = headerRow.getCell(excelColumnIndex);
                    if (headerCell == null) {
                        jsonKey = "UNKNOWN_COL_" + excelColumnIndex;
                    } else {
                        String originalHeader = headerCell.getStringCellValue();
                        jsonKey = formatHeaderKey(originalHeader);
                    }
                }
                
                // Map Excel column index với JSON key
                columnHeaderMap.put(excelColumnIndex, jsonKey);
            }

            int lastRowIndex = firstSheet.getLastRowNum();
            for (int rowIndex = 1; rowIndex <= lastRowIndex; rowIndex++) {
                Row dataRow = firstSheet.getRow(rowIndex);
                if (isEmptyRow(dataRow)) continue;
                Map<String, Object> jsonObject = new LinkedHashMap<>();
                for (Map.Entry<Integer, String> headerEntry : columnHeaderMap.entrySet()) {
                    int columnIndex = headerEntry.getKey();
                    String jsonKey = headerEntry.getValue();
                    Cell dataCell = dataRow.getCell(columnIndex);
                    Object cellValue = getCellValue(dataCell);
                    jsonObject.put(jsonKey, cellValue);
                }
                jsonDataList.add(jsonObject);
            }

        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Lỗi xử lý file Excel: " + e.getMessage());
        }

        return jsonDataList;
    }

    private String formatHeaderKey(String header) {
        if (header == null) return "UNKNOWN_COL";
        String normalized = header.trim().toUpperCase();
        normalized = removeAccents(normalized);
        return normalized.replaceAll("\\s+", "_");
    }


    private String removeAccents(String str) {
        String nfdNormalizedString = Normalizer.normalize(str, Normalizer.Form.NFD);
        Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
        return pattern.matcher(nfdNormalizedString).replaceAll("").replace('đ', 'd').replace('Đ', 'D');
    }


    private Object getCellValue(Cell cell) {
        if (cell == null) return null;

        switch (cell.getCellType()) {
            case STRING:
                String stringValue = cell.getStringCellValue().trim();
                return stringValue.isEmpty() ? null : stringValue;

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == (long) numericValue) {
                        return (long) numericValue;
                    }
                    return numericValue;
                }

            case BOOLEAN:
                return cell.getBooleanCellValue();

            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return cell.getNumericCellValue();
                }

            default:
                return null;
        }
    }

    private boolean isEmptyRow(Row row) {
        if (row == null) return true;
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

}
