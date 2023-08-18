package com.hitss.exceltojson.service;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {

    public List<HashMap<String, Object>> convertExcelToJson(String fileLocation) throws Exception {

        List<HashMap<String, Object>> array = new ArrayList<>();

        FileInputStream file = new FileInputStream(new File(fileLocation));
        try (Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            List<String> columnNames = getColumnNames(sheet);

            for (Row row : sheet) {
                int i = 0;
                if (row.getRowNum() == 0) {
                    continue;
                }
                HashMap<String, Object> object = new HashMap<>();
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            object.put(columnNames.get(i),
                                    cell.getStringCellValue().equalsIgnoreCase("null") ? null : cell.getStringCellValue());
                            i++;
                            break;
                        case NUMERIC:
                            object.put(columnNames.get(i), (int) cell.getNumericCellValue());
                            i++;
                            break;
                        case BOOLEAN:
                            object.put(columnNames.get(i),
                                    cell.getBooleanCellValue());
                            i++;
                        case FORMULA:
                            object.put(columnNames.get(i),
                                    cell.getCellFormula());
                            i++;
                            break;
                        default:
                            break;
                    }

                }
                array.add(object);

            }
        }catch (Exception e) {
           e.printStackTrace();
           throw new Exception("Erro ao converter excel para json");
        }

        return array;
    }

    private List<String> getColumnNames(Sheet sheet) {
        List<String> columnNames = new ArrayList<>();

        for (Cell cell : sheet.getRow(0)) {
            columnNames.add(cell.getStringCellValue());
        }

        return columnNames;
    }
}
