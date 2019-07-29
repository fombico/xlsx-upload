package com.fombico.xlsxupload;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
public class ApiController {

    @PostMapping("/upload")
    public void upload(@RequestPart MultipartFile file) throws IOException {
        Workbook workbook = WorkbookFactory.create(file.getInputStream());
        DataFormatter formatter = new DataFormatter();

        for (Sheet sheet : workbook) {
            System.out.println("Sheet " + sheet.getSheetName() + " has " + sheet.getLastRowNum() + " rows");

            for (Row row : sheet) {
                List<String> cellValues = new ArrayList<>();
                for (Cell cell : row) {
                    cellValues.add(formatter.formatCellValue(cell));
                }

                if (foundLastRow(cellValues)) {
                    break;
                }

                System.out.println(sheet.getSheetName() + " - row " + row.getRowNum() + " has values: " + cellValues);
            }
            System.out.println("=========================================");
        }
    }

    private boolean foundLastRow(List<String> cellValues) {
        int totalCharCount = 0;
        for (String value : cellValues) {
            totalCharCount += value.length();
        }
        return totalCharCount == 0;
    }
}
