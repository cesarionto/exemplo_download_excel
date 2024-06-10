package com.cesario.excel.api.exemplo_download_excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {
    public byte[] generateExcel() throws IOException {
        // Cria uma nova planilha
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Exemplo");

        // Adiciona cabeçalhos
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("Nome");
        headerRow.createCell(2).setCellValue("Idade");

        // Adiciona alguns dados de exemplo
        Object[][] data = {
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Carol", 40 }
        };

        int rowNum = 1;
        for (Object[] rowData : data) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue((Integer) rowData[0]);
            row.createCell(1).setCellValue((String) rowData[1]);
            row.createCell(2).setCellValue((Integer) rowData[2]);
        }

        // Escreve o conteúdo da planilha em um byte array
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        return outputStream.toByteArray();
    }
}
