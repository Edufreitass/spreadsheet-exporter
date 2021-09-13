package io.redspark;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadExcel {

    public static void main(String[] args) throws IOException {

        String filePath = "products.xls";

        try {
            // Abrindo o arquivo e recuperando a planilha
            FileInputStream file = new FileInputStream(filePath);
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(0);

            List<Product> products = new ArrayList<>();

            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Descantando a primeira linha com o header
                if (row.getRowNum() == 0) {
                    continue;
                }

                Iterator<Cell> cellIterator = row.cellIterator();
                Product product = new Product();
                products.add(product);

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getColumnIndex()) {
                        case 0:
                            product.setId(((Double) cell.getNumericCellValue()).longValue());
                            break;
                        case 1:
                            product.setName(cell.getStringCellValue());
                            break;
                        case 2:
                            product.setPrice(cell.getNumericCellValue());
                            break;
                    }
                }
            }

            for (Product product : products) {
                System.out.println(product.getId() + " - " + product.getName() + " - " + product.getPrice());
            }

            file.close();
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

}
