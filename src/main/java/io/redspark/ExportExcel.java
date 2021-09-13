package io.redspark;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExportExcel {

    public static void main(String[] args) {

        // Criando o arquivo e uma planilha chamada "Product"
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Product");

        // Definindo alguns padrões de layout
        sheet.setDefaultColumnWidth(15);
        sheet.setDefaultRowHeight((short) 400);

        //Carregando os produtos
        List<Product> products = getProducts();

        int rowNum = 0;
        int cellNum = 0;
        Cell cell;
        Row row;

        //Configurando estilos de células (Cores, alinhamento, formatação, etc..)
        HSSFDataFormat numberFormat = workbook.createDataFormat();

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setAlignment(HorizontalAlignment.CENTER);
        textStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.setDataFormat(numberFormat.getFormat("#,##0.00"));
        numberStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Configurando Header
        row = sheet.createRow(rowNum++);
        cell = row.createCell(cellNum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Code");

        cell = row.createCell(cellNum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Name");

        cell = row.createCell(cellNum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Price");

        // Adicionando os dados dos produtos na planilha
        for (Product product : products) {
            row = sheet.createRow(rowNum++);
            cellNum = 0;

            cell = row.createCell(cellNum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(product.getId());

            cell = row.createCell(cellNum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(product.getName());

            cell = row.createCell(cellNum++);
            cell.setCellStyle(numberStyle);
            cell.setCellValue(product.getPrice());
        }

        try {
            //Escrevendo o arquivo em disco
            FileOutputStream out = new FileOutputStream("products.xls");
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Success!!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //Simulando uma listagem de produtos
    private static List<Product> getProducts() {
        List<Product> products = new ArrayList<>();

        products.add(new Product(1L, "Produto 1", 200.5));
        products.add(new Product(2L, "Produto 2", 1050.5));
        products.add(new Product(3L, "Produto 3", 50.0));
        products.add(new Product(4L, "Produto 4", 200.0));
        products.add(new Product(5L, "Produto 5", 450.0));
        products.add(new Product(6L, "Produto 6", 150.5));
        products.add(new Product(7L, "Produto 7", 300.99));
        products.add(new Product(8L, "Produto 8", 1000.0));
        products.add(new Product(9L, "Produto 9", 350.0));
        products.add(new Product(10L, "Produto 10", 200.0));

        return products;
    }

}
