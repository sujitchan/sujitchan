package com.example.ImageConvert;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.*;
public class ExcelLoader {
	 public static void main(String[] args) throws IOException {
	        // Load the Excel workbook
	        FileInputStream excelFile = new FileInputStream(new File("C:/Users/AGT-18/Downloads/sample_vendors.xlsx"));
	        Workbook workbook = new XSSFWorkbook(excelFile);
	        Sheet sheet = workbook.getSheetAt(0);

	        // Create a new PDF document
	        PDDocument pdfDocument = new PDDocument();
	        PDPage pdfPage = new PDPage();
	        pdfDocument.addPage(pdfPage);

	        // Create a PDF content stream
	        PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, pdfPage);
	        contentStream.setFont(PDType1Font.HELVETICA, 12);

	        // Iterate through the Excel rows and columns
	        int yPosition = 750;  // Starting y position for the first row in the PDF
	        for (Row row : sheet) {
	            int xPosition = 50;  // Starting x position for the first column in the PDF
	            for (Cell cell : row) {
	                String cellValue = cell.toString();
	                contentStream.beginText();
	                contentStream.newLineAtOffset(xPosition, yPosition);
	                contentStream.showText(cellValue);
	                contentStream.endText();
	                xPosition += 100;  // Adjust the x position for the next cell
	            }
	            yPosition -= 20;  // Adjust the y position for the next row
	        }

	        // Close the content stream
	        contentStream.close();

	        // Save the PDF document
	        pdfDocument.save("example.pdf");
	        pdfDocument.close();
	        workbook.close();

	        System.out.println("Excel file converted to PDF successfully!");
	    }
	}
