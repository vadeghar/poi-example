package com.poi;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.springframework.boot.test.context.SpringBootTest;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;

@SpringBootTest
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class PoiExampleApplicationTests {
	private static final String FILE_NAME = "/MyFirstExcel.xlsx";

	@BeforeAll
	public void setup() {
		String path = Paths.get(".").toAbsolutePath().normalize().toString();
		File f = new File(path+FILE_NAME);
		f.delete();
	}

	@Test
	void contextLoads() {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
		Object[][] datatypes = {
				{"Datatype", "Type", "Size(in bytes)"},
				{"int", "Primitive", 2},
				{"float", "Primitive", 4},
				{"double", "Primitive", 8},
				{"char", "Primitive", 1},
				{"String", "Non-Primitive", "No fixed size"}
		};

		int rowNum = 0;
		System.out.println("Creating excel");
		Row row = sheet.createRow(rowNum++);
		int colNum = 0;
		Cell cell = row.createCell(colNum++);
		Font underlineFont = workbook.createFont();
		underlineFont.setUnderline(HSSFFont.U_SINGLE);
		underlineFont.setBold(true);
		underlineFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
		Font boldFont = workbook.createFont();
		boldFont.setBold(true);
		Font italicFont = workbook.createFont();
		italicFont.setItalic(true);
//		CellStyle style = workbook.createCellStyle();
//		style.setFont(underlineFont);
//		cell.setCellStyle(style);

		String cellText = "Maharashtra CM Eknath Shinde's faction forms <a href='/text/tezt/url'>national executive</a>, may 'show' EC it's real Sena today";
		String linkText = cellText.substring(cellText.indexOf("<a"), cellText.indexOf("</a>")+4);
		String linkAddress = linkText.substring(linkText.indexOf("'")+1, linkText.lastIndexOf("'"));
		String linkHeader = linkText.substring(linkText.indexOf(">")+1, linkText.lastIndexOf(">")-3);
		cellText = cellText.replace(linkText, linkHeader);
		System.out.println("Link: "+linkText);
		System.out.println("Link Address: "+linkAddress);
		System.out.println("Link Header: "+linkHeader);
		System.out.println("Cell Text: "+cellText);
		CreationHelper createHelper = workbook.getCreationHelper();
		XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
		link.setAddress(linkAddress);
		RichTextString richString = new XSSFRichTextString(cellText);
		richString.applyFont(0, cellText.length(), boldFont);
//		richString.applyFont(cellText.indexOf(linkHeader), cellText.length(), underlineFont);
//		richString.applyFont(cellText.indexOf(linkHeader)+linkHeader.length(), cellText.length(), boldFont);
		richString.applyFont(cellText.indexOf(linkHeader), cellText.indexOf(linkHeader)+linkHeader.length(), underlineFont);

		cell.setCellValue(richString);

//		cell.setHyperlink(link);


		try {
			String path = Paths.get(".").toAbsolutePath().normalize().toString();
			System.out.println("****** PATH: "+path);
			File f = new File(path+FILE_NAME);
			f.createNewFile();
			FileOutputStream outputStream = new FileOutputStream(path+FILE_NAME);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Done");
	}

}
