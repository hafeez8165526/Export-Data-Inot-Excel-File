package com.hafeez.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;


@Service

public class PopulateExcelService {
	
	
	public ResponseEntity<Resource> populateExcel() throws FileNotFoundException {
//		W
		HSSFWorkbook workbook=new HSSFWorkbook();
		FileOutputStream out=new FileOutputStream(new File("C://Users//hafee//Documents//workspace-spring-tool-suite-4-4.20.1.RELEASE//reports//GeneratedFile.xlsx"));
		DataFormat format=workbook.createDataFormat();
		
		Sheet sheet=workbook.createSheet("New Sheet");
		Integer headerRowPos = 0;
		Integer startRowPos = 1;
		Row headerRow=sheet.createRow(headerRowPos);
		// creating one header cell
		Cell headerCell=headerRow.createCell(0);
		headerCell.setCellValue("Name");
		//creating two rows with dummy values
		//first row
		Row row1=sheet.createRow(startRowPos);
		row1.createCell(0).setCellValue("hafeez");
		startRowPos++;
		//second row
		Row row2=sheet.createRow(startRowPos);
		row2.createCell(0).setCellValue("Chotu");
		
		try {
			workbook.write(out);
			File res=new File("C://Users//hafee//Documents//workspace-spring-tool-suite-4-4.20.1.RELEASE//reports//GeneratedFile.xlsx");
			InputStreamResource resource=new InputStreamResource(new FileInputStream(res));
			HttpHeaders headers=new HttpHeaders();
			Long size=res.length();
			headers.setContentType(MediaType.parseMediaType("application/vnd.ms-excel"));
			headers.set("filename","GeneratedFile");
			headers.set("ext", "xlsx");
			headers.set("size", size.toString());
			headers.set("status", "true");
			return new ResponseEntity<>(resource, headers, HttpStatus.OK);
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
			
		
		
		
	}
}
