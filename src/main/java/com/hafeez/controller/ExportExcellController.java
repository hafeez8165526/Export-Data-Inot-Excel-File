package com.hafeez.controller;

import java.io.FileNotFoundException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.hafeez.service.PopulateExcelService;


@RestController
@CrossOrigin(exposedHeaders = {"status","filename"})

public class ExportExcellController {

	
	@Autowired
	PopulateExcelService ser;
	
	
	@GetMapping("/getExcel")
	public ResponseEntity<Resource> getExcel(){
		try {
			return ser.populateExcel();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
}
