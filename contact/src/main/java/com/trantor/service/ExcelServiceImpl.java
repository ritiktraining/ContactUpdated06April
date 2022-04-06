package com.trantor.service;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import javax.transaction.Transactional;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import com.trantor.entity.Contact;
import com.trantor.repository.ContactRepository;

@Service
@Transactional
public class ExcelServiceImpl implements ExcelService {

	private static final String CSV_FILE_LOCATION = "C:/Users/ritik.kumar/Downloads/ContactsDataExcel.xlsx";
	private static final org.slf4j.Logger logger = LoggerFactory.getLogger(ExcelServiceImpl.class);
	private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
	@Autowired
	private UserExcelExporterService userExcelExporterService;

	@Autowired
	private ContactRepository contactRepo;

	public void exportData(ResponseEntity responseEntity) throws IOException {
		userExcelExporterService.export((HttpServletResponse) responseEntity);
	}

	public List<Contact> listAll(HttpServletResponse response) {

		response.setContentType("application/octet-stream");

		String headerKey = "Content-Disposition";
		String headerValue = "attachment; filename=Contacts_" + "DataExcel" + ".xlsx";
		response.setHeader(headerKey, headerValue);

		List<Contact> all = contactRepo.findAll();

		UserExcelExporterService excelExporter = new UserExcelExporterService(all);

		try {
			excelExporter.export(response);
		} catch (IOException e) {
			e.printStackTrace();
		}

		return all;
	}

	@Transactional
	public List<Contact> uploadAll() {

		List<Contact> courses = new ArrayList<>();

		Workbook workbook = null;
		try {
			// Creating a Workbook from an Excel file (.xls or .xlsx)
			workbook = WorkbookFactory.create(new File(CSV_FILE_LOCATION));

			// Retrieving the number of sheets in the Workbook
			logger.info("Number of sheets: ", workbook.getNumberOfSheets());
			// Print all sheets name
			workbook.forEach(sheet -> {
				logger.info(" => " + sheet.getSheetName());

				// Create a DataFormatter to format and get each cell's value as String
				DataFormatter dataFormatter = new DataFormatter();

				// loop through all rows and columns and create Course object
				int index = 0;
				for (Row row : sheet) {
					if (index++ == 0)
						continue;

					Contact course = new Contact();

					course.setFirstName(dataFormatter.formatCellValue(row.getCell(1)));
					course.setLastName(dataFormatter.formatCellValue(row.getCell(2)));
					course.setEmailAddress(dataFormatter.formatCellValue(row.getCell(3)));
					course.setIsActive(String.valueOf(dataFormatter.formatCellValue(row.getCell(4))));
					course.setCreatedBy(dataFormatter.formatCellValue(row.getCell(6)));
					courses.add(course);
				}

				contactRepo.saveAll(courses);
			});
		} catch (EncryptedDocumentException | IOException e) {
			logger.error(e.getMessage(), e);
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			}
		}

		return courses;
	}
}
