package com.trantor.service;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.trantor.entity.Contact;
import com.trantor.entity.Mobile;

@Service
public class UserExcelExporterService {
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<Contact> listUsers;

	public UserExcelExporterService(List<Contact> listUsers) {
		this.listUsers = listUsers;
		workbook = new XSSFWorkbook();
	}

	private void writeHeaderLine() {
		sheet = workbook.createSheet("Users");

		Row row = sheet.createRow(0);

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(16);
		style.setFont(font);
		style.setFillBackgroundColor(IndexedColors.BLUE_GREY.getIndex());
		style.setFillPattern(FillPatternType.BIG_SPOTS);

		createCell(row, 0, "Contact ID", style);
		createCell(row, 1, "First Name", style);
		createCell(row, 2, "Last Name", style);
		createCell(row, 3, "Email Address", style);
		createCell(row, 4, "Is Active", style);
		createCell(row, 5, "Created Date", style);
		createCell(row, 6, "Created By", style);
		createCell(row, 7, "Updated Date", style);
		createCell(row, 8, "Updated By", style);

		// mobile
		createCell(row, 9, "MobileId", style);
		createCell(row, 10, "Mobile Number", style);
		createCell(row, 11, "Country Cd", style);
		createCell(row, 12, "Type", style);
		createCell(row, 13, "Created Date", style);
		createCell(row, 14, "Created By", style);
		createCell(row, 15, "Updated Date", style);
		createCell(row, 16, "Updated By", style);

	}

	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}

	private void writeDataLines() {
		int rowCount = 1;

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);

		for (Contact user : listUsers) {

			CellStyle cellStyle = setColor(user);
			style = cellStyle;

			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;

			createCell(row, columnCount++, user.getContactId(), style);
			createCell(row, columnCount++, user.getFirstName(), style);
			createCell(row, columnCount++, user.getLastName(), style);
			createCell(row, columnCount++, user.getEmailAddress(), style);
			createCell(row, columnCount++, user.getIsActive(), style);
			createCell(row, columnCount++, user.getCreatedDate().toString(), style);
			createCell(row, columnCount++, user.getCreatedBy(), style);
			createCell(row, columnCount++, user.getUpdatedDate().toString(), style);
			createCell(row, columnCount++, user.getUpdatedBy(), style);

			List<Mobile> mobile = user.getMobile();

			for (Mobile mobileModel : mobile) {

				createCell(row, columnCount++, mobileModel.getMobileId(), style);

				String mno = new String();
				mno = String.valueOf(mobileModel.getMobileNumber());

				createCell(row, columnCount++, mno.toString(), style);

				createCell(row, columnCount++, mobileModel.getCountryCd(), style);
				createCell(row, columnCount++, mobileModel.getType(), style);

				createCell(row, columnCount++, mobileModel.getCreatedDate().toString(), style);
				createCell(row, columnCount++, mobileModel.getCreatedBy(), style);

				createCell(row, columnCount++, mobileModel.getUpdatedDate().toString(), style);
				createCell(row, columnCount++, mobileModel.getUpdatedBy(), style);

			}
		}
	}

	CellStyle setColor(Contact contact) {
		CellStyle cellStyle = workbook.createCellStyle();

		if (contact.getContactId() % 2 == 0) {
			cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
			cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
		} else {
			cellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
			cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
		}

		return cellStyle;
	}

	public void export(HttpServletResponse response) throws IOException {
		writeHeaderLine();
		writeDataLines();

		ServletOutputStream outputStream = response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();

		outputStream.close();

	}
}
