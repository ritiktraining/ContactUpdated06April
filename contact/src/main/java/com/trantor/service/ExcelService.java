package com.trantor.service;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.http.ResponseEntity;

import com.trantor.entity.Contact;

public interface ExcelService {
	public void exportData(ResponseEntity responseEntity) throws IOException;

	List<Contact> listAll(HttpServletResponse response);

	public List<Contact> uploadAll();
}
