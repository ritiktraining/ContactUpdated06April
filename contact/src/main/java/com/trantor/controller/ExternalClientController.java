package com.trantor.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.trantor.dto.ContactDto;
import com.trantor.service.ExternalClientService;

@RestController
public class ExternalClientController {

	@Autowired
	private ExternalClientService externalClientService;

	@GetMapping("/data")
	public ResponseEntity<String> fetchAll(@RequestParam("flag") Integer flag) {
		return externalClientService.fetchAll(flag);
	}

	@PostMapping("/data")
	public ResponseEntity<String> postData(@RequestParam("flag") int flag, @RequestBody ContactDto contactDto) {
		return externalClientService.postData(flag, contactDto);
	}
	
	@DeleteMapping("/data/{id}")
	public ResponseEntity<String> deleteData(@RequestParam("flag") int flag,@PathVariable Long id) {
		return externalClientService.deleteData(flag, id);
	}
	
}
