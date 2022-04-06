package com.trantor.controller;

import java.util.List;

import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.trantor.dto.ContactDto;
import com.trantor.entity.Contact;
import com.trantor.service.ContactServiceImpl;

@RestController
public class ContactController {

	@Autowired
	private ContactServiceImpl contactServiceImpl;

	@Autowired
	private ModelMapper modelMapper;

	@PostMapping("/contact")
	public ContactDto saveContact(@RequestBody ContactDto contactDto) {

		return contactServiceImpl.saveContact(contactDto);

	}

	@GetMapping("/contact")
	public List<ContactDto> fetchContctList() {
		return contactServiceImpl.fetchContactList();

	}

	@GetMapping("/contact/name/{name}")
	public ContactDto fetchbyfirstName(@PathVariable("name") String firstName) {
		return contactServiceImpl.fetchbyfirstName(firstName);
	}

	@DeleteMapping("/contact/{id}")
	public ContactDto deleteById(@PathVariable("id") Long contactId) {

		Contact deleteContact = contactServiceImpl.deleteContact(contactId);

		return modelMapper.map(deleteContact, ContactDto.class);

	}

}
