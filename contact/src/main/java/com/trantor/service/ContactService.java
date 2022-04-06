package com.trantor.service;

import java.util.List;

import com.trantor.dto.ContactDto;
import com.trantor.entity.Contact;

public interface ContactService {

	public List<ContactDto> fetchContactList();

	public ContactDto fetchbyfirstName(String firstName);

	ContactDto saveContact(ContactDto contactDto);
	
	public Contact deleteContact(Long contactId);

}
