package com.trantor.service;

import java.util.List;
import java.util.stream.Collectors;

import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.trantor.dto.ContactDto;
import com.trantor.entity.Contact;
import com.trantor.repository.ContactRepository;

@Service
public class ContactServiceImpl implements ContactService {

	@Autowired
	ContactRepository contactRepository;

	@Autowired
	private ModelMapper modelMapper;

	@Override
	public ContactDto saveContact(ContactDto contactDto) {

		Contact contact = modelMapper.map(contactDto, Contact.class);
		Contact saveContact = contactRepository.save(contact);
		return modelMapper.map(saveContact, ContactDto.class);

	}

	@Override
	public List<ContactDto> fetchContactList() {
		return contactRepository.findAll().stream().map(contact -> modelMapper.map(contact, ContactDto.class))
				.collect(Collectors.toList());
	}

	@Override
	public ContactDto fetchbyfirstName(String firstName) {
		Contact fetchbyfirstName = contactRepository.findByFirstName(firstName);

		return modelMapper.map(fetchbyfirstName, ContactDto.class);
	}

	@Override
	public Contact deleteContact(Long contactId) {
		Contact contact = contactRepository.findById(contactId).orElseThrow();

		contact.setIsActive("N");

		return contactRepository.save(contact);
	}

}
