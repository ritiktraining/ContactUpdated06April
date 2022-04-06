package com.trantor.dto;

import java.util.Date;
import java.util.List;

import com.trantor.entity.Mobile;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Data
@ToString
@AllArgsConstructor
@NoArgsConstructor
public class ContactDto {

	private Integer contactId;

//	@JsonProperty("abc")
	private String firstName;

//	@JsonProperty("XYZ")
	private String lastName;
	private String emailAddress;
	private String createdBy;
	private Date createdDate;
	private String updatedBy;
	private Date updatedDate;
	private String isActive;

	private List<Mobile> mobile;

}
