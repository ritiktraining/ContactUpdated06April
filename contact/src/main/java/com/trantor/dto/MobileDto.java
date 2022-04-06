package com.trantor.dto;

import java.util.Date;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class MobileDto {

	private Integer mobileId;

	private String mobileNumber;

	private String countryCd;
	private String type;
	private String createdBy;

	private Date createdDate;

	private String updatedBy;

	private Date updatedDate;

}
