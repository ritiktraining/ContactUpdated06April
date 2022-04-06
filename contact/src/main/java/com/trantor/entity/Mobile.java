package com.trantor.entity;

import java.util.Date;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;
import javax.validation.constraints.NotNull;

import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.UpdateTimestamp;
import org.springframework.data.annotation.CreatedBy;
import org.springframework.data.annotation.LastModifiedBy;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Entity
@Table(name = "MOBILE")
@Data
@ToString
@AllArgsConstructor
@NoArgsConstructor
public class Mobile {

	@Id
	@GeneratedValue(strategy = GenerationType.AUTO)
	private Integer mobileId;

	@NotNull(message = "This field cannot be Empty")
	private String mobileNumber;

	@NotNull(message = "This field cannot be Empty")
	private String countryCd;

	@NotNull(message = "This field cannot be Empty")
	private String type;

	@CreatedBy
	@NotNull(message = "This field cannot be Empty")
	private String createdBy;

	@CreationTimestamp
	private Date createdDate;

	@LastModifiedBy
	private String updatedBy;

	@UpdateTimestamp
	private Date updatedDate;

}
