package com.trantor.entity;

import java.util.Date;
import java.util.List;

import javax.persistence.CascadeType;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.OneToMany;
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
@Table(name = "CONTACT")
@Data
@ToString
@AllArgsConstructor
@NoArgsConstructor
public class Contact {

	@Id
	@GeneratedValue(strategy = GenerationType.AUTO)
	private Integer contactId;

	@NotNull(message = "This field cannot be Empty")
	private String firstName;

	@NotNull(message = "This field cannot be Empty")
	private String lastName;

	@NotNull(message = "This field cannot be Empty")
	private String emailAddress;

	@CreatedBy
	@NotNull(message = "This field cannot be Empty")
	private String createdBy;

	@CreationTimestamp
	private Date createdDate;

	@LastModifiedBy
	private String updatedBy;

	@UpdateTimestamp
	private Date updatedDate;

	private String isActive;

	@OneToMany(targetEntity = Mobile.class, cascade = CascadeType.ALL)
	@JoinColumn(name = "contactId", referencedColumnName = "contactId")
	private List<Mobile> mobile;

}
