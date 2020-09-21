USE [Hospital]
Go
create table patient
		(patient_id integer not null,
		 patient_name char(30) not null,
                 patient_details varchar(40) not null,
	         ward_no integer not null,
		 primary key(patient_id));

