USE [Hospital]
Go
create table department
		(Department_no integer not null,
		 Department_name char(20) not null,
		 Name_of_doctor char(30) not null,
                 Contact_no integer not null,
		 primary key(department_no));