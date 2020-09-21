USE [Hospital]
Go
create table billing
		(bill_no integer not null,
		 patient_name char(30) not null,
                 total_amount int not null,
                 amount_paid int ,  
	         balance integer,
		 primary key(bill_no));