USE [Hospital]
Go
create table Medicinal_stores
		(Name_of_the_Medicine char(20) not null,
		 Manufacture_date datetime not null,
		 Expiry_date datetime not null,
		 Price integer not null,
		 primary key(Name_of_the_Medicine));