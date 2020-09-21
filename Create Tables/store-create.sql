USE [Hospital]
Go
create table stores
		(no_of_medicines integer not null,
		 name_of_medicine char(20) not null,
		 price integer not null,
		 primary key(name_of_medicine));