USE [Hospital]
Go
create table Staff_info
			(Staff_id integer,
			 Staff_name char(30) not null,
			 Department char(20) not null,
			 Post char(20) not null,
			 Salary integer not null,
			 primary key(Staff_id));