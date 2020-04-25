use master
create table employee_management(employee_id int primary key,employee_name varchar(60),date_of_birth varchar(20),gender varchar(10),date_of_join varchar(20),mobile numeric(18,0),email_id varchar(50),basic_pay numeric(18,0),
house_name varchar(50),village varchar(50),city varchar(40),town varchar(50),pin_code int,states varchar(40),country varchar(40),pic varchar(300))


select * from employee_management
drop  employee_management