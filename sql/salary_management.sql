use master
create table salary_management(salary_id int primary key,employee_id int references employee_management(employee_id)on delete cascade on update cascade,employee_name varchar(50),basic_pay varchar(50),da varchar(50),hra varchar(50),providend_fund varchar(10),total_salary varchar(10));

select * from salary_management
drop table salary_management