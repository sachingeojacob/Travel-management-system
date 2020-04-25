use master
create table login_page(user_name varchar(20)primary key not null,employee_id int references employee_management(employee_id)on delete cascade on update cascade,password varchar(20),account_type varchar(30))

insert into login_page(user_name,employee_id,password,account_type) values('admin',1,'admin','ADMIN')

select * from login_page
drop table login_page
select * from employee_management
