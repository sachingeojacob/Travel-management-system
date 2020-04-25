use master
create table booking(booking_id int primary key,customer_id int references customer(customer_id)on delete cascade on update cascade,package_id int references packages(package_id)on delete cascade on update cascade,customer_name varchar(50),mobile numeric(18,0),state varchar(50),hotel varchar(50),number_of_persons int,boarding_date varchar(10),email_id varchar(20),spot varchar(30));

select * from booking
drop table booking

create table booking(booking_id int primary key,customer_id int references customer(customer_id)on delete cascade on update cascade,package_id int references packages(package_id)on delete cascade on update cascade,customer_name varchar(50),mobile varchar(50),state varchar(50),hotel varchar(50),number_of_persons varchar(10),boarding_date varchar(10),email_id varchar(40),spot varchar(30));


