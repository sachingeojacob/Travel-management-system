use master
create table packages(package_id int primary key not null,state varchar(20),spot varchar(30),distance float,car_charges float,bus_charges float,stay_cost float,hotel1 varchar(30),hotel2 varchar(30),hotel3 varchar(30))
select *from packages
drop table packages