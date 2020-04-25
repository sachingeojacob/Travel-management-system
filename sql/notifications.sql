use master
create table notifications(notification_id int primary key,purpose varchar(100),notification_date varchar(20),notification_time_H int,notification_time_M int,am_pm varchar(15),venue varchar(50))
SELECT * FROM notifications 

drop table notifications 