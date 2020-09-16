use project;
 create table spare(id int not null,name varchar(10) not null,cc int not null,quantity int not null,model varchar(10) not null,colour varchar(10) not null,price int not null);
create table used(years int not null,name varchar(10) not null,man varchar(10) not null,rc varchar(10) not null,ins varchar(10) not null,pol varchar(10) not null,body varchar(10) not null,variant varchar(10) not null,trans varchar(10) not null,nos int not null,clutch varchar(10) not null,len int not null,height int not null,front int not null,rear int not null);
select * from used;