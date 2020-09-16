use project ;
drop table emp1;
create table emp1(name varchar(20),id int primary key,gender varchar(10),dob varchar(20),department varchar(20),doj varchar(20),title varchar(20),bp varchar(20),hname varchar(20),village varchar(20),city varchar(20),town varchar(20),pincode varchar(20),states varchar(20),country varchar(20),email varchar(20),acdno varchar(20),branch varchar(20),mob varchar(20),alt varchar(20));
drop table salary ;
create table salary(eid int foreign key references emp1(id),department varchar(20),ename varchar(20),da int,hra int,pf int,total int)