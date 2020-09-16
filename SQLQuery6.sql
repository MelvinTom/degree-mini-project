use project;
create table billing1( fname varchar(10),lname varchar(10),id int primary key,hname varchar(10),town varchar(10),pin int,mob int,email varchar(10));
create table billing2(model varchar(10),cc int,variant varchar(10),colour varchar(10),exprice int,insurance int,advance int,accessories int,total varchar(10));
select * from billing2;
drop table billing2;