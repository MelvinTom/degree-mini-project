use project;
create table stock1(id int primary key not null,model varchar(10) not null,variant varchar(10) not null,colour varchar(10) not null,tank varchar(10) not null,display varchar(10) not null,fuel varchar(10) not null);
create table stock2(gear int not null,clutch varchar(10)not null);
create table stock3(length int not null,breadth int not null,tiref int not null,tirer int not null);
select * from stock1