use learnSQL;

create table sales
(
	product		varchar(10) not null,
	orderdate	date not null,
	quantity	int not null
);

insert into sales(product, orderdate,quantity)
values
('Product A','2020-01-01',100),
('Product A','2020-02-01',200),
('Product A','2020-03-01',150),
('Product A','2020-04-01',250),
('Product A','2021-01-01',120),
('Product A','2021-02-01',220),
('Product A','2021-03-01',130),
('Product A','2021-04-01',240),
('Product B','2021-01-01',130),
('Product B','2021-02-01',120),
('Product B','2022-03-01',120),
('Product B','2022-04-01',130),
('Product B','2022-01-01',140),
('Product B','2022-02-01',140),
('Product B','2022-03-01',130),
('Product B','2022-04-01',120),
('Product B','2022-01-01',120),
('Product B','2022-02-01',130),
('Product B','2022-03-01',310);