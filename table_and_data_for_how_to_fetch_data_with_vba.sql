SQL Server table and data

CREATE TABLE [dbo].[Sales](
	[Product] [varchar](50) NULL,
	[OrderDate] [date] NULL,
	[Quantity] [int] NULL
) ON [PRIMARY]
GO


insert into sales
		(Product, OrderDate,Quantity)
values
('Product A',	'2020-01-01',	100),
('Product A',	'2020-02-01',	200),
('Product A',	'2020-03-01',	150),
('Product A',	'2020-04-01',	250),
('Product A',	'2021-01-01',	120),
('Product A',	'2021-02-01',	220),
('Product A',	'2021-03-01',	130),
('Product A',	'2021-04-01',	240),
('Product A',	'2022-01-01',	130),
('Product A',	'2022-02-01',	250),
('Product A',	'2022-03-01',	140),
('Product A',	'2022-04-01',	260),
('Product B',	'2020-01-01',	150),
('Product B',	'2020-02-01',	250),
('Product B',	'2020-03-01',	200),
('Product B',	'2020-04-01',	300),
('Product B',	'2021-01-01',	160),
('Product B',	'2021-02-01',	260),
('Product B',	'2021-03-01',	170),
('Product B',	'2021-04-01',	280),
('Product B',	'2022-01-01',	170),
('Product B',	'2022-02-01',	290),
('Product B',	'2022-03-01',	180),
('Product B',	'2022-04-01',	310)















