
create table t_PO
(
	ID int not null primary key identity,
	PO varchar(255) not null unique, --采购订单号
	PUser varchar(255), --采购员
	SupplierNO varchar(255), --供应商代码

	PayTerm varchar(255), --付款条款代码
	PayTermStype varchar(255), --付款条款类型
	PayTermDay int, --付款条款天数
	PayTermDesc  varchar(255), --付款条款描述
	PayTermAuto varchar(255), --付款条款对应流程

	PaymentCurrency varchar(255),--支付币种
	PaymentStype varchar(255),--支付类型
	PGroup varchar(255),--采购组
	Class2 varchar(255),--二级分类
	SupplierShortName  varchar(255),--供应商简称 （不检查完整性）
	CompanyNO varchar(255),--公司代码
	BankOf  varchar(255),--对公对私
	DeliveryDate date,--交货日期
	InvoiceDate date,--发票提供日期
	IsLinkPR varchar(255),--是否关联资本性支出请购流程
	PurchaseStype varchar(255),--分期类型
	SWIFTCode varchar(255),--银行国家代码
	FilePath varchar(255),--相关附件
	IsInfrastructure varchar(255),--是否基建类
	IsExistPO varchar(255),--是否有采购订单号
	PayFinish date,--付款要求完成时间
	ScadaDateTime datetime not null default(getdate()),

	IsIntact int not null default(0),	--完整性：0=未完整，1=完整	
	RequestID varchar(255),				--OA保存时触发的RequestID（IsRequest=0或者IsRequest=1）
	IsRequest int not null default(-1),	--未操作OA=-1；保存OA=0；提交=1;其他异常不再执行=2
	ResInfo varchar(max)				--结果信息	
);
/*
采购订单号（款到发货）
申请人
供应商代码	(名称自动)

支付币种
付款类型
二级分类
供应商简称(代码)

付款公司
对公对私
预计金税发票提供时间
是否关联资本性支出请购流程
分期类型
收款人国家/地区
是否基建类
是否有采购订单号
付款要求完成时间
付款日期说明

相关附件
*/

create database db_RpaZF003
use db_RpaZF003
drop table t_PO
create table t_PO
(
	---SE11+EKKO,日期、状态G，仅记录Y020-----------------------------------------------
	ID int not null primary key identity,
	PO varchar(255) not null unique,--采购订单号  
	PayCode varchar(255),			--付款条款代码（直接固定Y020，取消付款条款代码配置表）

	---ME23N------------------------------------------------------------------------------
	PUser varchar(255),				--采购员描述
	Supplier varchar(255),			--供应商
	SupplierCode varchar(255),		--供应商代码
	SupplierShortName varchar(255),	--供应商简称
	Currency varchar(255),			--货币
	PGroup varchar(255),			--采购组	
	CompanyCode varchar(255),		--公司代码
	DeliveryDate nvarchar(max),		--交货日期(json)

	---ZMMR010------------------------------------------------------------------------------
	SWIFTCode varchar(255),		--银行国家代码（供应商代码+公司代码）

	---ZMMF007------------------------------------------------------------------------------
	FilePath varchar(255),		--相关附件 (工厂+采购订单号)

	----------------------------------------------------------------------------------------
	ScadaDateTime datetime not null default(getdate()),

	-----计算值--------------------------------------------------------------------------------------------
	Class2 varchar(255),			--二级分类（采购组+二级分类配置表）
	InvoiceDate date,				--预计金税发票提供时间 （交货日期<20号为交货月月末，否则为交货月次月月末）
	PayFinish date,					--付款要求完成时间（交期日期的前3天）
	PayDateExplained varchar(255),	--付款日期说明（款到发货）

	-----固定值--------------------------------------------------------------------------------
	PayStype varchar(255) not null default('预付材料款'),	--支付类型	
	BankOf  varchar(255) not null default('对公'),			--对公对私
	IsLinkPR varchar(255)not null default('否'),			--是否关联资本性支出请购流程
	PurchaseStype varchar(255) not null default('款到发货'),--分期类型Y020
	IsInfrastructure varchar(255) not null default('否'),	--是否基建类
	IsExistPO varchar(255) not null default('是'),			--是否有采购订单号

	----检查------------------------------------------------------------------------------------
	IsScadaFinsh int not null default(0),	--采集完毕：0=未完整，1=完整	
	IsIntact int not null default(0),		--检查完整：0=未完整，1=完整	

	----OA------------------------------------------------------------------------------------
	RequestID varchar(255),				--OA保存时触发的RequestID（IsRequest=0或者IsRequest=1）
	IsRequest int not null default(-1),	--未操作OA=-1；保存OA=0；提交=1;其他异常不再执行=2
	ResInfo nvarchar(max)				--结果信息

);


select * from t_PO where IsScadaFinsh=0 and IsRequest=-1

select * from t_PO where IsScadaFinsh=1 and IsRequest=-1

select * from t_PO where   IsRequest<=0 and (SWIFTCode is null or SWIFTCode='')

select * from t_PO where   IsRequest<=0 and SWIFTCode is null 





