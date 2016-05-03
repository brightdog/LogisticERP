/*==============================================================*/
/* DBMS name:      Microsoft SQL Server 2008                    */
/* Created on:     2014-7-21 16:01:56                           */
/*==============================================================*/


/*==============================================================*/
/* Table: AStock                                                */
/*==============================================================*/
create table AStock (
   AStockID             int                  not null,
   OrderID              int                  null,
   DeliveryID           int                  null,
   SendOutReceiptID     int                  null,
   WarehouseID          int                  null,
   constraint PK_ASTOCK primary key nonclustered (AStockID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '目的地库存表',
   'user', @CurrentUser, 'table', 'AStock'
go

/*==============================================================*/
/* Table: City                                                  */
/*==============================================================*/
create table City (
   CityID               int                  not null,
   constraint PK_CITY primary key nonclustered (CityID)
)
go

/*==============================================================*/
/* Table: Customer                                              */
/*==============================================================*/
create table Customer (
   CustomerID           int                  not null,
   CustomerName         varchar(100)         null,
   OrderCodePrefix      varchar(5)           null,
   constraint PK_CUSTOMER primary key nonclustered (CustomerID)
)
go

/*==============================================================*/
/* Table: DStock                                                */
/*==============================================================*/
create table DStock (
   DStockID             int                  not null,
   OrderID              int                  null,
   DeliveryID           int                  null,
   PickupReceiptID      int                  null,
   WarehouseID          int                  null,
   constraint PK_DSTOCK primary key nonclustered (DStockID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '出发地库存表',
   'user', @CurrentUser, 'table', 'DStock'
go

/*==============================================================*/
/* Table: DStockToDelivery                                      */
/*==============================================================*/
create table DStockToDelivery (
   ID                   int                  not null,
   DStockID             int                  not null,
   DeliveryID           int                  not null,
   constraint PK_DSTOCKTODELIVERY primary key nonclustered (ID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '集运分拣对应关系表，库存和实际的分拣单对应关系',
   'user', @CurrentUser, 'table', 'DStockToDelivery'
go

/*==============================================================*/
/* Table: Delivery                                              */
/*==============================================================*/
create table Delivery (
   DeliveryID           int                  not null,
   SupplierID           int                  null,
   constraint PK_DELIVERY primary key nonclustered (DeliveryID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '集运分拣单表，存放的主要是干线运输的快递单信息',
   'user', @CurrentUser, 'table', 'Delivery'
go

/*==============================================================*/
/* Table: DeliveryInReceipt                                     */
/*==============================================================*/
create table DeliveryInReceipt (
   DeliveryInReceiptID  int                  not null,
   DeliveryID           int                  null,
   constraint PK_DELIVERYINRECEIPT primary key nonclustered (DeliveryInReceiptID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '快递入库单表',
   'user', @CurrentUser, 'table', 'DeliveryInReceipt'
go

/*==============================================================*/
/* Table: DeliveryOutReceipt                                    */
/*==============================================================*/
create table DeliveryOutReceipt (
   DeliveryOutReceipt_ID int                  not null,
   DeliveryID           int                  null,
   constraint PK_DELIVERYOUTRECEIPT primary key nonclustered (DeliveryOutReceipt_ID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '快递出库单',
   'user', @CurrentUser, 'table', 'DeliveryOutReceipt'
go

/*==============================================================*/
/* Table: Operator                                              */
/*==============================================================*/
create table Operator (
   OperatorID           int                  not null,
   OperatorName         varchar(20)          null,
   RoleID               int                  null,
   constraint PK_OPERATOR primary key nonclustered (OperatorID)
)
go

/*==============================================================*/
/* Table: "Order"                                               */
/*==============================================================*/
create table "Order" (
   OrderID              int                  not null,
   OrderCode            varchar(50)          null,
   PickupReceiptID      int                  null,
   CustomerID           int                  null,
   OrderStatus          varchar(10)          null,
   constraint PK_ORDER primary key nonclustered (OrderID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '客户订单表',
   'user', @CurrentUser, 'table', 'Order'
go

/*==============================================================*/
/* Table: OrderLog                                              */
/*==============================================================*/
create table OrderLog (
   OrderLogID           int                  not null,
   OrderCode            varchar(50)          null,
   LogTime              datetime             null,
   constraint PK_ORDERLOG primary key nonclustered (OrderLogID)
)
go

/*==============================================================*/
/* Table: OrderToPickup                                         */
/*==============================================================*/
create table OrderToPickup (
   ID                   int                  not null,
   OrderID              int                  null,
   PickupReceiptID      int                  null,
   constraint PK_ORDERTOPICKUP primary key nonclustered (ID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '订单和取件单对应关系表',
   'user', @CurrentUser, 'table', 'OrderToPickup'
go

/*==============================================================*/
/* Table: PickupReceipt                                         */
/*==============================================================*/
create table PickupReceipt (
   PickupReceiptID      int                  not null,
   SupplierID           int                  null,
   constraint PK_PICKUPRECEIPT primary key nonclustered (PickupReceiptID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '取件单',
   'user', @CurrentUser, 'table', 'PickupReceipt'
go

/*==============================================================*/
/* Table: Role                                                  */
/*==============================================================*/
create table Role (
   RoleID               int                  not null,
   RoleName             varchar(40)          null,
   constraint PK_ROLE primary key nonclustered (RoleID)
)
go

/*==============================================================*/
/* Table: SendOutReceipt                                        */
/*==============================================================*/
create table SendOutReceipt (
   SendOutReceiptID     int                  not null,
   SupplierID           int                  null,
   constraint PK_SENDOUTRECEIPT primary key nonclustered (SendOutReceiptID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '送件单',
   'user', @CurrentUser, 'table', 'SendOutReceipt'
go

/*==============================================================*/
/* Table: Supplier                                              */
/*==============================================================*/
create table Supplier (
   SupplierID           int                  not null,
   SupplierName         varchar(100)         null,
   ContactName          varchar(40)          null,
   ContactTel           varchar(40)          null,
   constraint PK_SUPPLIER primary key nonclustered (SupplierID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   'di san fang gongyingshang xinxi biao ',
   'user', @CurrentUser, 'table', 'Supplier'
go

/*==============================================================*/
/* Table: Warehouse                                             */
/*==============================================================*/
create table Warehouse (
   WarehouseID          int                  not null,
   constraint PK_WAREHOUSE primary key nonclustered (WarehouseID)
)
go

declare @CurrentUser sysname
select @CurrentUser = user_name()
execute sp_addextendedproperty 'MS_Description', 
   '物流仓库表',
   'user', @CurrentUser, 'table', 'Warehouse'
go

