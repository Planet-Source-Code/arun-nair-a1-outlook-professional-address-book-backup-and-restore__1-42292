IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'bqinAB')
	DROP DATABASE [bqinAB]
GO

CREATE DATABASE [bqinAB]  ON (NAME = N'bqinAB_Data', FILENAME = N'D:\MSSQL7\Data\bqinAB_Data.MDF' , SIZE = 4, FILEGROWTH = 10%) LOG ON (NAME = N'bqinAB_Log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\Data\bqinAB_Log.LDF' , SIZE = 1, MAXSIZE = 100, FILEGROWTH = 10%)
 COLLATE SQL_Latin1_General_CP1_CI_AS
GO

exec sp_dboption N'bqinAB', N'autoclose', N'false'
GO

exec sp_dboption N'bqinAB', N'bulkcopy', N'false'
GO

exec sp_dboption N'bqinAB', N'trunc. log', N'false'
GO

exec sp_dboption N'bqinAB', N'torn page detection', N'true'
GO

exec sp_dboption N'bqinAB', N'read only', N'false'
GO

exec sp_dboption N'bqinAB', N'dbo use', N'false'
GO

exec sp_dboption N'bqinAB', N'single', N'false'
GO

exec sp_dboption N'bqinAB', N'autoshrink', N'false'
GO

exec sp_dboption N'bqinAB', N'ANSI null default', N'false'
GO

exec sp_dboption N'bqinAB', N'recursive triggers', N'false'
GO

exec sp_dboption N'bqinAB', N'ANSI nulls', N'false'
GO

exec sp_dboption N'bqinAB', N'concat null yields null', N'false'
GO

exec sp_dboption N'bqinAB', N'cursor close on commit', N'false'
GO

exec sp_dboption N'bqinAB', N'default to local cursor', N'false'
GO

exec sp_dboption N'bqinAB', N'quoted identifier', N'false'
GO

exec sp_dboption N'bqinAB', N'ANSI warnings', N'false'
GO

exec sp_dboption N'bqinAB', N'auto create statistics', N'true'
GO

exec sp_dboption N'bqinAB', N'auto update statistics', N'true'
GO

use [bqinAB]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AddProfile]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AddProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProfileView]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ProfileView]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[abView]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[abView]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AddressBookMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AddressBookMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProfileMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ProfileMaster]
GO

CREATE TABLE [dbo].[AddressBookMaster] (
	[WABId] [int] IDENTITY (1, 1) NOT NULL ,
	[ProfileId] [smallint] NOT NULL ,
	[Full Name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[E-mail Address] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[E-mail Display Name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Business2TelephoneNumber] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddress] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddressCity] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddressCountry] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddressPostalCode] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddressPostOfficeBox] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddressState] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessAddressStreet] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessFaxNumber] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessHomePage] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BusinessTelephoneNumber] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email1DisplayName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email2Address] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email2DisplayName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email3Address] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email3DisplayName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MobileTelephoneNumber] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Home2TelephoneNumber] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HomeTelephoneNumber] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[JobTitle] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ProfileMaster] (
	[ProfileId] [int] NOT NULL ,
	[User name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comment] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Date of the Profile] [datetime] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ProfileMaster] WITH NOCHECK ADD 
	CONSTRAINT [PK_ProfileMaster] PRIMARY KEY  CLUSTERED 
	(
		[ProfileId]
	)  ON [PRIMARY] 
GO

 CREATE  CLUSTERED  INDEX [idx_ProfileId] ON [dbo].[AddressBookMaster]([ProfileId]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[AddressBookMaster] WITH NOCHECK ADD 
	CONSTRAINT [PK_AddressBookMaster] PRIMARY KEY  NONCLUSTERED 
	(
		[WABId]
	)  ON [PRIMARY] 
GO

 CREATE  INDEX [idxUserName] ON [dbo].[ProfileMaster]([User name]) ON [PRIMARY]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE Procedure AddProfile  
@username nvarchar(50),
@comment nvarchar(250),
@dateOfProfile datetime,
@profileid integer output
AS

  declare @profId integer
  select @profid=max(profileid)+1 from ProfileMaster
 
  if @profid is null
  begin
	select @profid=1
  end

 

  insert into profilemaster
  values(@profid,@username,@comment,@dateOfProfile)

  select @profileid=@profid
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create Procedure ProfileView 
@username nvarchar(50)
as
if len(@username)<1
begin
   select distinct [User name] from profilemaster	
end
else
begin
   select ProfileId,Comment,[Date of the Profile]  from profilemaster
   where [User name]=@username	
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure abView
@ProfileId integer
as
select count(*) as "Ctr"   from AddressBookMaster
where profileid=@ProfileId
select [Full name],[E-Mail Address],
Business2TelephoneNumber,
BusinessAddress,
BusinessAddressCity,
BusinessAddressCountry,
BusinessAddressPostalCode,
BusinessAddressPostOfficeBox,
BusinessAddressState,
BusinessAddressStreet,
BusinessFaxNumber,
BusinessHomePage,
BusinessTelephoneNumber,
Email1DisplayName,
Email2Address,
Email2DisplayName,
Email3Address,
Email3DisplayName,
MobileTelephoneNumber,
Home2TelephoneNumber,
HomeTelephoneNumber,
JobTitle 
 from AddressBookMaster
where profileid=@ProfileId
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

