SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[subject]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[subject](
	[subid] [int] NOT NULL,
	[subname] [varchar](50) NOT NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[staff]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[staff](
	[Name] [varchar](50) NOT NULL,
	[Gender] [varchar](50) NOT NULL,
	[DOB] [datetime] NOT NULL,
	[Email] [varchar](50) NOT NULL,
	[Phno] [numeric](18, 0) NOT NULL,
	[Experience] [numeric](18, 0) NOT NULL,
	[DOJ] [datetime] NOT NULL,
	[Designation] [varchar](50) NOT NULL,
	[Qualification] [varchar](50) NOT NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[course crs]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[course crs](
	[Courseid] [numeric](18, 0) NOT NULL,
	[Subjectid] [numeric](18, 0) NOT NULL,
	[semester] [numeric](18, 0) NOT NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[student details]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[student details](
	[studid] [numeric](18, 0) NOT NULL,
	[name] [varchar](50) NOT NULL,
	[gender] [varchar](50) NOT NULL,
	[DOB] [datetime] NOT NULL,
	[Email] [varchar](50) NOT NULL,
	[Department] [varchar](50) NOT NULL,
	[Phone] [numeric](18, 0) NOT NULL,
	[Guadian] [varchar](50) NOT NULL,
	[Address] [varchar](50) NOT NULL,
	[Courseduration] [numeric](18, 0) NOT NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[course]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[course](
	[courseid] [decimal](18, 0) NOT NULL,
	[name] [varchar](50) NULL,
	[duration] [numeric](18, 0) NULL,
PRIMARY KEY CLUSTERED 
(
	[courseid] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
