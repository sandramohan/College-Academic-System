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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[coursecrs]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[coursecrs](
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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[student]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[student](
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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[class teacher]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[class teacher](
	[ctid] [int] NOT NULL,
	[courseid] [int] NOT NULL,
	[sem] [int] NOT NULL,
	[staffid] [int] NOT NULL
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
	[courseid] [int] NOT NULL,
	[name] [varchar](50) NOT NULL,
	[duration] [int] NOT NULL,
 CONSTRAINT [PK__course__00551192] PRIMARY KEY CLUSTERED 
(
	[courseid] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
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
	[StaffId] [int] NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Gender] [varchar](50) NOT NULL,
	[DOB] [varchar](20) NOT NULL,
	[Email] [varchar](50) NOT NULL,
	[Phno] [numeric](18, 0) NOT NULL,
	[Experience] [numeric](18, 0) NOT NULL,
	[DOJ] [varchar](20) NOT NULL,
	[Designation] [varchar](50) NOT NULL,
	[Qualification] [varchar](50) NOT NULL,
	[status] [int] NULL,
 CONSTRAINT [PK_staff_1] PRIMARY KEY CLUSTERED 
(
	[StaffId] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
