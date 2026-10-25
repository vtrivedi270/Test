/* Stationery Management System - SQL Server */
CREATE DATABASE StationeryDB;
GO
USE StationeryDB;
GO

/* ---------- MASTERS ---------- */
CREATE TABLE DivisionMaster (
    DivisionId     INT IDENTITY PRIMARY KEY,
    DivisionName   NVARCHAR(150) NOT NULL UNIQUE,
    DivisionHeadId INT NULL,                       -- FK added after EmployeeMaster
    IsActive       BIT NOT NULL DEFAULT 1,
    CreatedBy NVARCHAR(256), CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    ModifiedBy NVARCHAR(256), ModifiedAt DATETIME2 NULL
);

CREATE TABLE DepartmentMaster (
    DepartmentId     INT IDENTITY PRIMARY KEY,
    DepartmentName   NVARCHAR(150) NOT NULL,
    DivisionId       INT NOT NULL REFERENCES DivisionMaster(DivisionId),
    DepartmentHeadId INT NULL,
    IsActive         BIT NOT NULL DEFAULT 1,
    CreatedBy NVARCHAR(256), CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    ModifiedBy NVARCHAR(256), ModifiedAt DATETIME2 NULL,
    CONSTRAINT UQ_Dept UNIQUE (DepartmentName, DivisionId)
);

CREATE TABLE EmployeeMaster (
    EmployeeId         INT IDENTITY PRIMARY KEY,
    EmployeeCode       NVARCHAR(30) NOT NULL UNIQUE,
    FullName           NVARCHAR(200) NOT NULL,
    Email              NVARCHAR(256) NOT NULL UNIQUE,   -- matches Entra ID UPN
    DepartmentId       INT NOT NULL REFERENCES DepartmentMaster(DepartmentId),
    DivisionId         INT NOT NULL REFERENCES DivisionMaster(DivisionId),
    ReportingManagerId INT NULL REFERENCES EmployeeMaster(EmployeeId),
    Role               NVARCHAR(30) NOT NULL DEFAULT 'Requestor'
        CHECK (Role IN ('Requestor','ReportingManager','DeptHead','DivHead','HRAD','Admin')),
    IsActive           BIT NOT NULL DEFAULT 1,
    CreatedBy NVARCHAR(256), CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    ModifiedBy NVARCHAR(256), ModifiedAt DATETIME2 NULL
);
ALTER TABLE DivisionMaster   ADD CONSTRAINT FK_Div_Head  FOREIGN KEY (DivisionHeadId)   REFERENCES EmployeeMaster(EmployeeId);
ALTER TABLE DepartmentMaster ADD CONSTRAINT FK_Dept_Head FOREIGN KEY (DepartmentHeadId) REFERENCES EmployeeMaster(EmployeeId);

CREATE TABLE ItemCategoryMaster (
    CategoryId INT IDENTITY PRIMARY KEY,
    CategoryName NVARCHAR(100) NOT NULL UNIQUE,
    IsActive BIT NOT NULL DEFAULT 1
);

CREATE TABLE ItemMaster (
    ItemId        INT IDENTITY PRIMARY KEY,
    ItemCode      NVARCHAR(30)  NOT NULL UNIQUE,
    ItemName      NVARCHAR(200) NOT NULL,
    CategoryId    INT NULL REFERENCES ItemCategoryMaster(CategoryId),
    Unit          NVARCHAR(30)  NOT NULL,              -- Nos, Box, Ream
    Rate          DECIMAL(18,2) NOT NULL DEFAULT 0,
    StockQty      DECIMAL(18,2) NOT NULL DEFAULT 0,    -- current store availability
    ReorderLevel  DECIMAL(18,2) NOT NULL DEFAULT 0,
    IsActive      BIT NOT NULL DEFAULT 1,
    CreatedBy NVARCHAR(256), CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    ModifiedBy NVARCHAR(256), ModifiedAt DATETIME2 NULL
);

CREATE TABLE VendorMaster (
    VendorId      INT IDENTITY PRIMARY KEY,
    VendorName    NVARCHAR(200) NOT NULL,
    ContactPerson NVARCHAR(150),
    Email         NVARCHAR(256) NOT NULL,
    Phone         NVARCHAR(30),
    Address       NVARCHAR(500),
    GSTNo         NVARCHAR(30),
    IsActive      BIT NOT NULL DEFAULT 1,
    CreatedBy NVARCHAR(256), CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    ModifiedBy NVARCHAR(256), ModifiedAt DATETIME2 NULL
);

/* ---------- TRANSACTIONS ---------- */
CREATE TABLE RequestHeader (
    RequestId       INT IDENTITY PRIMARY KEY,
    RequestNo       NVARCHAR(30) NOT NULL UNIQUE,
    RequestDate     DATE NOT NULL DEFAULT CAST(SYSUTCDATETIME() AS DATE),
    RequestorId     INT NOT NULL REFERENCES EmployeeMaster(EmployeeId),
    DepartmentId    INT NOT NULL REFERENCES DepartmentMaster(DepartmentId),
    DivisionId      INT NOT NULL REFERENCES DivisionMaster(DivisionId),
    Remarks         NVARCHAR(1000),
    TotalAmount     DECIMAL(18,2) NOT NULL DEFAULT 0,
    Status          NVARCHAR(30) NOT NULL DEFAULT 'Draft'
        CHECK (Status IN ('Draft','InReview','AmendRequired','Rejected','HRApproval',
                          'Ordered','Delivered','Closed','PartialClosed')),
    CurrentStage    NVARCHAR(30) NULL
        CHECK (CurrentStage IN ('ReportingManager','DeptHead','DivHead','HRAD','Requestor')),
    CurrentApproverId INT NULL REFERENCES EmployeeMaster(EmployeeId),
    VendorId        INT NULL REFERENCES VendorMaster(VendorId),
    PONumber        NVARCHAR(30) NULL,
    PODate          DATE NULL,
    SubmittedAt     DATETIME2 NULL,
    ClosedAt        DATETIME2 NULL,
    RowVer          ROWVERSION,
    CreatedBy NVARCHAR(256), CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    ModifiedBy NVARCHAR(256), ModifiedAt DATETIME2 NULL
);
CREATE INDEX IX_Req_Requestor ON RequestHeader(RequestorId, Status);
CREATE INDEX IX_Req_Approver  ON RequestHeader(CurrentApproverId, Status);
CREATE INDEX IX_Req_Date      ON RequestHeader(RequestDate) INCLUDE (DepartmentId, TotalAmount, Status);

CREATE TABLE RequestItem (
    RequestItemId INT IDENTITY PRIMARY KEY,
    RequestId     INT NOT NULL REFERENCES RequestHeader(RequestId) ON DELETE CASCADE,
    ItemId        INT NOT NULL REFERENCES ItemMaster(ItemId),
    Unit          NVARCHAR(30) NOT NULL,
    Rate          DECIMAL(18,2) NOT NULL,
    Quantity      DECIMAL(18,2) NOT NULL CHECK (Quantity > 0),
    ApprovedQty   DECIMAL(18,2) NULL,                 -- HRAD may reduce
    IssuedFromStockQty DECIMAL(18,2) NOT NULL DEFAULT 0,
    ReceivedQty   DECIMAL(18,2) NOT NULL DEFAULT 0,   -- verified by requestor
    LineTotal     AS (Quantity * Rate) PERSISTED,
    Remarks       NVARCHAR(500),
    CONSTRAINT UQ_ReqItem UNIQUE (RequestId, ItemId)
);

CREATE TABLE ApprovalHistory (                         -- the "history / timeline" tab
    HistoryId   BIGINT IDENTITY PRIMARY KEY,
    RequestId   INT NOT NULL REFERENCES RequestHeader(RequestId) ON DELETE CASCADE,
    Stage       NVARCHAR(30) NOT NULL,
    Action      NVARCHAR(30) NOT NULL
        CHECK (Action IN ('Submit','Approve','Reject','Amend','Resubmit','VendorSent',
                          'Delivered','Handover','Close','PartialClose','Query')),
    ActorId     INT NOT NULL REFERENCES EmployeeMaster(EmployeeId),
    Comments    NVARCHAR(1000),
    ActionDate  DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
);
CREATE INDEX IX_Hist_Req ON ApprovalHistory(RequestId, ActionDate);

CREATE TABLE StockTransaction (
    StockTxnId BIGINT IDENTITY PRIMARY KEY,
    ItemId     INT NOT NULL REFERENCES ItemMaster(ItemId),
    TxnType    NVARCHAR(20) NOT NULL CHECK (TxnType IN ('Receipt','Issue','Adjustment')),
    Quantity   DECIMAL(18,2) NOT NULL,
    RequestId  INT NULL REFERENCES RequestHeader(RequestId),
    TxnDate    DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
    CreatedBy  NVARCHAR(256)
);

/* ---------- LOGS ---------- */
CREATE TABLE AuditLog (
    AuditId    BIGINT IDENTITY PRIMARY KEY,
    EntityName NVARCHAR(100) NOT NULL,
    EntityId   NVARCHAR(50)  NOT NULL,
    Action     NVARCHAR(30)  NOT NULL,
    OldValue   NVARCHAR(MAX), NewValue NVARCHAR(MAX),
    UserEmail  NVARCHAR(256), LogDate DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
);
CREATE TABLE ExceptionLog (
    ExceptionId BIGINT IDENTITY PRIMARY KEY,
    Source NVARCHAR(200), Message NVARCHAR(MAX), StackTrace NVARCHAR(MAX),
    RequestPath NVARCHAR(500), UserEmail NVARCHAR(256),
    LogDate DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
);
CREATE TABLE EmailLog (
    EmailLogId BIGINT IDENTITY PRIMARY KEY,
    RequestId  INT NULL REFERENCES RequestHeader(RequestId),
    ToEmail NVARCHAR(500) NOT NULL, CcEmail NVARCHAR(500),
    Subject NVARCHAR(500) NOT NULL, Body NVARCHAR(MAX),
    Status NVARCHAR(20) NOT NULL DEFAULT 'Pending' CHECK (Status IN ('Pending','Sent','Failed')),
    ErrorMessage NVARCHAR(MAX), RetryCount INT NOT NULL DEFAULT 0,
    CreatedAt DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(), SentAt DATETIME2 NULL
);
GO

/* ---------- SEQUENCE + REQUEST NO ---------- */
CREATE SEQUENCE dbo.SeqRequestNo START WITH 1 INCREMENT BY 1;
GO
CREATE PROCEDURE dbo.usp_NextRequestNo @RequestNo NVARCHAR(30) OUTPUT AS
BEGIN
  SET @RequestNo = 'STN-' + CAST(YEAR(GETDATE()) AS NVARCHAR(4)) + '-' +
                   RIGHT('00000' + CAST(NEXT VALUE FOR dbo.SeqRequestNo AS NVARCHAR(10)),5);
END
GO

/* ---------- DASHBOARD PROCS ---------- */
CREATE PROCEDURE dbo.usp_Dash_Monthly @From DATE, @To DATE, @DeptId INT = NULL AS
SELECT FORMAT(h.RequestDate,'yyyy-MM') AS [Month], COUNT(*) AS Requests, SUM(h.TotalAmount) AS TotalCost
FROM RequestHeader h
WHERE h.RequestDate BETWEEN @From AND @To AND h.Status NOT IN ('Draft','Rejected')
  AND (@DeptId IS NULL OR h.DepartmentId=@DeptId)
GROUP BY FORMAT(h.RequestDate,'yyyy-MM') ORDER BY 1;
GO
CREATE PROCEDURE dbo.usp_Dash_ByDepartment @From DATE, @To DATE AS
SELECT d.DepartmentName, COUNT(*) AS Requests, SUM(h.TotalAmount) AS TotalCost
FROM RequestHeader h JOIN DepartmentMaster d ON d.DepartmentId=h.DepartmentId
WHERE h.RequestDate BETWEEN @From AND @To AND h.Status NOT IN ('Draft','Rejected')
GROUP BY d.DepartmentName ORDER BY TotalCost DESC;
GO
CREATE PROCEDURE dbo.usp_Dash_ByItem @From DATE, @To DATE, @Top INT = 10 AS
SELECT TOP (@Top) i.ItemName, SUM(ri.Quantity) AS Qty, SUM(ri.LineTotal) AS TotalCost
FROM RequestItem ri JOIN RequestHeader h ON h.RequestId=ri.RequestId
JOIN ItemMaster i ON i.ItemId=ri.ItemId
WHERE h.RequestDate BETWEEN @From AND @To AND h.Status NOT IN ('Draft','Rejected')
GROUP BY i.ItemName ORDER BY TotalCost DESC;
GO
CREATE VIEW dbo.vw_StoreAvailability AS
SELECT ItemId, ItemCode, ItemName, Unit, StockQty, ReorderLevel,
       CASE WHEN StockQty <= ReorderLevel THEN 1 ELSE 0 END AS IsLow
FROM ItemMaster WHERE IsActive=1;
GO

/* ---------- SEED ---------- */
INSERT INTO DivisionMaster(DivisionName) VALUES ('Operations'),('Finance');
INSERT INTO DepartmentMaster(DepartmentName,DivisionId) VALUES ('IT',1),('HR & Admin',1),('Accounts',2);
INSERT INTO ItemCategoryMaster(CategoryName) VALUES ('Paper'),('Writing'),('Files & Folders');
INSERT INTO ItemMaster(ItemCode,ItemName,CategoryId,Unit,Rate,StockQty,ReorderLevel)
VALUES ('STN001','A4 Paper Ream',1,'Ream',280,50,10),('STN002','Ball Pen Blue',2,'Nos',10,200,50);
