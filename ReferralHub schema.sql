/* =====================================================================
   Referral Hub - SQL Server schema (SQL Server 2017+ for STRING_AGG)
   Flow: Requisition -> 3-step approval -> Open -> Referral -> HR review
         -> Interview 1..3 (feedback each) -> Final approval (Div + HR head)
   Business-rule transitions live in the .NET service layer (one DB
   transaction per action); constraints here protect data integrity.
   ===================================================================== */

IF DB_ID('ReferralHub') IS NULL CREATE DATABASE ReferralHub;
GO
USE ReferralHub;
GO
IF SCHEMA_ID('rp') IS NULL EXEC('CREATE SCHEMA rp');
GO

/* ---------- People and organisation ---------- */
CREATE TABLE rp.Employee(
    EmployeeId          int IDENTITY(1,1) CONSTRAINT PK_Employee PRIMARY KEY,
    EntraObjectId       uniqueidentifier NULL,               -- Microsoft Entra ID user id (from the SPFx token)
    Email               nvarchar(256) NOT NULL,              -- UPN, used to map the signed-in user
    DisplayName         nvarchar(150) NOT NULL,
    JobTitle            nvarchar(150) NULL,
    DepartmentId        int NULL,                            -- FK added below (circular with Department)
    ManagerEmployeeId   int NULL CONSTRAINT FK_Employee_Manager REFERENCES rp.Employee(EmployeeId),
    IsActive            bit NOT NULL CONSTRAINT DF_Employee_Active DEFAULT 1,
    CreatedOn           datetime2(0) NOT NULL CONSTRAINT DF_Employee_Created DEFAULT SYSUTCDATETIME(),
    CONSTRAINT UQ_Employee_Email UNIQUE (Email)
);
CREATE UNIQUE INDEX UX_Employee_EntraObjectId ON rp.Employee(EntraObjectId) WHERE EntraObjectId IS NOT NULL;

CREATE TABLE rp.Division(
    DivisionId              int IDENTITY(1,1) CONSTRAINT PK_Division PRIMARY KEY,
    Name                    nvarchar(120) NOT NULL CONSTRAINT UQ_Division_Name UNIQUE,
    DivisionHeadEmployeeId  int NOT NULL CONSTRAINT FK_Division_Head REFERENCES rp.Employee(EmployeeId)
);

CREATE TABLE rp.Department(
    DepartmentId                int IDENTITY(1,1) CONSTRAINT PK_Department PRIMARY KEY,
    DivisionId                  int NOT NULL CONSTRAINT FK_Department_Division REFERENCES rp.Division(DivisionId),
    Name                        nvarchar(120) NOT NULL CONSTRAINT UQ_Department_Name UNIQUE,
    DepartmentHeadEmployeeId    int NOT NULL CONSTRAINT FK_Department_Head REFERENCES rp.Employee(EmployeeId)
);

ALTER TABLE rp.Employee ADD CONSTRAINT FK_Employee_Department
    FOREIGN KEY (DepartmentId) REFERENCES rp.Department(DepartmentId);

CREATE TABLE rp.AppRole(
    RoleId  tinyint NOT NULL CONSTRAINT PK_AppRole PRIMARY KEY,
    Code    varchar(20) NOT NULL CONSTRAINT UQ_AppRole_Code UNIQUE,
    Name    nvarchar(60) NOT NULL
);

CREATE TABLE rp.EmployeeRole(
    EmployeeId  int NOT NULL CONSTRAINT FK_EmployeeRole_Employee REFERENCES rp.Employee(EmployeeId),
    RoleId      tinyint NOT NULL CONSTRAINT FK_EmployeeRole_Role REFERENCES rp.AppRole(RoleId),
    CONSTRAINT PK_EmployeeRole PRIMARY KEY (EmployeeId, RoleId)
);

/* ---------- Lookups ---------- */
CREATE TABLE rp.Location(
    LocationId  int IDENTITY(1,1) CONSTRAINT PK_Location PRIMARY KEY,
    Name        nvarchar(100) NOT NULL CONSTRAINT UQ_Location_Name UNIQUE
);

CREATE TABLE rp.Skill(
    SkillId int IDENTITY(1,1) CONSTRAINT PK_Skill PRIMARY KEY,
    Name    nvarchar(80) NOT NULL CONSTRAINT UQ_Skill_Name UNIQUE
);

CREATE TABLE rp.ReferralStage(
    StageId     tinyint NOT NULL CONSTRAINT PK_ReferralStage PRIMARY KEY,
    Name        nvarchar(40) NOT NULL
);

/* ---------- Job requisition and its approval chain ---------- */
CREATE TABLE rp.JobRequisition(
    RequisitionId       int IDENTITY(1,1) CONSTRAINT PK_JobRequisition PRIMARY KEY,
    Title               nvarchar(200) NOT NULL,
    DepartmentId        int NOT NULL CONSTRAINT FK_Req_Department REFERENCES rp.Department(DepartmentId),
    LocationId          int NOT NULL CONSTRAINT FK_Req_Location REFERENCES rp.Location(LocationId),
    EmploymentType      varchar(12) NOT NULL CONSTRAINT DF_Req_Type DEFAULT 'FullTime'
                        CONSTRAINT CK_Req_Type CHECK (EmploymentType IN ('FullTime','PartTime','Contract')),
    NoOfPositions       smallint NOT NULL CONSTRAINT CK_Req_Positions CHECK (NoOfPositions > 0),
    PositionsFilled     smallint NOT NULL CONSTRAINT DF_Req_Filled DEFAULT 0,
    ExperienceRequired  nvarchar(50) NULL,
    HireByDate          date NOT NULL,
    Summary             nvarchar(max) NOT NULL,
    Requirements        nvarchar(max) NOT NULL,
    RequestedByEmployeeId int NOT NULL CONSTRAINT FK_Req_RequestedBy REFERENCES rp.Employee(EmployeeId),
    Status              varchar(16) NOT NULL CONSTRAINT DF_Req_Status DEFAULT 'PendingApproval'
                        CONSTRAINT CK_Req_Status CHECK (Status IN ('Draft','PendingApproval','Open','Rejected','Closed','Filled')),
    RejectionReason     nvarchar(500) NULL,
    PublishedOn         datetime2(0) NULL,                   -- set when the HR head approves
    CreatedOn           datetime2(0) NOT NULL CONSTRAINT DF_Req_Created DEFAULT SYSUTCDATETIME(),
    UpdatedOn           datetime2(0) NOT NULL CONSTRAINT DF_Req_Updated DEFAULT SYSUTCDATETIME(),
    RowVer              rowversion,
    CONSTRAINT CK_Req_Filled CHECK (PositionsFilled <= NoOfPositions)
);
CREATE INDEX IX_Req_Status_HireBy ON rp.JobRequisition(Status, HireByDate) INCLUDE (Title, DepartmentId, LocationId);
CREATE INDEX IX_Req_RequestedBy ON rp.JobRequisition(RequestedByEmployeeId);

CREATE TABLE rp.JobRequisitionSkill(
    RequisitionId   int NOT NULL CONSTRAINT FK_ReqSkill_Req REFERENCES rp.JobRequisition(RequisitionId) ON DELETE CASCADE,
    SkillId         int NOT NULL CONSTRAINT FK_ReqSkill_Skill REFERENCES rp.Skill(SkillId),
    CONSTRAINT PK_JobRequisitionSkill PRIMARY KEY (RequisitionId, SkillId)
);

/* One row per approval step, created when the requisition is submitted.
   Step 1 = Department head (Pending), steps 2 and 3 start as Waiting. */
CREATE TABLE rp.RequisitionApproval(
    ApprovalId          int IDENTITY(1,1) CONSTRAINT PK_RequisitionApproval PRIMARY KEY,
    RequisitionId       int NOT NULL CONSTRAINT FK_ReqApproval_Req REFERENCES rp.JobRequisition(RequisitionId) ON DELETE CASCADE,
    StepNo              tinyint NOT NULL CONSTRAINT CK_ReqApproval_Step CHECK (StepNo BETWEEN 1 AND 3),
    ApproverRole        varchar(10) NOT NULL CONSTRAINT CK_ReqApproval_Role CHECK (ApproverRole IN ('DEPT_HEAD','DIV_HEAD','HR_HEAD')),
    ApproverEmployeeId  int NOT NULL CONSTRAINT FK_ReqApproval_Approver REFERENCES rp.Employee(EmployeeId),
    Status              varchar(10) NOT NULL CONSTRAINT DF_ReqApproval_Status DEFAULT 'Waiting'
                        CONSTRAINT CK_ReqApproval_Status CHECK (Status IN ('Waiting','Pending','Approved','Rejected')),
    ActedOn             datetime2(0) NULL,
    Comments            nvarchar(500) NULL,
    CONSTRAINT UQ_ReqApproval_Step UNIQUE (RequisitionId, StepNo),
    CONSTRAINT UQ_ReqApproval_Role UNIQUE (RequisitionId, ApproverRole),
    CONSTRAINT CK_ReqApproval_Acted CHECK ((Status IN ('Approved','Rejected') AND ActedOn IS NOT NULL) OR Status IN ('Waiting','Pending'))
);
CREATE INDEX IX_ReqApproval_Inbox ON rp.RequisitionApproval(ApproverEmployeeId) WHERE Status = 'Pending';

/* ---------- Referrals ---------- */
CREATE TABLE rp.Referral(
    ReferralId              int IDENTITY(1,1) CONSTRAINT PK_Referral PRIMARY KEY,
    RequisitionId           int NOT NULL CONSTRAINT FK_Referral_Req REFERENCES rp.JobRequisition(RequisitionId),
    ReferredByEmployeeId    int NOT NULL CONSTRAINT FK_Referral_Referrer REFERENCES rp.Employee(EmployeeId),
    CandidateName           nvarchar(150) NOT NULL,
    CandidateEmail          nvarchar(256) NOT NULL,
    CandidatePhone          nvarchar(30) NOT NULL,
    TotalExperienceYears    decimal(4,1) NOT NULL CONSTRAINT CK_Referral_Exp CHECK (TotalExperienceYears >= 0),
    CurrentCompany          nvarchar(150) NULL,
    NoticePeriod            varchar(20) NULL,
    Relationship            nvarchar(60) NULL,               -- how the employee knows the candidate
    ReferrerNote            nvarchar(1000) NULL,             -- why they are a good fit
    CurrentStageId          tinyint NOT NULL CONSTRAINT DF_Referral_Stage DEFAULT 1
                            CONSTRAINT FK_Referral_Stage REFERENCES rp.ReferralStage(StageId),
    Outcome                 varchar(10) NOT NULL CONSTRAINT DF_Referral_Outcome DEFAULT 'Active'
                            CONSTRAINT CK_Referral_Outcome CHECK (Outcome IN ('Active','Selected','Rejected','Withdrawn')),
    ClosedReason            nvarchar(500) NULL,
    AssignedHrEmployeeId    int NULL CONSTRAINT FK_Referral_Hr REFERENCES rp.Employee(EmployeeId),
    ReferredOn              datetime2(0) NOT NULL CONSTRAINT DF_Referral_On DEFAULT SYSUTCDATETIME(),
    UpdatedOn               datetime2(0) NOT NULL CONSTRAINT DF_Referral_Updated DEFAULT SYSUTCDATETIME(),
    RowVer                  rowversion,
    CONSTRAINT UQ_Referral_Candidate UNIQUE (RequisitionId, CandidateEmail)   -- same person cannot be referred twice for one job
);
CREATE INDEX IX_Referral_Referrer ON rp.Referral(ReferredByEmployeeId, ReferredOn DESC);
CREATE INDEX IX_Referral_Pipeline ON rp.Referral(Outcome, CurrentStageId) INCLUDE (RequisitionId, CandidateName);

/* Resume file lives in a SharePoint document library; the DB keeps the link. */
CREATE TABLE rp.ReferralDocument(
    DocumentId      int IDENTITY(1,1) CONSTRAINT PK_ReferralDocument PRIMARY KEY,
    ReferralId      int NOT NULL CONSTRAINT FK_RefDoc_Referral REFERENCES rp.Referral(ReferralId) ON DELETE CASCADE,
    DocType         varchar(20) NOT NULL CONSTRAINT DF_RefDoc_Type DEFAULT 'Resume',
    FileName        nvarchar(260) NOT NULL,
    ContentType     varchar(100) NOT NULL,
    SizeBytes       bigint NOT NULL CONSTRAINT CK_RefDoc_Size CHECK (SizeBytes BETWEEN 1 AND 5242880),
    StorageUrl      nvarchar(1000) NOT NULL,
    UploadedOn      datetime2(0) NOT NULL CONSTRAINT DF_RefDoc_On DEFAULT SYSUTCDATETIME()
);
CREATE INDEX IX_RefDoc_Referral ON rp.ReferralDocument(ReferralId);

/* Timeline: contacted, shortlisted, rejected, stage moves, notes */
CREATE TABLE rp.ReferralActivity(
    ActivityId      bigint IDENTITY(1,1) CONSTRAINT PK_ReferralActivity PRIMARY KEY,
    ReferralId      int NOT NULL CONSTRAINT FK_RefAct_Referral REFERENCES rp.Referral(ReferralId) ON DELETE CASCADE,
    ActivityType    varchar(30) NOT NULL,   -- Submitted, Contacted, Shortlisted, InterviewScheduled, FeedbackSubmitted, FinalApproved, Rejected...
    ActorEmployeeId int NOT NULL CONSTRAINT FK_RefAct_Actor REFERENCES rp.Employee(EmployeeId),
    Notes           nvarchar(1000) NULL,
    CreatedOn       datetime2(0) NOT NULL CONSTRAINT DF_RefAct_On DEFAULT SYSUTCDATETIME()
);
CREATE INDEX IX_RefAct_Referral ON rp.ReferralActivity(ReferralId, CreatedOn);

/* ---------- Interviews (3 rounds) and feedback ---------- */
CREATE TABLE rp.Interview(
    InterviewId             int IDENTITY(1,1) CONSTRAINT PK_Interview PRIMARY KEY,
    ReferralId              int NOT NULL CONSTRAINT FK_Interview_Referral REFERENCES rp.Referral(ReferralId) ON DELETE CASCADE,
    RoundNo                 tinyint NOT NULL CONSTRAINT CK_Interview_Round CHECK (RoundNo BETWEEN 1 AND 3),
    InterviewerEmployeeId   int NOT NULL CONSTRAINT FK_Interview_Interviewer REFERENCES rp.Employee(EmployeeId),
    ScheduledAt             datetime2(0) NOT NULL,
    Mode                    varchar(8) NOT NULL CONSTRAINT DF_Interview_Mode DEFAULT 'Online'
                            CONSTRAINT CK_Interview_Mode CHECK (Mode IN ('Online','OnSite')),
    MeetingLink             nvarchar(500) NULL,
    Status                  varchar(10) NOT NULL CONSTRAINT DF_Interview_Status DEFAULT 'Scheduled'
                            CONSTRAINT CK_Interview_Status CHECK (Status IN ('Scheduled','Completed','Cancelled')),
    ScheduledByEmployeeId   int NOT NULL CONSTRAINT FK_Interview_ScheduledBy REFERENCES rp.Employee(EmployeeId),
    CreatedOn               datetime2(0) NOT NULL CONSTRAINT DF_Interview_Created DEFAULT SYSUTCDATETIME(),
    CONSTRAINT UQ_Interview_Round UNIQUE (ReferralId, RoundNo)     -- reschedule = update the row
);
CREATE INDEX IX_Interview_Interviewer ON rp.Interview(InterviewerEmployeeId, ScheduledAt) WHERE Status = 'Scheduled';

CREATE TABLE rp.InterviewFeedback(
    FeedbackId              int IDENTITY(1,1) CONSTRAINT PK_InterviewFeedback PRIMARY KEY,
    InterviewId             int NOT NULL CONSTRAINT FK_Feedback_Interview REFERENCES rp.Interview(InterviewId) ON DELETE CASCADE,
    SubmittedByEmployeeId   int NOT NULL CONSTRAINT FK_Feedback_By REFERENCES rp.Employee(EmployeeId),
    Rating                  tinyint NOT NULL CONSTRAINT CK_Feedback_Rating CHECK (Rating BETWEEN 1 AND 5),
    Recommendation          varchar(8) NOT NULL CONSTRAINT CK_Feedback_Rec CHECK (Recommendation IN ('Proceed','Reject')),
    Comments                nvarchar(2000) NOT NULL,
    SubmittedOn             datetime2(0) NOT NULL CONSTRAINT DF_Feedback_On DEFAULT SYSUTCDATETIME(),
    CONSTRAINT UQ_Feedback_Once UNIQUE (InterviewId, SubmittedByEmployeeId)
);

/* ---------- Final approval: Division head and HR head both required ---------- */
CREATE TABLE rp.FinalApproval(
    ReferralId          int NOT NULL CONSTRAINT FK_FinalApproval_Referral REFERENCES rp.Referral(ReferralId) ON DELETE CASCADE,
    ApproverRole        varchar(10) NOT NULL CONSTRAINT CK_FinalApproval_Role CHECK (ApproverRole IN ('DIV_HEAD','HR_HEAD')),
    Decision            varchar(10) NOT NULL CONSTRAINT DF_FinalApproval_Decision DEFAULT 'Pending'
                        CONSTRAINT CK_FinalApproval_Decision CHECK (Decision IN ('Pending','Approved','Rejected')),
    DecidedByEmployeeId int NULL CONSTRAINT FK_FinalApproval_By REFERENCES rp.Employee(EmployeeId),
    DecidedOn           datetime2(0) NULL,
    Comments            nvarchar(500) NULL,
    CONSTRAINT PK_FinalApproval PRIMARY KEY (ReferralId, ApproverRole)
);
CREATE INDEX IX_FinalApproval_Inbox ON rp.FinalApproval(ApproverRole) WHERE Decision = 'Pending';

/* ---------- Notification outbox (email / Teams via Graph or Power Automate) ---------- */
CREATE TABLE rp.Notification(
    NotificationId      bigint IDENTITY(1,1) CONSTRAINT PK_Notification PRIMARY KEY,
    RecipientEmployeeId int NOT NULL CONSTRAINT FK_Notification_To REFERENCES rp.Employee(EmployeeId),
    Channel             varchar(8) NOT NULL CONSTRAINT DF_Notification_Channel DEFAULT 'Email'
                        CONSTRAINT CK_Notification_Channel CHECK (Channel IN ('Email','Teams')),
    Subject             nvarchar(200) NOT NULL,
    Body                nvarchar(max) NOT NULL,
    RelatedEntity       varchar(20) NULL,       -- Requisition, Referral, Interview
    RelatedId           int NULL,
    Status              varchar(8) NOT NULL CONSTRAINT DF_Notification_Status DEFAULT 'Queued'
                        CONSTRAINT CK_Notification_Status CHECK (Status IN ('Queued','Sent','Failed')),
    CreatedOn           datetime2(0) NOT NULL CONSTRAINT DF_Notification_Created DEFAULT SYSUTCDATETIME(),
    SentOn              datetime2(0) NULL
);
CREATE INDEX IX_Notification_Queue ON rp.Notification(CreatedOn) WHERE Status = 'Queued';
GO

/* ---------- Seed data ---------- */
INSERT rp.AppRole(RoleId, Code, Name) VALUES
 (1,'EMPLOYEE','Employee'),(2,'MANAGER','Hiring manager'),(3,'DEPT_HEAD','Department head'),
 (4,'DIV_HEAD','Division head'),(5,'HR_HEAD','HR head'),(6,'HR','HR team');

INSERT rp.ReferralStage(StageId, Name) VALUES
 (1,'Referred'),(2,'HR review'),(3,'Interview 1'),(4,'Interview 2'),(5,'Interview 3'),(6,'Final approval');

INSERT rp.Location(Name) VALUES ('Rajkot'),('Ahmedabad'),('Remote');
GO

/* ---------- Views used by the API ---------- */
CREATE VIEW rp.vw_OpenPositions AS
SELECT  j.RequisitionId, j.Title, d.Name AS Department, l.Name AS Location, j.EmploymentType,
        j.NoOfPositions, j.PositionsFilled, j.NoOfPositions - j.PositionsFilled AS OpenPositions,
        j.ExperienceRequired, j.HireByDate, DATEDIFF(DAY, CAST(SYSUTCDATETIME() AS date), j.HireByDate) AS DaysLeft,
        j.Summary, j.Requirements,
        (SELECT STRING_AGG(s.Name, ', ') FROM rp.JobRequisitionSkill rs JOIN rp.Skill s ON s.SkillId = rs.SkillId
          WHERE rs.RequisitionId = j.RequisitionId) AS Skills,
        (SELECT COUNT(*) FROM rp.Referral r WHERE r.RequisitionId = j.RequisitionId) AS ReferralCount
FROM    rp.JobRequisition j
JOIN    rp.Department d ON d.DepartmentId = j.DepartmentId
JOIN    rp.Location l   ON l.LocationId = j.LocationId
WHERE   j.Status = 'Open';
GO

CREATE VIEW rp.vw_Pipeline AS
SELECT  r.ReferralId, r.CandidateName, r.CandidateEmail, j.RequisitionId, j.Title AS JobTitle,
        e.DisplayName AS ReferredBy, r.ReferredByEmployeeId, r.ReferredOn,
        r.CurrentStageId, s.Name AS StageName, r.Outcome,
        (SELECT AVG(CAST(f.Rating AS decimal(3,1))) FROM rp.Interview i
           JOIN rp.InterviewFeedback f ON f.InterviewId = i.InterviewId WHERE i.ReferralId = r.ReferralId) AS AvgRating,
        (SELECT MIN(i.ScheduledAt) FROM rp.Interview i
          WHERE i.ReferralId = r.ReferralId AND i.Status = 'Scheduled') AS NextInterviewAt
FROM    rp.Referral r
JOIN    rp.JobRequisition j ON j.RequisitionId = r.RequisitionId
JOIN    rp.Employee e       ON e.EmployeeId = r.ReferredByEmployeeId
JOIN    rp.ReferralStage s  ON s.StageId = r.CurrentStageId;
GO
