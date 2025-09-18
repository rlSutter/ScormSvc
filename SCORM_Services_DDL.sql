-- =====================================================
-- SCORM Learning Support Services Database Schema
-- SQL Server DDL Script
-- =====================================================
-- This script creates the database schema for the SCORM Learning Support
-- web services that integrate with SCORM course players and content storage.

-- =====================================================
-- Database Creation
-- =====================================================
USE master;
GO

-- Create databases if they don't exist
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'siebeldb')
    CREATE DATABASE siebeldb;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'elearning')
    CREATE DATABASE elearning;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'DMS')
    CREATE DATABASE DMS;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'reports')
    CREATE DATABASE reports;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'scormmedia')
    CREATE DATABASE scormmedia;
GO

-- =====================================================
-- Siebel Database Schema (Main Application Database)
-- =====================================================
USE siebeldb;
GO

-- Contact/User Management Tables
CREATE TABLE S_CONTACT (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    FST_NAME NVARCHAR(50),
    LAST_NAME NVARCHAR(50),
    EMAIL_ADDR NVARCHAR(100),
    LOGIN NVARCHAR(50),
    X_PASSWORD NVARCHAR(50),
    X_REGISTRATION_NUM NVARCHAR(20),
    X_PR_LANG_CD NVARCHAR(5) DEFAULT 'ENU',
    PR_DEPT_OU_ID NVARCHAR(15),
    PR_PER_ADDR_ID NVARCHAR(15),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- Subscription Management
CREATE TABLE CX_SUBSCRIPTION (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    DOMAIN NVARCHAR(20),
    SVC_TYPE NVARCHAR(50),
    SVC_TERM_DT DATETIME,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- Subscription Contact Association
CREATE TABLE CX_SUB_CON (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    SUB_ID NVARCHAR(15),
    CON_ID NVARCHAR(15),
    USER_EXP_DATE DATETIME,
    TRAINING_ACCESS NVARCHAR(10),
    TRAINER_ACC_FLG NVARCHAR(1),
    PAID_USER_FLG NVARCHAR(1),
    LAST_INST NVARCHAR(500),
    LAST_LOGIN DATETIME,
    LAST_SESS_ID NVARCHAR(50),
    TOKEN NVARCHAR(100),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (SUB_ID) REFERENCES CX_SUBSCRIPTION(ROW_ID),
    FOREIGN KEY (CON_ID) REFERENCES S_CONTACT(ROW_ID)
);

-- Subscription Domain Configuration
CREATE TABLE CX_SUB_DOMAIN (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    DOMAIN NVARCHAR(20),
    HOME_URL NVARCHAR(200),
    DEF_SUB_ID NVARCHAR(15),
    UNSUB_URL NVARCHAR(200),
    LOGOUT_URL NVARCHAR(200),
    ETIPS_FLG NVARCHAR(1),
    SRC_URL NVARCHAR(100),
    CS_EMAIL NVARCHAR(100),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- Session History for Login Tracking
CREATE TABLE CX_SUB_CON_HIST (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    SUB_CON_ID NVARCHAR(15),
    USER_ID NVARCHAR(20),
    SESSION_ID NVARCHAR(50),
    REMOTE_ADDR NVARCHAR(50),
    LOGOUT_DT DATETIME,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (SUB_CON_ID) REFERENCES CX_SUB_CON(ROW_ID)
);

-- Course Management
CREATE TABLE S_CRSE (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    NAME NVARCHAR(200),
    X_SCORM_FLG NVARCHAR(1),
    X_FORMAT NVARCHAR(20),
    X_LANG_CD NVARCHAR(5),
    X_CRSE_CONTENT_URL NVARCHAR(500),
    X_EXAM_REQD NVARCHAR(1),
    X_RESOLUTION NVARCHAR(20),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- Training Offerings
CREATE TABLE CX_TRAIN_OFFR (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    CRSE_ID NVARCHAR(15),
    TRAIN_TYPE NVARCHAR(20),
    DOMAIN NVARCHAR(20),
    MS_IDENT NVARCHAR(50),
    STATUS_CD NVARCHAR(20),
    ALLOWED_REFERRER NVARCHAR(200),
    LANG_ID NVARCHAR(5),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (CRSE_ID) REFERENCES S_CRSE(ROW_ID)
);

-- Session Registrations
CREATE TABLE CX_SESS_REG (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    TRAIN_OFFR_ID NVARCHAR(15),
    CONTACT_ID NVARCHAR(15),
    CRSE_ID NVARCHAR(15),
    STATUS_CD NVARCHAR(20),
    JURIS_ID NVARCHAR(15),
    REF_CON_ID NVARCHAR(15),
    OU_ID NVARCHAR(15),
    ADDR_ID NVARCHAR(15),
    PER_ADDR_ID NVARCHAR(15),
    RETAKE_FLG NVARCHAR(1),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (TRAIN_OFFR_ID) REFERENCES CX_TRAIN_OFFR(ROW_ID),
    FOREIGN KEY (CONTACT_ID) REFERENCES S_CONTACT(ROW_ID),
    FOREIGN KEY (CRSE_ID) REFERENCES S_CRSE(ROW_ID)
);

-- Training Access Control
CREATE TABLE CX_TRAIN_OFFR_ACCESS (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    REG_ID NVARCHAR(15),
    ENTER_FLG NVARCHAR(1),
    EXIT_FLG NVARCHAR(1),
    MOBILE NVARCHAR(1),
    CALL_ID NVARCHAR(50),
    CALL_SCREEN NVARCHAR(50),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (REG_ID) REFERENCES CX_SESS_REG(ROW_ID)
);

-- Course Tests
CREATE TABLE S_CRSE_TST (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    CRSE_ID NVARCHAR(15),
    X_ENGINE NVARCHAR(20),
    X_SURVEY_FLG NVARCHAR(1),
    X_ANON_REF_ID NVARCHAR(50),
    X_LANG_ID NVARCHAR(5),
    X_CONTINUABLE_FLG NVARCHAR(1),
    X_KBA_QUES_NUM INT,
    MAX_POINTS INT,
    X_TIME_ALLOWED INT,
    PASSING_SCORE DECIMAL(5,2),
    STATUS_CD NVARCHAR(20),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (CRSE_ID) REFERENCES S_CRSE(ROW_ID)
);

-- Test Runs
CREATE TABLE S_CRSE_TSTRUN (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    CRSE_TST_ID NVARCHAR(15),
    X_PART_ID NVARCHAR(15),
    PERSON_ID NVARCHAR(15),
    X_REDIRECT_URL NVARCHAR(500),
    STATUS_CD NVARCHAR(20),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (CRSE_TST_ID) REFERENCES S_CRSE_TST(ROW_ID)
);

-- Test Access Control
CREATE TABLE S_CRSE_TSTRUN_ACCESS (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    CRSE_TSTRUN_ID NVARCHAR(15),
    ENTER_FLG NVARCHAR(1),
    EXIT_FLG NVARCHAR(1),
    CALL_ID NVARCHAR(50),
    CALL_SCREEN NVARCHAR(50),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (CRSE_TSTRUN_ID) REFERENCES S_CRSE_TSTRUN(ROW_ID)
);

-- Course Test Questions
CREATE TABLE S_CRSE_TST_QUES (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    CRSE_TST_ID NVARCHAR(15),
    QUES_TYPE_CD NVARCHAR(20),
    QUES_SEQ_NUM INT,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (CRSE_TST_ID) REFERENCES S_CRSE_TST(ROW_ID)
);

-- Jurisdiction Management
CREATE TABLE CX_JURISDICTION_X (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    NAME NVARCHAR(100),
    JURIS_LVL NVARCHAR(20),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- Jurisdiction Course Configuration
CREATE TABLE CX_JURIS_CRSE (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    JURIS_ID NVARCHAR(15),
    CRSE_ID NVARCHAR(15),
    KBA_REQD NVARCHAR(1),
    KBA_QUESTIONS INT,
    KBA_NOTICE NVARCHAR(500),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (JURIS_ID) REFERENCES CX_JURISDICTION_X(ROW_ID),
    FOREIGN KEY (CRSE_ID) REFERENCES S_CRSE(ROW_ID)
);

-- SCORM Course Attributes
CREATE TABLE CX_CRSE_SCORM_ATTR (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    CRSE_ID NVARCHAR(15),
    PRE_CRSE_DOC_ID INT,
    POST_CRSE_DOC_ID INT,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (CRSE_ID) REFERENCES S_CRSE(ROW_ID)
);

-- Organization Management
CREATE TABLE S_ORG_EXT (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    PR_ADDR_ID NVARCHAR(15),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- Address Management
CREATE TABLE S_ADDR_ORG (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    X_JURIS_ID NVARCHAR(15),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

CREATE TABLE S_ADDR_PER (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    X_JURIS_ID NVARCHAR(15),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- =====================================================
-- E-Learning Database Schema
-- =====================================================
USE elearning;
GO

-- KBA (Knowledge-Based Assessment) Questions
CREATE TABLE KBA_QUES (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    QUES_TEXT NVARCHAR(MAX),
    LANG_CD NVARCHAR(5),
    ACTIVE_FLG NVARCHAR(1) DEFAULT 'Y',
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- KBA Jurisdiction Mapping
CREATE TABLE KBA_JURIS (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    QUES_ID NVARCHAR(15),
    JURIS_ID NVARCHAR(15),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (QUES_ID) REFERENCES KBA_QUES(ROW_ID)
);

-- KBA Course Mapping
CREATE TABLE KBA_CRSE (
    ROW_ID NVARCHAR(15) PRIMARY KEY,
    QUES_ID NVARCHAR(15),
    CRSE_ID NVARCHAR(15),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (QUES_ID) REFERENCES KBA_QUES(ROW_ID)
);

-- KBA Answers
CREATE TABLE KBA_ANSR (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    REG_ID NVARCHAR(15),
    USER_ID NVARCHAR(20),
    QUES_ID NVARCHAR(15),
    ANSR_TEXT NVARCHAR(MAX),
    ENC_ANSR_TEXT NVARCHAR(MAX),
    VAL_TEXT NVARCHAR(MAX),
    VAL_DATE DATETIME,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0,
    FOREIGN KEY (QUES_ID) REFERENCES KBA_QUES(ROW_ID)
);

-- E-Learning Player Data
CREATE TABLE Elearning_Player_Data (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    reg_id NVARCHAR(15),
    crse_id NVARCHAR(15),
    crse_type NVARCHAR(1),
    from_db NVARCHAR(20),
    player_reg_id NVARCHAR(50),
    completion_status NVARCHAR(20),
    entry NVARCHAR(20),
    exit_mode NVARCHAR(20),
    success_status NVARCHAR(20),
    activity_absolute_dur NVARCHAR(50),
    score_scaled DECIMAL(5,4),
    shell_exit_dt DATETIME,
    active BIT DEFAULT 0,
    suspended BIT DEFAULT 0,
    update_dt DATETIME DEFAULT GETDATE(),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning Package Data
CREATE TABLE Elearning_Package_Data (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    crse_id NVARCHAR(15),
    package_id NVARCHAR(50),
    version_id INT,
    from_db NVARCHAR(20),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning Course Map
CREATE TABLE ELN_COURSE_MAP (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    CRSE_ID NVARCHAR(15),
    XML_CRSE_ID NVARCHAR(50),
    NUM_ELEMENTS INT,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning User Progress
CREATE TABLE ELN_USER_PROGRESS (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    SESS_REG_ID NVARCHAR(15),
    HIGH_WATER INT,
    START_DATE DATETIME,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning Course Elements
CREATE TABLE ELN_COURSE_ELEMENT (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    CRSE_ID NVARCHAR(15),
    ELEMENT_ID NVARCHAR(50),
    ELEMENT_NAME NVARCHAR(200),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning Test Answers
CREATE TABLE ELN_TEST_ANSWER (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    REG_ID NVARCHAR(15),
    QUESTION_ID NVARCHAR(50),
    ANSWER_TEXT NVARCHAR(MAX),
    IS_CORRECT BIT,
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning Feedback
CREATE TABLE ELN_FEEDBACK (
    ID INT IDENTITY(1,1) PRIMARY KEY,
    USER_NAME NVARCHAR(50),
    PAGE NVARCHAR(100),
    COMM NVARCHAR(MAX),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- E-Learning Notes
CREATE TABLE ELN_NOTES (
    ROW_ID NVARCHAR(50) PRIMARY KEY,
    REG_ID NVARCHAR(15),
    USER_ID NVARCHAR(20),
    NOTE_TEXT NVARCHAR(MAX),
    SCREEN NVARCHAR(100),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- =====================================================
-- Document Management System (DMS) Schema
-- =====================================================
USE DMS;
GO

-- Documents
CREATE TABLE Documents (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    name NVARCHAR(200),
    description NVARCHAR(MAX),
    data_type_id INT,
    deleted DATETIME NULL,
    old_row_id NVARCHAR(50),
    dfilename NVARCHAR(200),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Document Versions
CREATE TABLE Document_Versions (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    document_id INT,
    version_number INT,
    dsize BIGINT,
    dimage VARBINARY(MAX),
    minio_flg BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0,
    FOREIGN KEY (document_id) REFERENCES Documents(row_id)
);

-- Document Categories
CREATE TABLE Categories (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    name NVARCHAR(100),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Document Types
CREATE TABLE Document_Types (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    name NVARCHAR(100),
    extension NVARCHAR(10),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Document Categories Association
CREATE TABLE Document_Categories (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    doc_id INT,
    cat_id INT,
    pr_flag NVARCHAR(1),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0,
    FOREIGN KEY (doc_id) REFERENCES Documents(row_id),
    FOREIGN KEY (cat_id) REFERENCES Categories(row_id)
);

-- Associations
CREATE TABLE Associations (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    name NVARCHAR(100),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Document Associations
CREATE TABLE Document_Associations (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    doc_id INT,
    association_id INT,
    fkey NVARCHAR(50),
    pr_flag NVARCHAR(1),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0,
    FOREIGN KEY (doc_id) REFERENCES Documents(row_id),
    FOREIGN KEY (association_id) REFERENCES Associations(row_id)
);

-- User Access Control
CREATE TABLE User_Access (
    user_access_id INT IDENTITY(1,1) PRIMARY KEY,
    user_id NVARCHAR(50),
    access_type NVARCHAR(20),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Document User Access
CREATE TABLE Document_User_Access (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    document_id INT,
    user_access_id INT,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0,
    FOREIGN KEY (document_id) REFERENCES Documents(row_id),
    FOREIGN KEY (user_access_id) REFERENCES User_Access(user_access_id)
);

-- User Sessions
CREATE TABLE User_Sessions (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    user_id NVARCHAR(20),
    session_key NVARCHAR(50),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- =====================================================
-- SCORM Media Services Database Schema
-- =====================================================
USE scormmedia;
GO

-- User Sessions for authentication tracking
CREATE TABLE User_Sessions (
    session_id NVARCHAR(50) PRIMARY KEY,
    user_id NVARCHAR(50),
    login_time DATETIME DEFAULT GETDATE(),
    last_activity DATETIME DEFAULT GETDATE(),
    ip_address NVARCHAR(50),
    user_agent NVARCHAR(500),
    is_active BIT DEFAULT 1,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Course Media Assets
CREATE TABLE Course_Media_Assets (
    asset_id NVARCHAR(50) PRIMARY KEY,
    course_id NVARCHAR(50),
    asset_name NVARCHAR(200),
    asset_type NVARCHAR(20),
    file_extension NVARCHAR(10),
    content_type NVARCHAR(100),
    file_size BIGINT,
    is_public BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Course Resources
CREATE TABLE Course_Resources (
    resource_id NVARCHAR(50) PRIMARY KEY,
    course_id NVARCHAR(50),
    resource_name NVARCHAR(200),
    resource_type NVARCHAR(20),
    file_extension NVARCHAR(10),
    content_type NVARCHAR(100),
    file_size BIGINT,
    is_public BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Course Images
CREATE TABLE Course_Images (
    image_id NVARCHAR(50) PRIMARY KEY,
    course_id NVARCHAR(50),
    image_name NVARCHAR(200),
    image_type NVARCHAR(20),
    file_extension NVARCHAR(10),
    content_type NVARCHAR(100),
    file_size BIGINT,
    is_public BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Document Images (for DMS integration)
CREATE TABLE Document_Images (
    doc_image_id NVARCHAR(50) PRIMARY KEY,
    document_id NVARCHAR(50),
    domain NVARCHAR(50),
    public_key NVARCHAR(100),
    image_name NVARCHAR(200),
    file_extension NVARCHAR(10),
    content_type NVARCHAR(100),
    file_size BIGINT,
    is_public BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Secure Images (for authenticated access)
CREATE TABLE Secure_Images (
    secure_image_id NVARCHAR(50) PRIMARY KEY,
    user_id NVARCHAR(50),
    domain NVARCHAR(50),
    public_key NVARCHAR(100),
    user_key NVARCHAR(100),
    image_name NVARCHAR(200),
    file_extension NVARCHAR(10),
    content_type NVARCHAR(100),
    file_size BIGINT,
    is_public BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Course Help Media
CREATE TABLE Course_Help_Media (
    help_media_id NVARCHAR(50) PRIMARY KEY,
    course_id NVARCHAR(50),
    media_name NVARCHAR(200),
    media_type NVARCHAR(20),
    file_extension NVARCHAR(10),
    content_type NVARCHAR(100),
    file_size BIGINT,
    is_public BIT DEFAULT 0,
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- Media Access Log
CREATE TABLE Media_Access_Log (
    log_id INT IDENTITY(1,1) PRIMARY KEY,
    asset_type NVARCHAR(20),
    asset_id NVARCHAR(50),
    user_id NVARCHAR(50),
    session_id NVARCHAR(50),
    ip_address NVARCHAR(50),
    user_agent NVARCHAR(500),
    access_time DATETIME DEFAULT GETDATE(),
    success BIT DEFAULT 1,
    error_message NVARCHAR(500),
    created DATETIME DEFAULT GETDATE(),
    last_upd DATETIME DEFAULT GETDATE(),
    created_by NVARCHAR(15),
    last_upd_by NVARCHAR(15),
    modification_num INT DEFAULT 0,
    conflict_id INT DEFAULT 0
);

-- =====================================================
-- Reports Database Schema
-- =====================================================
USE reports;
GO

-- Certification Manager Log
CREATE TABLE CM_LOG (
    ROW_ID INT IDENTITY(1,1) PRIMARY KEY,
    REG_ID NVARCHAR(20),
    SESSION_ID NVARCHAR(50),
    RECORD_ID NVARCHAR(50),
    REMOTE_ADDR NVARCHAR(50),
    ACTION NVARCHAR(100),
    BROWSER NVARCHAR(500),
    CREATED DATETIME DEFAULT GETDATE(),
    LAST_UPD DATETIME DEFAULT GETDATE(),
    CREATED_BY NVARCHAR(15),
    LAST_UPD_BY NVARCHAR(15),
    MODIFICATION_NUM INT DEFAULT 0,
    CONFLICT_ID INT DEFAULT 0
);

-- =====================================================
-- Indexes for Performance
-- =====================================================

-- Siebel Database Indexes
USE siebeldb;
GO

CREATE INDEX IX_S_CONTACT_X_REGISTRATION_NUM ON S_CONTACT(X_REGISTRATION_NUM);
CREATE INDEX IX_S_CONTACT_EMAIL_ADDR ON S_CONTACT(EMAIL_ADDR);
CREATE INDEX IX_S_CONTACT_LOGIN ON S_CONTACT(LOGIN);
CREATE INDEX IX_CX_SUB_CON_CON_ID ON CX_SUB_CON(CON_ID);
CREATE INDEX IX_CX_SUB_CON_SUB_ID ON CX_SUB_CON(SUB_ID);
CREATE INDEX IX_CX_SUB_CON_HIST_USER_SESSION ON CX_SUB_CON_HIST(USER_ID, SESSION_ID);
CREATE INDEX IX_CX_SESS_REG_CONTACT_ID ON CX_SESS_REG(CONTACT_ID);
CREATE INDEX IX_CX_SESS_REG_CRSE_ID ON CX_SESS_REG(CRSE_ID);
CREATE INDEX IX_CX_TRAIN_OFFR_ACCESS_REG_ID ON CX_TRAIN_OFFR_ACCESS(REG_ID);
CREATE INDEX IX_S_CRSE_TSTRUN_ACCESS_CRSE_TSTRUN_ID ON S_CRSE_TSTRUN_ACCESS(CRSE_TSTRUN_ID);
CREATE INDEX IX_S_CRSE_TST_CRSE_ID ON S_CRSE_TST(CRSE_ID);
CREATE INDEX IX_S_CRSE_TSTRUN_CRSE_TST_ID ON S_CRSE_TSTRUN(CRSE_TST_ID);
CREATE INDEX IX_S_CRSE_TSTRUN_PERSON_ID ON S_CRSE_TSTRUN(PERSON_ID);
CREATE INDEX IX_S_CRSE_TST_QUES_CRSE_TST_ID ON S_CRSE_TST_QUES(CRSE_TST_ID);

-- E-Learning Database Indexes
USE elearning;
GO

CREATE INDEX IX_KBA_ANSR_REG_ID ON KBA_ANSR(REG_ID);
CREATE INDEX IX_KBA_ANSR_USER_ID ON KBA_ANSR(USER_ID);
CREATE INDEX IX_KBA_ANSR_QUES_ID ON KBA_ANSR(QUES_ID);
CREATE INDEX IX_Elearning_Player_Data_reg_crse ON Elearning_Player_Data(reg_id, crse_id);
CREATE INDEX IX_Elearning_Player_Data_player_reg_id ON Elearning_Player_Data(player_reg_id);
CREATE INDEX IX_Elearning_Package_Data_crse_id ON Elearning_Package_Data(crse_id);
CREATE INDEX IX_ELN_USER_PROGRESS_SESS_REG_ID ON ELN_USER_PROGRESS(SESS_REG_ID);
CREATE INDEX IX_ELN_COURSE_MAP_CRSE_ID ON ELN_COURSE_MAP(CRSE_ID);
CREATE INDEX IX_ELN_COURSE_ELEMENT_CRSE_ID ON ELN_COURSE_ELEMENT(CRSE_ID);
CREATE INDEX IX_ELN_TEST_ANSWER_REG_ID ON ELN_TEST_ANSWER(REG_ID);
CREATE INDEX IX_ELN_FEEDBACK_USER_PAGE ON ELN_FEEDBACK(USER_NAME, PAGE);
CREATE INDEX IX_ELN_NOTES_REG_USER ON ELN_NOTES(REG_ID, USER_ID);

-- DMS Database Indexes
USE DMS;
GO

CREATE INDEX IX_Documents_name ON Documents(name);
CREATE INDEX IX_Documents_dfilename ON Documents(dfilename);
CREATE INDEX IX_Document_Versions_document_id ON Document_Versions(document_id);
CREATE INDEX IX_Document_Associations_doc_id ON Document_Associations(doc_id);
CREATE INDEX IX_Document_Associations_fkey ON Document_Associations(fkey);
CREATE INDEX IX_Document_User_Access_document_id ON Document_User_Access(document_id);
CREATE INDEX IX_User_Access_user_id ON User_Access(user_id);
CREATE INDEX IX_User_Sessions_user_session ON User_Sessions(user_id, session_key);

-- SCORM Media Database Indexes
USE scormmedia;
GO

CREATE INDEX IX_User_Sessions_user_id ON User_Sessions(user_id);
CREATE INDEX IX_User_Sessions_session_id ON User_Sessions(session_id);
CREATE INDEX IX_Course_Media_Assets_course_id ON Course_Media_Assets(course_id);
CREATE INDEX IX_Course_Resources_course_id ON Course_Resources(course_id);
CREATE INDEX IX_Course_Images_course_id ON Course_Images(course_id);
CREATE INDEX IX_Document_Images_document_id ON Document_Images(document_id);
CREATE INDEX IX_Secure_Images_user_id ON Secure_Images(user_id);
CREATE INDEX IX_Course_Help_Media_course_id ON Course_Help_Media(course_id);
CREATE INDEX IX_Media_Access_Log_asset_id ON Media_Access_Log(asset_id);
CREATE INDEX IX_Media_Access_Log_user_id ON Media_Access_Log(user_id);
CREATE INDEX IX_Media_Access_Log_access_time ON Media_Access_Log(access_time);

-- Reports Database Indexes
USE reports;
GO

CREATE INDEX IX_CM_LOG_REG_ID ON CM_LOG(REG_ID);
CREATE INDEX IX_CM_LOG_SESSION_ID ON CM_LOG(SESSION_ID);
CREATE INDEX IX_CM_LOG_CREATED ON CM_LOG(CREATED);

-- =====================================================
-- Sample Data (Optional - for testing)
-- =====================================================

-- Insert sample domain configuration
USE siebeldb;
GO

INSERT INTO CX_SUB_DOMAIN (ROW_ID, DOMAIN, HOME_URL, DEF_SUB_ID, UNSUB_URL, LOGOUT_URL, ETIPS_FLG, SRC_URL, CS_EMAIL)
VALUES ('1-1', 'TIPS', 'https://www.example.com', '1-1', 'https://www.example.com/unsubscribe', 'https://www.example.com/logout', 'Y', 'example.com', 'support@example.com');

INSERT INTO CX_SUBSCRIPTION (ROW_ID, DOMAIN, SVC_TYPE, SVC_TERM_DT)
VALUES ('1-1', 'TIPS', 'PUBLIC ACCESS', DATEADD(YEAR, 1, GETDATE()));

-- Insert sample associations for DMS
USE DMS;
GO

INSERT INTO Associations (name) VALUES ('Course');
INSERT INTO Associations (name) VALUES ('Jurisdiction');
INSERT INTO Categories (name) VALUES ('Course Materials');

-- =====================================================
-- Stored Procedures for SCORM Operations
-- =====================================================

-- Stored procedure for deleting SCORM registrations
USE elearning;
GO

CREATE PROCEDURE usp_elearning_DeleteRegistration
    @scorm_registration_id NVARCHAR(37),
    @returns INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRY
        -- Delete from Elearning_Player_Data
        DELETE FROM Elearning_Player_Data 
        WHERE player_reg_id = @scorm_registration_id;
        
        SET @returns = @@ROWCOUNT;
    END TRY
    BEGIN CATCH
        SET @returns = -1;
        THROW;
    END CATCH
END;
GO

-- =====================================================
-- End of SCORM Learning Support Services Database Schema
-- =====================================================
