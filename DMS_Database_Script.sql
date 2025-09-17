-- =============================================
-- DMS (Document Management System) Database Script
-- SQL Server Database Schema
-- Generated from DMSServices Application Analysis
-- =============================================

-- Create Database
USE master;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'DMS')
BEGIN
    CREATE DATABASE DMS;
END
GO

USE DMS;
GO

-- =============================================
-- Core Tables
-- =============================================

-- Documents Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Documents' AND xtype='U')
BEGIN
    CREATE TABLE Documents (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        name NVARCHAR(255) NOT NULL,
        description NVARCHAR(MAX),
        dfilename NVARCHAR(255),
        ext_id NVARCHAR(50),
        deleted DATETIME NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- Document_Versions Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Document_Versions' AND xtype='U')
BEGIN
    CREATE TABLE Document_Versions (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        doc_id INT NOT NULL,
        dimage VARBINARY(MAX),
        dsize INT,
        backed_up DATETIME NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1,
        FOREIGN KEY (doc_id) REFERENCES Documents(row_id)
    );
END
GO

-- Document_Types Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Document_Types' AND xtype='U')
BEGIN
    CREATE TABLE Document_Types (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        extension NVARCHAR(10) NOT NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- =============================================
-- Classification Tables
-- =============================================

-- Categories Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Categories' AND xtype='U')
BEGIN
    CREATE TABLE Categories (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        name NVARCHAR(255) NOT NULL,
        public_flag CHAR(1) DEFAULT 'Y',
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- Keywords Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Keywords' AND xtype='U')
BEGIN
    CREATE TABLE Keywords (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        name NVARCHAR(255) NOT NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- Associations Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Associations' AND xtype='U')
BEGIN
    CREATE TABLE Associations (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        name NVARCHAR(255) NOT NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- =============================================
-- User Management Tables
-- =============================================

-- Users Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Users' AND xtype='U')
BEGIN
    CREATE TABLE Users (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        ext_user_id NVARCHAR(50),
        ext_id NVARCHAR(50),
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- Groups Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Groups' AND xtype='U')
BEGIN
    CREATE TABLE Groups (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        name NVARCHAR(255) NOT NULL,
        type_cd NVARCHAR(50),
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- User_Group_Access Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='User_Group_Access' AND xtype='U')
BEGIN
    CREATE TABLE User_Group_Access (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        type_id CHAR(1) NOT NULL, -- 'U' for User, 'G' for Group
        access_id INT NOT NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- User_Sessions Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='User_Sessions' AND xtype='U')
BEGIN
    CREATE TABLE User_Sessions (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        user_id NVARCHAR(50) NOT NULL,
        session_key NVARCHAR(255) NOT NULL,
        machine_id NVARCHAR(50),
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1
    );
END
GO

-- =============================================
-- Document Relationship Tables
-- =============================================

-- Document_Categories Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Document_Categories' AND xtype='U')
BEGIN
    CREATE TABLE Document_Categories (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        doc_id INT NOT NULL,
        cat_id INT NOT NULL,
        pr_flag CHAR(1) DEFAULT 'N', -- Primary flag
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1,
        FOREIGN KEY (doc_id) REFERENCES Documents(row_id),
        FOREIGN KEY (cat_id) REFERENCES Categories(row_id)
    );
END
GO

-- Document_Keywords Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Document_Keywords' AND xtype='U')
BEGIN
    CREATE TABLE Document_Keywords (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        doc_id INT NOT NULL,
        key_id INT NOT NULL,
        pr_flag CHAR(1) DEFAULT 'N', -- Primary flag
        val NVARCHAR(255),
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1,
        FOREIGN KEY (doc_id) REFERENCES Documents(row_id),
        FOREIGN KEY (key_id) REFERENCES Keywords(row_id)
    );
END
GO

-- Document_Associations Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Document_Associations' AND xtype='U')
BEGIN
    CREATE TABLE Document_Associations (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        doc_id INT NOT NULL,
        association_id INT NOT NULL,
        fkey NVARCHAR(255) NOT NULL,
        pr_flag CHAR(1) DEFAULT 'N', -- Primary flag
        access_flag CHAR(1) DEFAULT 'Y',
        access_type NVARCHAR(50),
        reqd_flag CHAR(1) DEFAULT 'N', -- Required flag
        expiration DATETIME NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1,
        FOREIGN KEY (doc_id) REFERENCES Documents(row_id),
        FOREIGN KEY (association_id) REFERENCES Associations(row_id)
    );
END
GO

-- Document_Users Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Document_Users' AND xtype='U')
BEGIN
    CREATE TABLE Document_Users (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        doc_id INT NOT NULL,
        user_access_id INT NOT NULL,
        owner_flag CHAR(1) DEFAULT 'N',
        access_type NVARCHAR(50),
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1,
        FOREIGN KEY (doc_id) REFERENCES Documents(row_id),
        FOREIGN KEY (user_access_id) REFERENCES User_Group_Access(row_id)
    );
END
GO

-- Category_Keywords Table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Category_Keywords' AND xtype='U')
BEGIN
    CREATE TABLE Category_Keywords (
        row_id INT IDENTITY(1,1) PRIMARY KEY,
        cat_id INT NOT NULL,
        key_id INT NOT NULL,
        created DATETIME DEFAULT GETDATE(),
        created_by INT DEFAULT 1,
        last_upd DATETIME DEFAULT GETDATE(),
        last_upd_by INT DEFAULT 1,
        FOREIGN KEY (cat_id) REFERENCES Categories(row_id),
        FOREIGN KEY (key_id) REFERENCES Keywords(row_id)
    );
END
GO

-- =============================================
-- Indexes for Performance
-- =============================================

-- Documents indexes
CREATE NONCLUSTERED INDEX IX_Documents_name ON Documents(name);
CREATE NONCLUSTERED INDEX IX_Documents_ext_id ON Documents(ext_id);
CREATE NONCLUSTERED INDEX IX_Documents_deleted ON Documents(deleted);

-- Document_Versions indexes
CREATE NONCLUSTERED INDEX IX_Document_Versions_doc_id ON Document_Versions(doc_id);

-- Document_Associations indexes
CREATE NONCLUSTERED INDEX IX_Document_Associations_doc_id ON Document_Associations(doc_id);
CREATE NONCLUSTERED INDEX IX_Document_Associations_association_id ON Document_Associations(association_id);
CREATE NONCLUSTERED INDEX IX_Document_Associations_fkey ON Document_Associations(fkey);
CREATE NONCLUSTERED INDEX IX_Document_Associations_pr_flag ON Document_Associations(pr_flag);

-- Document_Categories indexes
CREATE NONCLUSTERED INDEX IX_Document_Categories_doc_id ON Document_Categories(doc_id);
CREATE NONCLUSTERED INDEX IX_Document_Categories_cat_id ON Document_Categories(cat_id);
CREATE NONCLUSTERED INDEX IX_Document_Categories_pr_flag ON Document_Categories(pr_flag);

-- Document_Keywords indexes
CREATE NONCLUSTERED INDEX IX_Document_Keywords_doc_id ON Document_Keywords(doc_id);
CREATE NONCLUSTERED INDEX IX_Document_Keywords_key_id ON Document_Keywords(key_id);
CREATE NONCLUSTERED INDEX IX_Document_Keywords_pr_flag ON Document_Keywords(pr_flag);

-- Document_Users indexes
CREATE NONCLUSTERED INDEX IX_Document_Users_doc_id ON Document_Users(doc_id);
CREATE NONCLUSTERED INDEX IX_Document_Users_user_access_id ON Document_Users(user_access_id);

-- User_Group_Access indexes
CREATE NONCLUSTERED INDEX IX_User_Group_Access_type_id ON User_Group_Access(type_id);
CREATE NONCLUSTERED INDEX IX_User_Group_Access_access_id ON User_Group_Access(access_id);

-- User_Sessions indexes
CREATE NONCLUSTERED INDEX IX_User_Sessions_user_id ON User_Sessions(user_id);
CREATE NONCLUSTERED INDEX IX_User_Sessions_session_key ON User_Sessions(session_key);

-- Users indexes
CREATE NONCLUSTERED INDEX IX_Users_ext_user_id ON Users(ext_user_id);
CREATE NONCLUSTERED INDEX IX_Users_ext_id ON Users(ext_id);

-- Groups indexes
CREATE NONCLUSTERED INDEX IX_Groups_name ON Groups(name);
CREATE NONCLUSTERED INDEX IX_Groups_type_cd ON Groups(type_cd);

-- Categories indexes
CREATE NONCLUSTERED INDEX IX_Categories_name ON Categories(name);
CREATE NONCLUSTERED INDEX IX_Categories_public_flag ON Categories(public_flag);

-- Keywords indexes
CREATE NONCLUSTERED INDEX IX_Keywords_name ON Keywords(name);

-- Associations indexes
CREATE NONCLUSTERED INDEX IX_Associations_name ON Associations(name);

-- Category_Keywords indexes
CREATE NONCLUSTERED INDEX IX_Category_Keywords_cat_id ON Category_Keywords(cat_id);
CREATE NONCLUSTERED INDEX IX_Category_Keywords_key_id ON Category_Keywords(key_id);

-- =============================================
-- Sample Data Insertion
-- =============================================

-- Insert default associations
INSERT INTO Associations (name) VALUES 
('Contact'),
('Partner'),
('Trainer'),
('Employee'),
('Registration');

-- Insert default document types
INSERT INTO Document_Types (extension) VALUES 
('.pdf'),
('.doc'),
('.docx'),
('.xls'),
('.xlsx'),
('.ppt'),
('.pptx'),
('.txt'),
('.jpg'),
('.jpeg'),
('.png'),
('.gif'),
('.bmp'),
('.tiff'),
('.xml');

-- Insert default categories
INSERT INTO Categories (name, public_flag) VALUES 
('General', 'Y'),
('Training', 'Y'),
('Compliance', 'Y'),
('Reports', 'Y'),
('Policies', 'Y');

-- Insert default keywords
INSERT INTO Keywords (name) VALUES 
('Confidential'),
('Public'),
('Internal'),
('External'),
('Required'),
('Optional'),
('Training'),
('Compliance'),
('Policy'),
('Procedure'),
('Manual'),
('Guide'),
('Form'),
('Report'),
('Certificate');

-- Insert default groups
INSERT INTO Groups (name, type_cd) VALUES 
('System Administrators', 'System'),
('Domain Users', 'Domain'),
('Subscription Users', 'Subscription');

-- =============================================
-- Database Users and Permissions
-- =============================================

-- Create DMS user
IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = 'DMS')
BEGIN
    CREATE LOGIN DMS WITH PASSWORD = '5241200';
END
GO

-- Create database user
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = 'DMS')
BEGIN
    CREATE USER DMS FOR LOGIN DMS;
END
GO

-- Grant permissions
ALTER ROLE db_datareader ADD MEMBER DMS;
ALTER ROLE db_datawriter ADD MEMBER DMS;
ALTER ROLE db_ddladmin ADD MEMBER DMS;

-- =============================================
-- Stored Procedures (Optional)
-- =============================================

-- Procedure to get document count for a user
CREATE OR ALTER PROCEDURE GetDocumentCount
    @ContactId NVARCHAR(50),
    @TrainerNum NVARCHAR(50) = NULL,
    @PartId NVARCHAR(50) = NULL,
    @MtId NVARCHAR(50) = NULL,
    @Domain NVARCHAR(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @DocCount INT = 0;
    
    -- Complex query to count accessible documents
    -- This is a simplified version of the logic found in the application
    SELECT @DocCount = COUNT(DISTINCT D.row_id)
    FROM Documents D
    LEFT OUTER JOIN Document_Categories DC ON DC.doc_id = D.row_id
    LEFT OUTER JOIN Category_Keywords CK ON CK.cat_id = DC.cat_id
    INNER JOIN Document_Associations DA ON DA.doc_id = D.row_id
    INNER JOIN Document_Users DU ON DU.doc_id = D.row_id
    WHERE DC.pr_flag = 'Y' 
    AND (CK.key_id IN (3,5,7,8,13,15,16,14))
    AND (
        (DA.association_id = 3 AND DA.fkey = @ContactId AND DA.pr_flag = 'Y') OR
        (@TrainerNum IS NOT NULL AND DA.association_id = 5 AND DA.fkey = @TrainerNum AND DA.pr_flag = 'Y') OR
        (@PartId IS NOT NULL AND DA.association_id = 4 AND DA.fkey = @PartId AND DA.pr_flag = 'Y') OR
        (@MtId IS NOT NULL AND DA.association_id = 37 AND DA.fkey = @MtId AND DA.pr_flag = 'Y')
    )
    AND D.deleted IS NULL;
    
    SELECT @DocCount AS DocumentCount;
END
GO

-- =============================================
-- Script Completion
-- =============================================

PRINT 'DMS Database schema created successfully!';
PRINT 'Database: DMS';
PRINT 'Tables created: 15';
PRINT 'Indexes created: 25+';
PRINT 'Sample data inserted';
PRINT 'User DMS created with appropriate permissions';
