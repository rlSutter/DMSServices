# DMS Services - Customization and Implementation Guide

## Table of Contents
1. [Environment Setup](#environment-setup)
2. [Database Customization](#database-customization)
3. [Application Configuration](#application-configuration)
4. [User Management Customization](#user-management-customization)
5. [Document Classification System](#document-classification-system)
6. [Access Control Customization](#access-control-customization)
7. [External System Integration](#external-system-integration)
8. [UI and Branding Customization](#ui-and-branding-customization)
9. [Performance Tuning](#performance-tuning)
10. [Security Hardening](#security-hardening)
11. [Deployment Strategies](#deployment-strategies)
12. [Maintenance and Updates](#maintenance-and-updates)

## Environment Setup

### Prerequisites
Before implementing DMS Services in your environment, ensure you have:

- **Windows Server 2016/2019/2022** with IIS 10.0+
- **SQL Server 2016/2017/2019/2022** (Standard or Enterprise)
- **.NET Framework 4.8** (latest version recommended)
- **Visual Studio 2019/2022** (for development and customization)
- **Active Directory** (for Windows Authentication)
- **Network access** to external systems (Siebel, Certegrity, etc.)

### Server Requirements
- **CPU**: 4+ cores recommended
- **RAM**: 8GB minimum, 16GB recommended
- **Storage**: SSD recommended for database and file storage
- **Network**: Stable connection to external services

### Installation Steps

#### 1. Database Server Setup
```sql
-- Run the DMS_Database_Script.sql on your SQL Server instance
-- Ensure SQL Server Agent is running for maintenance tasks
-- Configure appropriate backup strategies
```

#### 2. Web Server Setup
```powershell
# Install IIS with required features
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServer
Enable-WindowsOptionalFeature -Online -FeatureName IIS-CommonHttpFeatures
Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpErrors
Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpLogging
Enable-WindowsOptionalFeature -Online -FeatureName IIS-RequestFiltering
Enable-WindowsOptionalFeature -Online -FeatureName IIS-StaticContent
Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASPNET48
```

#### 3. File System Setup
```batch
# Create directory structure
mkdir C:\Inetpub\dms
mkdir C:\Inetpub\dms\work_dir
mkdir C:\Inetpub\dms\work_dir\DOC
mkdir C:\Inetpub\dms\work_dir\jpg
mkdir C:\Inetpub\dms\work_dir\pdf
mkdir C:\Inetpub\dms\work_dir\XML
mkdir C:\Inetpub\dms\logs
```

## Database Customization

### Initial Database Configuration

#### 1. Create Custom Database User
```sql
-- Create a dedicated service account for DMS
CREATE LOGIN [DOMAIN\DMS-Service] FROM WINDOWS;
USE DMS;
CREATE USER [DOMAIN\DMS-Service] FOR LOGIN [DOMAIN\DMS-Service];
ALTER ROLE db_datareader ADD MEMBER [DOMAIN\DMS-Service];
ALTER ROLE db_datawriter ADD MEMBER [DOMAIN\DMS-Service];
ALTER ROLE db_ddladmin ADD MEMBER [DOMAIN\DMS-Service];
```

#### 2. Customize Document Types
```sql
-- Add your organization's specific file types
INSERT INTO Document_Types (extension) VALUES 
('.dwg'),      -- AutoCAD files
('.zip'),      -- Archive files
('.mp4'),      -- Video files
('.wav'),      -- Audio files
('.csv'),      -- Data files
('.json');     -- Configuration files
```

#### 3. Set Up Custom Categories
```sql
-- Define your organization's document categories
INSERT INTO Categories (name, public_flag) VALUES 
('HR Documents', 'N'),
('Financial Reports', 'N'),
('Technical Specifications', 'Y'),
('Legal Documents', 'N'),
('Training Materials', 'Y'),
('Project Documentation', 'Y'),
('Vendor Information', 'N'),
('Customer Data', 'N');
```

#### 4. Configure Custom Keywords
```sql
-- Add organization-specific keywords
INSERT INTO Keywords (name) VALUES 
('Confidential - Internal Only'),
('Public - External Access'),
('Restricted - Management Only'),
('Sensitive - HR Only'),
('Financial - Accounting Only'),
('Technical - Engineering Only'),
('Legal - Compliance Only'),
('Archive - Historical'),
('Active - Current Use'),
('Draft - Under Review');
```

### Database Performance Optimization

#### 1. Create Additional Indexes
```sql
-- Performance indexes for common queries
CREATE NONCLUSTERED INDEX IX_Documents_created_deleted 
ON Documents(created, deleted) 
INCLUDE (name, dfilename);

CREATE NONCLUSTERED INDEX IX_Document_Associations_composite 
ON Document_Associations(association_id, fkey, pr_flag) 
INCLUDE (doc_id, access_flag, reqd_flag);

CREATE NONCLUSTERED INDEX IX_User_Sessions_composite 
ON User_Sessions(user_id, session_key) 
INCLUDE (machine_id, created);
```

#### 2. Partition Large Tables (Optional)
```sql
-- For high-volume environments, consider partitioning
-- Example for Documents table by creation date
CREATE PARTITION FUNCTION PF_Documents_Date (datetime)
AS RANGE RIGHT FOR VALUES 
('2020-01-01', '2021-01-01', '2022-01-01', '2023-01-01', '2024-01-01');

CREATE PARTITION SCHEME PS_Documents_Date
AS PARTITION PF_Documents_Date
TO ([PRIMARY], [PRIMARY], [PRIMARY], [PRIMARY], [PRIMARY], [PRIMARY]);
```

## Application Configuration

### Web.config Customization

#### 1. Connection Strings
```xml
<connectionStrings>
  <!-- Production Database -->
  <add name="dms" 
       connectionString="server=YOUR-SQL-SERVER;uid=DOMAIN\DMS-Service;Integrated Security=true;database=DMS;Min Pool Size=5;Max Pool Size=20;Connect Timeout=30" 
       providerName="System.Data.SqlClient"/>
  
  <!-- Siebel Integration -->
  <add name="siebeldb" 
       connectionString="server=SIEBEL-SERVER;uid=SIEBEL_USER;pwd=SIEBEL_PASSWORD;database=siebeldb;Min Pool Size=3;Max Pool Size=10" 
       providerName="System.Data.SqlClient"/>
  
  <!-- Email Database -->
  <add name="email" 
       connectionString="server=YOUR-SQL-SERVER;uid=DOMAIN\DMS-Service;Integrated Security=true;database=scanner;Min Pool Size=2;Max Pool Size=5" 
       providerName="System.Data.SqlClient"/>
</connectionStrings>
```

#### 2. Application Settings
```xml
<appSettings>
  <!-- File Storage Path -->
  <add key="basepath" value="C:\Inetpub\dms\"/>
  
  <!-- External Service URLs -->
  <add key="processing.com.certegrity.cloudsvc.service" value="https://your-processing-service.certegrity.com/service.asmx"/>
  <add key="basic.com.certegrity.cloudsvc.service" value="https://your-basic-service.certegrity.com/service.asmx"/>
  <add key="local.hq.dms.service" value="https://your-dms-service.certegrity.com/service.asmx"/>
  
  <!-- Debug Settings -->
  <add key="SaveDMSDocDoc_debug" value="N"/>
  <add key="SaveDMSDocAssoc_debug" value="N"/>
  <add key="SaveDMSDocCat_debug" value="N"/>
  <add key="SaveDMSDocKey_debug" value="N"/>
  <add key="SaveDMSDocUser_debug" value="N"/>
  <add key="UpdDMSDocCount_debug" value="N"/>
  <add key="UpdDMSDoc_debug" value="N"/>
  <add key="PublishDMSDoc_debug" value="N"/>
  <add key="CheckDocAccess_debug" value="N"/>
  <add key="UserLogin_debug" value="N"/>
  <add key="UserLogout_debug" value="N"/>
  
  <!-- Custom Settings -->
  <add key="MaxFileSize" value="104857600"/> <!-- 100MB -->
  <add key="AllowedFileTypes" value=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.jpg,.jpeg,.png"/>
  <add key="SessionTimeout" value="30"/> <!-- minutes -->
  <add key="MaxConcurrentSessions" value="100"/>
</appSettings>
```

#### 3. Logging Configuration
```xml
<log4net>
  <!-- Production Logging -->
  <appender name="ProductionLogAppender" type="log4net.Appender.RollingFileAppender">
    <file value="C:\Inetpub\dms\logs\dms.log"/>
    <appendToFile value="true"/>
    <rollingStyle value="Date"/>
    <datePattern value="yyyyMMdd"/>
    <maxSizeRollBackups value="30"/>
    <staticLogFileName value="false"/>
    <layout type="log4net.Layout.PatternLayout">
      <conversionPattern value="%date [%thread] %-5level %logger - %message%newline"/>
    </layout>
  </appender>
  
  <!-- Error Logging -->
  <appender name="ErrorLogAppender" type="log4net.Appender.RollingFileAppender">
    <file value="C:\Inetpub\dms\logs\errors.log"/>
    <appendToFile value="true"/>
    <rollingStyle value="Size"/>
    <maxSizeRollBackups value="10"/>
    <maximumFileSize value="10MB"/>
    <filter type="log4net.Filter.LevelRangeFilter">
      <levelMin value="ERROR"/>
      <levelMax value="FATAL"/>
    </filter>
    <layout type="log4net.Layout.PatternLayout">
      <conversionPattern value="%date [%thread] %-5level %logger - %message%newline"/>
    </layout>
  </appender>
  
  <root>
    <level value="INFO"/>
    <appender-ref ref="ProductionLogAppender"/>
    <appender-ref ref="ErrorLogAppender"/>
  </root>
</log4net>
```

## User Management Customization

### 1. Active Directory Integration

#### Configure Group Mappings
```sql
-- Map AD groups to DMS groups
INSERT INTO Groups (name, type_cd) VALUES 
('DOMAIN\DMS-Administrators', 'System'),
('DOMAIN\DMS-Users', 'Domain'),
('DOMAIN\DMS-Managers', 'Domain'),
('DOMAIN\DMS-ReadOnly', 'Domain'),
('DOMAIN\DMS-Contractors', 'Domain');
```

#### Create User Import Procedure
```sql
CREATE PROCEDURE ImportADUsers
AS
BEGIN
    -- This procedure would integrate with AD to sync users
    -- Implementation depends on your AD structure
    DECLARE @ADQuery NVARCHAR(MAX) = '
    SELECT 
        sAMAccountName,
        displayName,
        mail,
        department,
        title
    FROM OPENQUERY(ADSI, 
        ''SELECT sAMAccountName, displayName, mail, department, title 
          FROM ''''LDAP://DC=yourdomain,DC=com''''
          WHERE objectClass = ''''user'''' AND objectCategory = ''''person''''
        '')';
    
    -- Process AD users and sync with DMS Users table
    -- Add your specific logic here
END
```

### 2. Custom User Profile Fields

#### Extend User Profile Class
```vb
' Add to Service.vb
Public Class ExtendedProfile
    Inherits profile
    
    Public DEPARTMENT As String
    Public TITLE As String
    Public MANAGER As String
    Public LOCATION As String
    Public EMPLOYEE_ID As String
    Public HIRE_DATE As String
    Public SECURITY_CLEARANCE As String
    Public CUSTOM_FIELD1 As String
    Public CUSTOM_FIELD2 As String
End Class
```

### 3. Role-Based Access Control

#### Define Custom Roles
```sql
-- Create role-based access control
CREATE TABLE User_Roles (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    user_id INT NOT NULL,
    role_name NVARCHAR(50) NOT NULL,
    granted_by INT,
    granted_date DATETIME DEFAULT GETDATE(),
    expires_date DATETIME NULL,
    FOREIGN KEY (user_id) REFERENCES Users(row_id)
);

-- Insert standard roles
INSERT INTO User_Roles (user_id, role_name) 
SELECT U.row_id, 'Document_Admin'
FROM Users U 
WHERE U.ext_user_id IN ('admin1', 'admin2');

INSERT INTO User_Roles (user_id, role_name) 
SELECT U.row_id, 'Content_Manager'
FROM Users U 
WHERE U.ext_user_id IN ('manager1', 'manager2');
```

## Document Classification System

### 1. Custom Document Types

#### Add Organization-Specific Types
```sql
-- Add your organization's document types
INSERT INTO Document_Types (extension) VALUES 
('.dwg'),      -- AutoCAD
('.zip'),      -- Archives
('.mp4'),      -- Videos
('.wav'),      -- Audio
('.csv'),      -- Data
('.json'),     -- Config
('.xml'),      -- XML
('.html'),     -- Web
('.css'),      -- Styles
('.js');       -- Scripts
```

### 2. Custom Categories and Keywords

#### Industry-Specific Categories
```sql
-- Manufacturing/Engineering
INSERT INTO Categories (name, public_flag) VALUES 
('Engineering Drawings', 'N'),
('Quality Control', 'Y'),
('Safety Procedures', 'Y'),
('Equipment Manuals', 'Y'),
('Maintenance Records', 'N'),
('Test Results', 'N'),
('Compliance Certificates', 'Y');

-- Healthcare
INSERT INTO Categories (name, public_flag) VALUES 
('Patient Records', 'N'),
('Medical Procedures', 'N'),
('Drug Information', 'N'),
('Insurance Forms', 'N'),
('HIPAA Compliance', 'N'),
('Clinical Trials', 'N'),
('Research Data', 'N');

-- Financial Services
INSERT INTO Categories (name, public_flag) VALUES 
('Account Statements', 'N'),
('Investment Reports', 'N'),
('Regulatory Filings', 'Y'),
('Audit Documents', 'N'),
('Risk Assessments', 'N'),
('Compliance Reports', 'Y'),
('Client Communications', 'N');
```

### 3. Custom Metadata Fields

#### Extend Document Keywords
```sql
-- Add custom metadata fields
ALTER TABLE Document_Keywords ADD 
    custom_field1 NVARCHAR(255),
    custom_field2 NVARCHAR(255),
    custom_field3 NVARCHAR(255),
    custom_date1 DATETIME,
    custom_date2 DATETIME;

-- Create custom keyword types
INSERT INTO Keywords (name) VALUES 
('Project Code'),
('Client ID'),
('Contract Number'),
('Version Number'),
('Approval Status'),
('Review Date'),
('Next Review Date'),
('Retention Period'),
('Disposal Date'),
('Confidentiality Level');
```

## Access Control Customization

### 1. Custom Access Rules

#### Create Access Control Matrix
```sql
CREATE TABLE Access_Control_Matrix (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    user_type NVARCHAR(50) NOT NULL,
    document_category NVARCHAR(50) NOT NULL,
    access_level NVARCHAR(20) NOT NULL, -- READ, WRITE, DELETE, ADMIN
    conditions NVARCHAR(MAX), -- JSON conditions
    created DATETIME DEFAULT GETDATE()
);

-- Define access rules
INSERT INTO Access_Control_Matrix (user_type, document_category, access_level, conditions) VALUES 
('Employee', 'General', 'READ', '{"time_restrictions": "business_hours"}'),
('Manager', 'General', 'WRITE', '{"approval_required": false}'),
('HR', 'HR Documents', 'ADMIN', '{"audit_required": true}'),
('Contractor', 'General', 'READ', '{"ip_restrictions": "office_only"}'),
('External', 'Public', 'READ', '{"session_timeout": 15}');
```

### 2. Time-Based Access Control

#### Implement Time Restrictions
```sql
CREATE TABLE Access_Time_Rules (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    user_id INT,
    group_id INT,
    document_id INT,
    start_time TIME,
    end_time TIME,
    days_of_week NVARCHAR(7), -- 'MTWTFSS'
    timezone NVARCHAR(50),
    FOREIGN KEY (user_id) REFERENCES Users(row_id),
    FOREIGN KEY (group_id) REFERENCES Groups(row_id),
    FOREIGN KEY (document_id) REFERENCES Documents(row_id)
);

-- Example: Restrict access to sensitive documents during business hours only
INSERT INTO Access_Time_Rules (group_id, start_time, end_time, days_of_week) 
SELECT G.row_id, '08:00', '17:00', 'MTWTFSS'
FROM Groups G 
WHERE G.name = 'DOMAIN\DMS-Users';
```

### 3. IP-Based Access Control

#### Implement IP Restrictions
```sql
CREATE TABLE Access_IP_Rules (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    user_id INT,
    group_id INT,
    ip_address NVARCHAR(45),
    ip_range_start NVARCHAR(45),
    ip_range_end NVARCHAR(45),
    allowed BIT DEFAULT 1,
    FOREIGN KEY (user_id) REFERENCES Users(row_id),
    FOREIGN KEY (group_id) REFERENCES Groups(row_id)
);

-- Example: Restrict contractors to office IPs only
INSERT INTO Access_IP_Rules (group_id, ip_range_start, ip_range_end, allowed)
SELECT G.row_id, '192.168.1.1', '192.168.1.254', 1
FROM Groups G 
WHERE G.name = 'DOMAIN\DMS-Contractors';
```

## External System Integration

### 1. Siebel CRM Integration

#### Customize Siebel Connection
```xml
<!-- Update web.config for your Siebel environment -->
<appSettings>
  <add key="siebel_server" value="YOUR-SIEBEL-SERVER"/>
  <add key="siebel_database" value="SIEBELDB"/>
  <add key="siebel_user" value="SIEBEL_USER"/>
  <add key="siebel_password" value="SIEBEL_PASSWORD"/>
  <add key="siebel_connection_timeout" value="30"/>
</appSettings>
```

#### Custom Siebel Query Procedures
```sql
CREATE PROCEDURE GetSiebelContactInfo
    @ContactId NVARCHAR(50)
AS
BEGIN
    -- Custom query for your Siebel schema
    SELECT 
        C.FST_NAME,
        C.LAST_NAME,
        C.EMAIL_ADDR,
        C.X_REGISTRATION_NUM,
        D.DOMAIN,
        D.CS_EMAIL,
        C.X_PART_ID,
        C.X_TRAINER_NUM,
        C.REG_AS_EMP_ID,
        C.DEPARTMENT,
        C.TITLE,
        C.MANAGER_ID
    FROM siebeldb.dbo.S_CONTACT C
    LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID = C.ROW_ID
    LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID = SC.SUB_ID
    LEFT OUTER JOIN siebeldb.dbo.CX_SUB_DOMAIN D ON D.DOMAIN = S.DOMAIN
    WHERE C.ROW_ID = @ContactId;
END
```

### 2. Email Integration

#### Configure SMTP Settings
```xml
<system.net>
  <mailSettings>
    <smtp deliveryMethod="Network">
      <network 
        host="YOUR-SMTP-SERVER" 
        port="587" 
        userName="DMS-Service@yourdomain.com"
        password="YOUR-EMAIL-PASSWORD"
        enableSsl="true"/>
    </smtp>
  </mailSettings>
</system.net>
```

#### Custom Email Templates
```vb
' Add to Service.vb
Private Function GetEmailTemplate(templateType As String) As String
    Select Case templateType
        Case "DocumentPublished"
            Return "A new document has been published for your review. " & _
                   "Document: {0} " & _
                   "Access the document at: {1}"
        Case "DocumentExpiring"
            Return "Document {0} will expire on {1}. " & _
                   "Please review and take necessary action."
        Case "AccessDenied"
            Return "Access to document {0} was denied. " & _
                   "Contact your administrator for assistance."
        Case Else
            Return "DMS Notification: {0}"
    End Select
End Function
```

### 3. Cloud Service Integration

#### Customize Certegrity Integration
```xml
<appSettings>
  <!-- Update URLs for your Certegrity environment -->
  <add key="processing.com.certegrity.cloudsvc.service" value="https://your-processing.certegrity.com/service.asmx"/>
  <add key="basic.com.certegrity.cloudsvc.service" value="https://your-basic.certegrity.com/service.asmx"/>
  <add key="local.hq.dms.service" value="https://your-dms.certegrity.com/service.asmx"/>
  
  <!-- Add authentication settings -->
  <add key="certegrity_username" value="YOUR_USERNAME"/>
  <add key="certegrity_password" value="YOUR_PASSWORD"/>
  <add key="certegrity_api_key" value="YOUR_API_KEY"/>
</appSettings>
```

## UI and Branding Customization

### 1. Custom Help Page

#### Create Custom Help Content
```html
<!-- help.aspx -->
<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="help.aspx.vb" Inherits="help" %>

<!DOCTYPE html>
<html>
<head>
    <title>DMS Services - Help</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #YOUR-BRAND-COLOR; color: white; padding: 10px; }
        .content { margin: 20px 0; }
        .method { margin: 15px 0; padding: 10px; border: 1px solid #ccc; }
        .method-name { font-weight: bold; color: #YOUR-BRAND-COLOR; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Your Organization DMS Services</h1>
        <p>Document Management System API Help</p>
    </div>
    
    <div class="content">
        <h2>Available Web Methods</h2>
        
        <div class="method">
            <div class="method-name">SaveDMSDoc</div>
            <p>Creates or updates a document in the system.</p>
            <p><strong>Parameters:</strong> DocId, ItemName, DFileName, Description, FileExt, etc.</p>
        </div>
        
        <div class="method">
            <div class="method-name">PublishDMSDoc</div>
            <p>Publishes a document to specific users with optional notifications.</p>
            <p><strong>Parameters:</strong> DocId, ContactId, NotifyFlg, ReqdFlag, Domain, Expiration</p>
        </div>
        
        <!-- Add more methods as needed -->
    </div>
</body>
</html>
```

### 2. Custom Error Pages

#### Create Custom Error Handling
```xml
<!-- Add to web.config -->
<system.web>
  <customErrors mode="On" defaultRedirect="~/Error.aspx">
    <error statusCode="404" redirect="~/NotFound.aspx"/>
    <error statusCode="500" redirect="~/ServerError.aspx"/>
  </customErrors>
</system.web>
```

### 3. Branding and Styling

#### Custom CSS for Help Pages
```css
/* Custom styles for your organization */
:root {
    --primary-color: #YOUR-BRAND-COLOR;
    --secondary-color: #YOUR-SECONDARY-COLOR;
    --text-color: #333;
    --background-color: #f5f5f5;
}

body {
    font-family: 'Your-Font', Arial, sans-serif;
    background-color: var(--background-color);
    color: var(--text-color);
}

.header {
    background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
    color: white;
    padding: 20px;
    border-radius: 5px;
    margin-bottom: 20px;
}

.method {
    background: white;
    border-left: 4px solid var(--primary-color);
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
```

## Performance Tuning

### 1. Database Optimization

#### Query Optimization
```sql
-- Create covering indexes for common queries
CREATE NONCLUSTERED INDEX IX_Documents_Covering 
ON Documents(deleted, created) 
INCLUDE (row_id, name, dfilename, description);

-- Optimize document access queries
CREATE NONCLUSTERED INDEX IX_Document_Access_Optimized 
ON Document_Associations(association_id, fkey, pr_flag, access_flag) 
INCLUDE (doc_id, reqd_flag, expiration);
```

#### Database Maintenance
```sql
-- Create maintenance procedures
CREATE PROCEDURE DMS_Maintenance
AS
BEGIN
    -- Update statistics
    EXEC sp_updatestats;
    
    -- Rebuild fragmented indexes
    DECLARE @sql NVARCHAR(MAX) = '';
    SELECT @sql = @sql + 'ALTER INDEX ' + name + ' ON ' + OBJECT_NAME(object_id) + ' REBUILD;' + CHAR(13)
    FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'LIMITED') ips
    INNER JOIN sys.indexes i ON ips.object_id = i.object_id AND ips.index_id = i.index_id
    WHERE ips.avg_fragmentation_in_percent > 30;
    
    EXEC sp_executesql @sql;
    
    -- Clean up old sessions
    DELETE FROM User_Sessions 
    WHERE created < DATEADD(day, -30, GETDATE());
    
    -- Archive old document versions
    -- Add your archiving logic here
END
```

### 2. Application Performance

#### Caching Configuration
```xml
<!-- Add to web.config -->
<system.web>
  <caching>
    <outputCacheSettings>
      <outputCacheProfiles>
        <add name="DMS_Cache" duration="300" varyByParam="*"/>
      </outputCacheProfiles>
    </outputCacheSettings>
  </caching>
</system.web>
```

#### Connection Pool Optimization
```xml
<connectionStrings>
  <add name="dms" 
       connectionString="server=YOUR-SERVER;database=DMS;Integrated Security=true;Min Pool Size=10;Max Pool Size=50;Connection Timeout=30;Command Timeout=60" 
       providerName="System.Data.SqlClient"/>
</connectionStrings>
```

### 3. File System Optimization

#### Implement File Compression
```vb
' Add to Service.vb
Private Function CompressFile(filePath As String) As String
    Try
        Dim compressedPath As String = filePath & ".gz"
        Using inputStream As New FileStream(filePath, FileMode.Open)
            Using outputStream As New FileStream(compressedPath, FileMode.Create)
                Using gzipStream As New GZipStream(outputStream, CompressionMode.Compress)
                    inputStream.CopyTo(gzipStream)
                End Using
            End Using
        End Using
        Return compressedPath
    Catch ex As Exception
        Return filePath ' Return original if compression fails
    End Try
End Function
```

## Security Hardening

### 1. Authentication Security

#### Implement Multi-Factor Authentication
```vb
' Add to Service.vb
Private Function ValidateMFA(userId As String, token As String) As Boolean
    ' Integrate with your MFA provider (Azure MFA, Google Authenticator, etc.)
    ' This is a placeholder - implement based on your MFA solution
    Return True
End Function
```

#### Session Security
```xml
<!-- Add to web.config -->
<system.web>
  <sessionState 
    mode="InProc" 
    timeout="30" 
    regenerateExpiredSessionId="true"
    cookieName="DMSSession"
    cookieSecure="true"
    cookieHttpOnly="true"/>
</system.web>
```

### 2. Data Encryption

#### Encrypt Sensitive Data
```vb
' Add to Service.vb
Private Function EncryptSensitiveData(data As String) As String
    ' Implement encryption for sensitive document metadata
    ' Use AES encryption or your organization's standard
    Return data ' Placeholder
End Function

Private Function DecryptSensitiveData(encryptedData As String) As String
    ' Implement decryption
    Return encryptedData ' Placeholder
End Function
```

### 3. Audit Logging

#### Implement Comprehensive Auditing
```sql
CREATE TABLE Audit_Log (
    row_id INT IDENTITY(1,1) PRIMARY KEY,
    user_id NVARCHAR(50),
    action_type NVARCHAR(50),
    table_name NVARCHAR(50),
    record_id NVARCHAR(50),
    old_values NVARCHAR(MAX),
    new_values NVARCHAR(MAX),
    ip_address NVARCHAR(45),
    user_agent NVARCHAR(500),
    timestamp DATETIME DEFAULT GETDATE()
);

-- Create audit triggers
CREATE TRIGGER TR_Documents_Audit
ON Documents
AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    -- Log document changes
    INSERT INTO Audit_Log (user_id, action_type, table_name, record_id, new_values)
    SELECT 
        SYSTEM_USER,
        CASE 
            WHEN EXISTS(SELECT * FROM inserted) AND EXISTS(SELECT * FROM deleted) THEN 'UPDATE'
            WHEN EXISTS(SELECT * FROM inserted) THEN 'INSERT'
            ELSE 'DELETE'
        END,
        'Documents',
        ISNULL(i.row_id, d.row_id),
        (SELECT * FROM inserted FOR JSON AUTO)
    FROM inserted i
    FULL OUTER JOIN deleted d ON i.row_id = d.row_id;
END
```

## Deployment Strategies

### 1. Development Environment

#### Local Development Setup
```batch
@echo off
echo Setting up DMS Development Environment...

REM Create local directories
mkdir C:\DMS-Dev
mkdir C:\DMS-Dev\work_dir
mkdir C:\DMS-Dev\logs

REM Set up local database
sqlcmd -S localhost -i DMS_Database_Script.sql

REM Configure local web.config
copy web.config.dev web.config

echo Development environment ready!
pause
```

#### Development Configuration
```xml
<!-- web.config.dev -->
<appSettings>
  <add key="basepath" value="C:\DMS-Dev\"/>
  <add key="SaveDMSDocDoc_debug" value="Y"/>
  <add key="SaveDMSDocAssoc_debug" value="Y"/>
  <add key="SaveDMSDocCat_debug" value="Y"/>
  <add key="SaveDMSDocKey_debug" value="Y"/>
  <add key="SaveDMSDocUser_debug" value="Y"/>
  <add key="UpdDMSDocCount_debug" value="Y"/>
  <add key="UpdDMSDoc_debug" value="Y"/>
  <add key="PublishDMSDoc_debug" value="Y"/>
  <add key="CheckDocAccess_debug" value="Y"/>
  <add key="UserLogin_debug" value="Y"/>
  <add key="UserLogout_debug" value="Y"/>
</appSettings>
```

### 2. Staging Environment

#### Staging Deployment Script
```powershell
# Deploy to staging
param(
    [string]$StagingServer = "staging-server",
    [string]$StagingPath = "C:\Inetpub\wwwroot\DMS-Staging"
)

Write-Host "Deploying to staging environment..."

# Stop application pool
Invoke-Command -ComputerName $StagingServer -ScriptBlock {
    Import-Module WebAdministration
    Stop-WebAppPool -Name "DMS-Staging"
}

# Copy files
Copy-Item -Path ".\*" -Destination "\\$StagingServer\$StagingPath" -Recurse -Force

# Update configuration
$configPath = "\\$StagingServer\$StagingPath\web.config"
$config = [xml](Get-Content $configPath)
$config.configuration.appSettings.add | Where-Object {$_.key -eq "basepath"} | ForEach-Object {$_.value = "C:\DMS-Staging\"}
$config.Save($configPath)

# Start application pool
Invoke-Command -ComputerName $StagingServer -ScriptBlock {
    Import-Module WebAdministration
    Start-WebAppPool -Name "DMS-Staging"
}

Write-Host "Staging deployment complete!"
```

### 3. Production Deployment

#### Production Deployment Checklist
```markdown
## Pre-Deployment Checklist
- [ ] Database backup completed
- [ ] Application backup completed
- [ ] Configuration files reviewed
- [ ] Security settings verified
- [ ] Performance tests passed
- [ ] User acceptance testing completed
- [ ] Rollback plan prepared

## Deployment Steps
1. Stop production application pool
2. Backup current application
3. Deploy new version
4. Update configuration
5. Run database migrations
6. Start application pool
7. Verify functionality
8. Monitor for issues

## Post-Deployment Verification
- [ ] All web methods responding
- [ ] Database connections working
- [ ] File upload/download working
- [ ] User authentication working
- [ ] External integrations working
- [ ] Logging functioning
- [ ] Performance within acceptable limits
```

## Maintenance and Updates

### 1. Regular Maintenance Tasks

#### Daily Maintenance
```sql
-- Daily maintenance script
CREATE PROCEDURE DMS_Daily_Maintenance
AS
BEGIN
    -- Clean up expired sessions
    DELETE FROM User_Sessions 
    WHERE created < DATEADD(hour, -24, GETDATE());
    
    -- Update document access statistics
    UPDATE Documents 
    SET last_accessed = GETDATE()
    WHERE row_id IN (
        SELECT DISTINCT doc_id 
        FROM Document_Associations 
        WHERE created > DATEADD(day, -1, GETDATE())
    );
    
    -- Log maintenance completion
    INSERT INTO Maintenance_Log (task_name, completion_time, status)
    VALUES ('Daily Maintenance', GETDATE(), 'Completed');
END
```

#### Weekly Maintenance
```sql
-- Weekly maintenance script
CREATE PROCEDURE DMS_Weekly_Maintenance
AS
BEGIN
    -- Update statistics
    EXEC sp_updatestats;
    
    -- Check for fragmented indexes
    DECLARE @fragmented_indexes INT;
    SELECT @fragmented_indexes = COUNT(*)
    FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'LIMITED') ips
    INNER JOIN sys.indexes i ON ips.object_id = i.object_id AND ips.index_id = i.index_id
    WHERE ips.avg_fragmentation_in_percent > 30;
    
    IF @fragmented_indexes > 0
    BEGIN
        -- Log fragmentation warning
        INSERT INTO Maintenance_Log (task_name, completion_time, status, notes)
        VALUES ('Index Fragmentation Check', GETDATE(), 'Warning', 
                'Found ' + CAST(@fragmented_indexes AS VARCHAR(10)) + ' fragmented indexes');
    END
    
    -- Archive old audit logs
    DELETE FROM Audit_Log 
    WHERE timestamp < DATEADD(month, -6, GETDATE());
END
```

### 2. Monitoring and Alerting

#### Performance Monitoring
```sql
-- Create monitoring views
CREATE VIEW DMS_Performance_Metrics AS
SELECT 
    'Active Sessions' AS metric_name,
    COUNT(*) AS metric_value,
    GETDATE() AS timestamp
FROM User_Sessions
WHERE created > DATEADD(hour, -1, GETDATE())

UNION ALL

SELECT 
    'Documents Created Today' AS metric_name,
    COUNT(*) AS metric_value,
    GETDATE() AS timestamp
FROM Documents
WHERE created > DATEADD(day, -1, GETDATE())

UNION ALL

SELECT 
    'Database Size (MB)' AS metric_name,
    CAST(SUM(size) * 8.0 / 1024 AS INT) AS metric_value,
    GETDATE() AS timestamp
FROM sys.database_files;
```

#### Alert Configuration
```vb
' Add to Service.vb
Private Sub CheckSystemHealth()
    Try
        ' Check database connectivity
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("dms").ConnectionString)
            conn.Open()
            ' Database is accessible
        End Using
        
        ' Check file system space
        Dim drive As New DriveInfo("C:")
        If drive.AvailableFreeSpace < 1024 * 1024 * 1024 Then ' Less than 1GB
            SendAlert("Low disk space on C: drive")
        End If
        
        ' Check memory usage
        Dim process As Process = Process.GetCurrentProcess()
        If process.WorkingSet64 > 1024 * 1024 * 1024 Then ' More than 1GB
            SendAlert("High memory usage: " & (process.WorkingSet64 / 1024 / 1024) & " MB")
        End If
        
    Catch ex As Exception
        SendAlert("System health check failed: " & ex.Message)
    End Try
End Sub

Private Sub SendAlert(message As String)
    ' Implement your alerting mechanism (email, SMS, etc.)
    ' This is a placeholder
End Sub
```

### 3. Backup and Recovery

#### Database Backup Strategy
```sql
-- Create backup procedures
CREATE PROCEDURE DMS_Backup_Database
AS
BEGIN
    DECLARE @backupPath NVARCHAR(500);
    DECLARE @backupName NVARCHAR(500);
    
    SET @backupPath = 'C:\Backups\DMS\';
    SET @backupName = @backupPath + 'DMS_' + FORMAT(GETDATE(), 'yyyyMMdd_HHmmss') + '.bak';
    
    BACKUP DATABASE DMS 
    TO DISK = @backupName
    WITH FORMAT, INIT, COMPRESSION;
    
    -- Clean up old backups (keep 30 days)
    DECLARE @deleteCmd NVARCHAR(1000);
    SET @deleteCmd = 'DEL "' + @backupPath + 'DMS_*.bak" /Q';
    
    -- This would need to be executed via xp_cmdshell or SQL Agent job
END
```

#### Application Backup
```powershell
# Application backup script
param(
    [string]$BackupPath = "C:\Backups\DMS-App\",
    [string]$AppPath = "C:\Inetpub\wwwroot\DMS"
)

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFile = "$BackupPath\DMS-App_$timestamp.zip"

# Create backup directory if it doesn't exist
if (!(Test-Path $BackupPath)) {
    New-Item -ItemType Directory -Path $BackupPath -Force
}

# Create zip backup
Compress-Archive -Path $AppPath -DestinationPath $backupFile -Force

# Clean up old backups (keep 30 days)
Get-ChildItem -Path $BackupPath -Name "DMS-App_*.zip" | 
Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-30) } | 
Remove-Item -Force

Write-Host "Application backup created: $backupFile"
```

---

## Conclusion

This customization guide provides comprehensive instructions for implementing DMS Services in your organization. The key to successful implementation is:

1. **Start with a pilot environment** to test configurations
2. **Customize gradually** - implement core functionality first, then add customizations
3. **Document all changes** for future maintenance
4. **Test thoroughly** before production deployment
5. **Monitor performance** and adjust as needed
6. **Plan for maintenance** and regular updates

Remember to adapt these guidelines to your specific organizational requirements, security policies, and technical environment. Regular review and updates of this customization guide will ensure it remains relevant as your DMS implementation evolves.
