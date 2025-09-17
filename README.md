# DMS Services - Document Management System

## Overview

DMS Services is a web service application that provides a front-end interface to a comprehensive document management system. Built using ASP.NET Web Services (ASMX) with VB.NET, it integrates with multiple external systems including Siebel CRM and various cloud services to manage document storage, access control, and user authentication.

## Architecture

### Technology Stack
- **Framework**: ASP.NET 4.0 Web Services (ASMX)
- **Language**: VB.NET
- **Database**: SQL Server
- **Authentication**: Windows Authentication
- **Logging**: log4net
- **External Integrations**: 
  - Siebel CRM (siebeldb)
  - Certegrity Cloud Services
  - Email Services

### Key Components

#### 1. Web Service Interface (`Service.asmx`)
- Main entry point for all DMS operations
- Provides SOAP-based web service endpoints
- Handles document management, user authentication, and access control

#### 2. Service Logic (`App_Code/Service.vb`)
- Core business logic implementation
- Document CRUD operations
- User session management
- Integration with external systems
- Comprehensive logging and debugging capabilities

#### 3. Database Layer
- **Primary Database**: DMS (Document Management System)
- **Integration Database**: siebeldb (Siebel CRM)
- **Email Database**: scanner
- Multiple connection strings for different database operations

## Database Schema

### Core Tables

#### Documents Management
- **Documents**: Main document metadata
- **Document_Versions**: Document file storage and versioning
- **Document_Types**: Supported file extensions

#### Classification System
- **Categories**: Document categorization
- **Keywords**: Document tagging and metadata
- **Associations**: Document association types

#### User Management
- **Users**: User accounts and external system integration
- **Groups**: User groups (Domain, Subscription, System)
- **User_Group_Access**: Access control matrix
- **User_Sessions**: Active user sessions

#### Document Relationships
- **Document_Categories**: Document-to-category mapping
- **Document_Keywords**: Document-to-keyword mapping
- **Document_Associations**: Document access associations
- **Document_Users**: User-specific document access
- **Category_Keywords**: Category-to-keyword relationships

### Key Features

#### Document Management
- **Document Storage**: Binary file storage with versioning
- **Metadata Management**: Comprehensive document classification
- **Access Control**: Multi-level permission system
- **Expiration Handling**: Time-based document access control

#### User Authentication & Authorization
- **Multi-System Integration**: Siebel CRM user integration
- **Role-Based Access**: Subscription, Domain, and System-level permissions
- **Session Management**: Secure session handling with machine tracking
- **External User Support**: Partner, Trainer, and Employee access types

#### Integration Capabilities
- **Siebel CRM**: User data synchronization
- **Email Services**: Automated notifications
- **Cloud Services**: External document processing
- **Reporting**: Document access and usage tracking

## Installation & Setup

### Prerequisites
- Windows Server with IIS
- SQL Server 2012 or later
- .NET Framework 4.0
- Visual Studio 2010 or later (for development)

### Database Setup

1. **Run the Database Script**
   ```sql
   -- Execute the DMS_Database_Script.sql file
   -- This creates the complete database schema with:
   -- - 15 core tables
   -- - 25+ performance indexes
   -- - Sample data
   -- - Database user with appropriate permissions
   ```

2. **Configure Connection Strings**
   Update `web.config` with your database server details:
   ```xml
   <connectionStrings>
     <add name="dms" connectionString="server=YOUR_SERVER;uid=DMS;pwd=YOUR_PASSWORD;database=DMS" />
     <add name="siebeldb" connectionString="server=YOUR_SERVER;uid=sa;pwd=YOUR_PASSWORD;database=siebeldb" />
     <!-- Additional connection strings as needed -->
   </connectionStrings>
   ```

### Application Configuration

1. **Web.config Settings**
   - Update `basepath` for document storage location
   - Configure external service URLs
   - Set debug flags for development/testing
   - Configure logging settings

2. **File System Setup**
   - Create document storage directories:
     - `work_dir/DOC/`
     - `work_dir/jpg/`
     - `work_dir/pdf/`
     - `work_dir/XML/`

3. **IIS Configuration**
   - Deploy to IIS with .NET 4.0 application pool
   - Enable Windows Authentication
   - Configure appropriate permissions for file system access

## API Reference

### Core Web Methods

#### Document Management
- **SaveDMSDoc**: Create/update documents with metadata
- **PublishDMSDoc**: Publish documents to specific users
- **UpdDMSDoc**: Update document information
- **UpdDMSDocCount**: Update document access counts

#### Document Classification
- **SaveDMSDocCat**: Associate documents with categories
- **SaveDMSDocKey**: Add keywords to documents
- **SaveDMSDocAssoc**: Create document associations

#### User Management
- **UserLogin**: Authenticate users and create sessions
- **UserLogout**: Terminate user sessions
- **CheckDocAccess**: Verify document access permissions

#### Document Retrieval
- **GetDMSDoc**: Retrieve document content
- **GetDMSDocList**: List accessible documents for users

### Request/Response Format
All methods use SOAP XML format with comprehensive error handling and debug logging capabilities.

## Security Features

### Access Control
- **Multi-Level Permissions**: User, Group, Domain, and System-level access
- **Document-Level Security**: Individual document access control
- **Session Management**: Secure session tracking with machine identification
- **Expiration Handling**: Time-based access control

### Integration Security
- **Windows Authentication**: Integrated with Active Directory
- **Database Security**: Dedicated database users with minimal required permissions
- **External System Integration**: Secure communication with Siebel and cloud services

## Logging & Monitoring

### Log4net Configuration
- **Multiple Appenders**: File-based and remote syslog logging
- **Debug Logging**: Comprehensive debug information for troubleshooting
- **Performance Monitoring**: Document access and processing metrics
- **Error Tracking**: Detailed error logging with context information

### Debug Modes
- **Y**: Enable debug logging
- **N**: Disable debug logging  
- **T**: Test mode (no database writes)

## External Dependencies

### Required Services
- **Siebel CRM**: User management and authentication
- **Email Services**: Notification system
- **Certegrity Cloud Services**: Document processing and validation
- **File System**: Document storage and retrieval

### Optional Integrations
- **Reporting Services**: Document usage analytics
- **Backup Services**: Document versioning and archival
- **Monitoring Systems**: Application performance monitoring

## Troubleshooting

### Common Issues

1. **Database Connection Errors**
   - Verify connection strings in web.config
   - Check database server accessibility
   - Confirm user permissions

2. **File Access Issues**
   - Verify IIS application pool identity permissions
   - Check file system directory permissions
   - Ensure basepath configuration is correct

3. **Authentication Problems**
   - Verify Windows Authentication is enabled
   - Check Siebel CRM connectivity
   - Confirm user exists in external systems

### Debug Mode
Enable debug mode by setting debug parameters to "Y" in web method calls for detailed logging and troubleshooting information.

## Development

### Code Structure
- **Service.vb**: Main service implementation (8,000+ lines)
- **web.config**: Configuration and connection strings
- **App_WebReferences**: External service references
- **Bin/**: Compiled assemblies and dependencies

### Key Classes
- **Service**: Main web service class
- **profile**: User profile data structure
- **enumObjectType**: Data type enumeration

### Best Practices
- Use debug mode during development
- Implement proper error handling
- Follow logging guidelines
- Test with various user types and permissions

## Support

For technical support and questions:
- Review debug logs for detailed error information
- Check database connectivity and permissions
- Verify external service integrations
- Consult system administrators for infrastructure issues

## Version Information
- **Framework**: .NET 4.0
- **Database**: SQL Server compatible
- **IIS**: 7.0 or later
- **Dependencies**: log4net 2.0.3

---

*This document management system provides comprehensive document storage, access control, and user management capabilities with extensive integration support for enterprise environments.*
