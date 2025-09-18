# SCORM Learning Support Services - Customization Guide

## Overview

This guide provides step-by-step instructions for implementing and customizing the SCORM Learning Support Services in your environment. The system is designed to be flexible and adaptable to various organizational needs while maintaining security and performance standards.

## Prerequisites

### System Requirements
- Windows Server 2012 R2 or later
- IIS 7.0 or later with ASP.NET 4.5.2
- SQL Server 2012 or later
- .NET Framework 4.5.2
- MinIO server (for object storage)

### Software Dependencies
- Visual Studio 2015 or later (for development)
- SQL Server Management Studio
- IIS Manager
- MinIO client tools

## Installation Steps

### 1. Database Setup

#### Step 1.1: Create Databases
Execute the provided DDL script (`SCORM_Services_DDL.sql`) to create all required databases:

```sql
-- Run the complete DDL script
-- This creates: siebeldb, elearning, DMS, scormmedia, reports
```

#### Step 1.2: Configure Database Users
Create dedicated database users for each service:

```sql
-- Create service accounts
CREATE LOGIN [scorm_service] WITH PASSWORD = 'YourSecurePassword123!';
CREATE USER [scorm_service] FOR LOGIN [scorm_service];

-- Grant permissions to each database
USE siebeldb;
ALTER ROLE db_datareader ADD MEMBER [scorm_service];
ALTER ROLE db_datawriter ADD MEMBER [scorm_service];

USE elearning;
ALTER ROLE db_datareader ADD MEMBER [scorm_service];
ALTER ROLE db_datawriter ADD MEMBER [scorm_service];

USE DMS;
ALTER ROLE db_datareader ADD MEMBER [scorm_service];
ALTER ROLE db_datawriter ADD MEMBER [scorm_service];

USE scormmedia;
ALTER ROLE db_datareader ADD MEMBER [scorm_service];
ALTER ROLE db_datawriter ADD MEMBER [scorm_service];

USE reports;
ALTER ROLE db_datareader ADD MEMBER [scorm_service];
ALTER ROLE db_datawriter ADD MEMBER [scorm_service];
```

### 2. Web Server Configuration

#### Step 2.1: IIS Setup
1. Install IIS with ASP.NET 4.5.2 support
2. Create a new application pool for SCORM services
3. Set the application pool to use .NET Framework 4.5.2
4. Configure the application pool identity

#### Step 2.2: Deploy Application
1. Copy the application files to your web server
2. Create a virtual directory in IIS pointing to the application folder
3. Set the virtual directory as an application
4. Assign the application to the SCORM application pool

### 3. Configuration

#### Step 3.1: Update web.config

Replace the placeholder values in `web.config` with your actual configuration:

```xml
<appSettings>
    <!-- Database Configuration -->
    <add key="dbuser" value="scorm_service"/>
    <add key="dbpass" value="YourSecurePassword123!"/>
    
    <!-- Service URLs -->
    <add key="processing.gettips.siebel.service" value="http://your-domain.com/Processing/service.asmx"/>
    <add key="com.certegrity.cloudsvc.basic.service" value="http://your-domain.com/basic/service.asmx"/>
    <add key="com.certegrity.scorm.svc.service" value="http://your-domain.com/svc/service.asmx"/>
    
    <!-- MinIO Configuration -->
    <add key="minio-key" value="your-minio-access-key"/>
    <add key="minio-secret" value="your-minio-secret-key"/>
    <add key="minio-region" value="us-east"/>
    <add key="minio-bucket" value="your-bucket-name"/>
    
    <!-- Debug Settings -->
    <add key="GetConfiguration_debug" value="N"/>
    <add key="GetLessonData_debug" value="N"/>
    <!-- Set other debug flags as needed -->
</appSettings>

<connectionStrings>
    <add name="siebeldb" connectionString="server=YOUR_SERVER\YOUR_INSTANCE;uid=scorm_service;pwd=YourSecurePassword123!;database=siebeldb;Min Pool Size=3;Max Pool Size=5" providerName="System.Data.SqlClient"/>
    <add name="email" connectionString="server=YOUR_SERVER\YOUR_INSTANCE;uid=scorm_service;pwd=YourSecurePassword123!;database=scanner;Min Pool Size=3;Max Pool Size=5" providerName="System.Data.SqlClient"/>
    <add name="hcidb" connectionString="server=YOUR_SERVER\YOUR_INSTANCE;uid=scorm_service;pwd=YourSecurePassword123!;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;database=" providerName="System.Data.SqlClient"/>
    <add name="hcidbro" connectionString="server=YOUR_SERVER\YOUR_INSTANCE;uid=scorm_service;pwd=YourSecurePassword123!;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;ApplicationIntent=ReadOnly;database=siebeldb" providerName="System.Data.SqlClient"/>
    <add name="dms" connectionString="server=YOUR_SERVER\YOUR_INSTANCE;uid=scorm_service;pwd=YourSecurePassword123!;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;ApplicationIntent=ReadOnly;database=DMS" providerName="System.Data.SqlClient"/>
</connectionStrings>
```

#### Step 3.2: Configure Logging
Update the log4net configuration:

```xml
<log4net>
    <appender name="RemoteSyslogAppender" type="log4net.Appender.RemoteSyslogAppender">
        <remoteAddress value="YOUR_LOG_SERVER_IP"/>
        <!-- Configure other logging settings -->
    </appender>
</log4net>
```

### 4. MinIO Setup

#### Step 4.1: Install MinIO
1. Download and install MinIO server
2. Configure MinIO with your desired storage backend
3. Create the required bucket for media storage

#### Step 4.2: Configure MinIO Access
1. Create access keys for the SCORM service
2. Configure bucket policies for read/write access
3. Test connectivity from the web server

### 5. Security Configuration

#### Step 5.1: SSL/TLS Setup
1. Install SSL certificates on the web server
2. Configure HTTPS bindings in IIS
3. Update service URLs to use HTTPS

#### Step 5.2: Firewall Configuration
Configure firewall rules to allow:
- HTTP/HTTPS traffic to the web server
- Database connections from web server to SQL Server
- MinIO connections from web server

#### Step 5.3: Application Security
1. Configure Windows Authentication if required
2. Set appropriate file permissions on the application folder
3. Configure IIS security settings

### 6. Testing and Validation

#### Step 6.1: Database Connectivity Test
Create a simple test script to verify database connections:

```vb
' Test database connectivity
Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString)
Try
    conn.Open()
    ' Connection successful
Catch ex As Exception
    ' Handle connection error
Finally
    conn.Close()
End Try
```

#### Step 6.2: Service Endpoint Testing
Test each web service endpoint:
1. Access the WSDL: `http://your-server/Service.asmx?WSDL`
2. Test individual methods using SOAP clients
3. Verify response formats and error handling

#### Step 6.3: Media Access Testing
1. Upload test media files to MinIO
2. Test media retrieval through the service
3. Verify access control and caching

### 7. Performance Optimization

#### Step 7.1: Database Optimization
1. Create appropriate indexes (included in DDL script)
2. Configure database maintenance plans
3. Monitor query performance

#### Step 7.2: Web Server Optimization
1. Configure application pool settings
2. Enable output caching where appropriate
3. Configure compression

#### Step 7.3: Caching Configuration
1. Configure database query caching
2. Set up media asset caching
3. Monitor cache hit rates

### 8. Monitoring and Maintenance

#### Step 8.1: Logging Setup
1. Configure log rotation
2. Set up log monitoring
3. Create alerting for critical errors

#### Step 8.2: Performance Monitoring
1. Set up performance counters
2. Monitor database performance
3. Track web service response times

#### Step 8.3: Backup Strategy
1. Configure database backups
2. Backup MinIO data
3. Document recovery procedures

## Customization Options

### 1. Adding New Web Methods

To add new functionality:

1. Add the method to `Service.vb`:

```vb
<WebMethod(Description:="Your method description")>
Public Function YourNewMethod(ByVal param1 As String, ByVal param2 As String) As XmlDocument
    ' Implementation here
End Function
```

2. Update the database schema if needed
3. Test the new method
4. Update documentation

### 2. Customizing Database Schema

To modify the database structure:

1. Create migration scripts for schema changes
2. Update the DDL script for new installations
3. Test changes in a development environment
4. Update application code to use new schema

### 3. Adding New Media Types

To support additional media types:

1. Update the media handling code
2. Configure MinIO bucket policies
3. Update content type mappings
4. Test with various file formats

### 4. Customizing Authentication

To implement custom authentication:

1. Modify the authentication logic in the service methods
2. Update user management tables
3. Configure session handling
4. Test authentication flows

## Troubleshooting

### Common Issues

#### Database Connection Errors
- Verify connection strings
- Check SQL Server service status
- Validate user permissions
- Test network connectivity

#### Media Access Issues
- Verify MinIO server status
- Check access keys and permissions
- Validate bucket configuration
- Test network connectivity

#### Performance Issues
- Monitor database performance
- Check web server resources
- Review caching configuration
- Analyze slow queries

#### Authentication Problems
- Verify user credentials
- Check session configuration
- Validate token handling
- Review security settings

### Log Analysis

#### Application Logs
- Check IIS logs for HTTP errors
- Review application event logs
- Monitor custom log files
- Analyze performance counters

#### Database Logs
- Review SQL Server error logs
- Monitor deadlock information
- Check connection pool status
- Analyze query execution plans

## Support and Maintenance

### Regular Maintenance Tasks

1. **Weekly**
   - Review error logs
   - Check disk space
   - Monitor performance metrics

2. **Monthly**
   - Update security patches
   - Review and rotate logs
   - Performance optimization review

3. **Quarterly**
   - Database maintenance
   - Security audit
   - Backup testing

### Documentation Updates

Keep documentation current by:
1. Updating configuration changes
2. Documenting custom modifications
3. Recording troubleshooting solutions
4. Maintaining deployment procedures

## Contact Information

For technical support and questions:
- System Administrator: [Your Contact Information]
- Database Administrator: [Your Contact Information]
- Development Team: [Your Contact Information]

---

*This customization guide should be updated as your implementation evolves and new requirements are identified.*
