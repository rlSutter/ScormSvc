# SCORM Learning Support Services

## Overview

The SCORM Learning Support Services is a comprehensive web service platform built on ASP.NET that provides backend infrastructure for SCORM-compliant Learning Management Systems (LMS). The system manages the complete e-learning lifecycle through a multi-database architecture supporting user management, course delivery, assessments, and media asset management. It features Knowledge-Based Authentication (KBA) for secure user verification, SCORM player data tracking for course progress monitoring, comprehensive assessment capabilities with question banks and scoring, and integrated media management with MinIO object storage for scalable content delivery. The platform includes over 20 web service methods covering everything from course configuration and lesson data retrieval to secure media access and exam management, all designed to support educational institutions and training organizations in delivering interactive, trackable, and compliant online learning experiences with robust security, performance optimization, and comprehensive logging capabilities.

## Architecture

The system is built using ASP.NET Web Services (ASMX) with VB.NET and follows a multi-database architecture to support different aspects of the learning platform:

### Core Components

1. **Main Web Service** (`Service.asmx`) - Primary SOAP web service providing SCORM functionality
2. **KBA Handler** (`KBALookup.ashx`) - HTTP handler for Knowledge-Based Authentication
3. **Database Layer** - Multi-database architecture supporting different functional areas
4. **Media Management** - Integration with MinIO for object storage
5. **Logging System** - Comprehensive logging using log4net

### Database Architecture

The system uses five separate SQL Server databases:

- **siebeldb** - Main application database for user management, courses, and registrations
- **elearning** - E-learning specific data including SCORM player data and assessments
- **DMS** - Document Management System for course materials and media
- **scormmedia** - Media asset management and access control
- **reports** - Reporting and analytics data

## Key Features

### 1. SCORM Compliance
- Full SCORM 1.2 and 2004 support
- SCORM player data management
- Course progress tracking
- Completion status monitoring

### 2. Knowledge-Based Authentication (KBA)
- Secure user identity verification
- Question and answer management
- Jurisdiction-specific KBA requirements
- Encrypted answer storage

### 3. Media Management
- Course asset delivery
- Secure media access control
- MinIO integration for scalable storage
- Caching mechanisms for performance

### 4. Assessment System
- Online testing capabilities
- Question bank management
- Score tracking and reporting
- Time-limited assessments

### 5. User Management
- Contact and registration management
- Session tracking
- Access control and permissions
- Multi-domain support

## Web Service Methods

### Core SCORM Methods

#### `GetConfiguration`
Retrieves course configuration data including:
- Course settings and parameters
- Media destinations
- Service URLs
- KBA requirements

#### `GetLessonData`
Provides lesson content and structure:
- Course elements and navigation
- Media asset references
- Progress tracking data

#### `GetScenarioData`
Delivers scenario-based learning content:
- Interactive scenarios
- Decision trees
- Outcome tracking

#### `QuizLibrary`
Manages question banks and assessments:
- Question retrieval
- Answer validation
- Score calculation

### Media Management Methods

#### `GetMedia`
Retrieves course media assets:
- Images, videos, and documents
- Access control validation
- Caching support

#### `GetMediaSecure`
Secure media access with user authentication:
- User-specific media delivery
- Access logging
- Permission validation

#### `GetHMedia`
Help media and documentation:
- Course help files
- User guides
- Support materials

#### `GetDImage`
Document images from DMS:
- Document thumbnails
- Preview images
- Access control

### Assessment Methods

#### `GetExamXML`
Retrieves exam configuration:
- Question structure
- Time limits
- Passing scores

#### `GetExamHTML5`
HTML5-compatible exam delivery:
- Modern web interface
- Responsive design
- Real-time validation

#### `StoreExamData`
Saves assessment results:
- Answer storage
- Score calculation
- Progress tracking

### KBA Methods

#### `KBALookup`
Retrieves KBA questions:
- Question selection
- Answer validation
- Jurisdiction-specific requirements

#### `KBAQLookup`
Question lookup for assessments:
- Question bank access
- Randomization
- Filtering

#### `KBAALookup`
Answer validation:
- Answer verification
- Security checks
- Audit logging

#### `KBAReset`
Resets KBA attempts:
- User retry management
- Security controls

### User Management Methods

#### `GetLocalRegs`
Local registration management:
- User session data
- Access permissions
- Status tracking

#### `SaveCourseNote`
User note management:
- Personal notes
- Screen-specific notes
- Timestamp tracking

#### `GetCourseNote`
Retrieves user notes:
- Note retrieval
- Organization
- Search capabilities

### Utility Methods

#### `CourseGlossary`
Course terminology and definitions:
- Term lookup
- Multi-language support
- Context-specific definitions

#### `SaveFeedback`
User feedback collection:
- Bug reports
- Feature requests
- User experience feedback

#### `GetStateBACLevels` / `GetAllBACLevels`
Blood Alcohol Content level management:
- Jurisdiction-specific levels
- Regulatory compliance
- Educational content

#### `GetIDExercise`
Identity verification exercises:
- Training scenarios
- Skill assessment
- Progress tracking

## Security Features

### Authentication
- User session management
- Token-based authentication
- Secure credential handling

### Data Protection
- Encrypted sensitive data
- SQL injection prevention
- Input validation and sanitization

### Access Control
- Role-based permissions
- Domain-specific access
- Media access restrictions

### Audit Logging
- Comprehensive activity logging
- Performance monitoring
- Error tracking and reporting

## Performance Features

### Caching
- Database query caching
- Media asset caching
- Session data caching

### Optimization
- Connection pooling
- Asynchronous operations
- Efficient data retrieval

### Monitoring
- Performance metrics
- Error tracking
- Usage analytics

## Integration Points

### External Services
- MinIO object storage
- Email services
- Logging services
- Cloud-based processing

### Database Integration
- Multi-database architecture
- Stored procedure support
- Transaction management

### API Support
- SOAP web services
- RESTful endpoints
- JSON data exchange

## Configuration

The system is configured through `web.config` with settings for:

- Database connections
- Service endpoints
- Media storage
- Logging configuration
- Security settings

## Dependencies

### Required Components
- .NET Framework 4.5.2
- SQL Server 2012 or later
- IIS 7.0 or later
- MinIO (for object storage)

### Third-Party Libraries
- log4net (logging)
- Newtonsoft.Json (JSON processing)
- Amazon S3 SDK (MinIO integration)

## Deployment

The system is designed for deployment in:
- Windows Server environments
- IIS web server
- SQL Server database
- Load-balanced configurations

## Monitoring and Maintenance

### Logging
- Application event logging
- Performance monitoring
- Error tracking
- User activity auditing

### Maintenance Tasks
- Database optimization
- Cache management
- Log rotation
- Security updates

## Support and Documentation

For technical support and additional documentation, refer to:
- API documentation
- Database schema documentation
- Deployment guides
- Troubleshooting guides

## Version Information

- **Version**: 1.0
- **Framework**: .NET Framework 4.5.2
- **Database**: SQL Server 2012+
- **Web Server**: IIS 7.0+

## License

This software is proprietary and confidential. All rights reserved.

---

*For implementation and customization instructions, see CUSTOMIZING.md*
