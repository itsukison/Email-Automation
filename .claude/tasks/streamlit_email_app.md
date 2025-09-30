# Streamlit Email Automation Frontend - Implementation Plan

## Project Overview
Create a user-friendly Streamlit frontend for the existing email automation system (AutoEmail.ipynb) that allows users to easily send bulk emails to companies from Excel files.

## Current System Analysis
The existing Jupyter notebook (`AutoEmail.ipynb`) provides:
- Excel file reading with company names and email addresses
- SMTP email sending with personalized content
- BCC functionality for internal recipients
- Success/error logging
- Hardcoded email templates for TikTok Shop business outreach

## Requirements
### Core Functionality
1. **File Upload**: Allow users to upload Excel files with company data
2. **Email Configuration**: Editable sender email, SMTP password, and server settings
3. **Content Management**: Editable email templates with dynamic company name insertion
4. **Bulk Email Sending**: Send personalized emails to all recipients
5. **Progress Tracking**: Real-time progress indication during email sending
6. **Results Reporting**: Success/error status for each email sent

### User Experience
- Intuitive step-by-step workflow
- Input validation and error handling
- Data preview before sending
- Confirmation dialogs for critical actions
- Clear success/error messaging

## Technical Architecture

### File Structure
```
/Users/bytedance/Desktop/code/email_Itsuki/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # Setup and usage instructions
└── .claude/
    └── tasks/
        └── streamlit_email_app.md  # This plan document
```

### Dependencies
- streamlit: Web application framework
- pandas: Excel file processing
- smtplib: Email sending (built-in)
- email: Email formatting (built-in)
- openpyxl: Excel file reading support

### Application Components

#### 1. Main Interface (`app.py`)
**Header Section**
- Application title and description
- Usage instructions

**File Upload Section**
- Excel file uploader
- File validation (format, required columns)
- Data preview table

**Email Configuration Section**
- Sender email input
- SMTP password input (masked)
- SMTP server and port settings
- BCC recipients configuration

**Email Template Section**
- Editable email content with company name placeholder
- Template preview functionality
- Default TikTok Shop template

**Sending Section**
- Send confirmation dialog
- Progress bar with current status
- Real-time sending updates
- Results summary

#### 2. Helper Functions
```python
def validate_email(email: str) -> bool
def process_excel_file(uploaded_file) -> pd.DataFrame
def generate_email_body(template: str, company_name: str) -> str
def send_emails(df, config, template) -> List[Dict]
def display_results(results: List[Dict]) -> None
```

#### 3. Session State Management
- Uploaded file data persistence
- Form input preservation
- Sending progress tracking
- Results storage

## Default Configuration
Based on the existing notebook:
- **SMTP Server**: smtp.larkoffice.com
- **SMTP Port**: 465 (SSL)
- **Default BCC**: michinari.nakai@bytedance.com, kentaro.koi@bytedance.com
- **Email Template**: TikTok Shop business outreach content

## Security Considerations
- Password input masking
- Security warnings for credential handling
- No credential storage in session state
- Input sanitization for email addresses

## Error Handling
- Invalid Excel file formats
- Missing required columns (entity name, email)
- SMTP authentication failures
- Network connectivity issues
- Invalid email address formats
- Timeout handling for email sending

## Implementation Steps

### Phase 1: Core Setup
1. Create project structure
2. Set up requirements.txt
3. Implement basic Streamlit layout
4. Add file upload functionality

### Phase 2: Email Configuration
1. Create email configuration form
2. Add input validation
3. Implement SMTP connection testing
4. Add template editor

### Phase 3: Email Processing
1. Port email sending logic from notebook
2. Add progress tracking
3. Implement error handling
4. Create results display

### Phase 4: User Experience
1. Add confirmation dialogs
2. Improve error messages
3. Add data preview functionality
4. Create documentation

### Phase 5: Testing & Refinement
1. Test with sample data
2. Verify email sending functionality
3. Handle edge cases
4. Performance optimization

## Success Criteria
- [ ] Users can upload Excel files with company data
- [ ] Email configuration is fully editable
- [ ] Email templates can be customized
- [ ] Bulk emails send successfully with progress tracking
- [ ] Clear error handling and user feedback
- [ ] Intuitive user interface requiring minimal training
- [ ] Maintains all functionality from original notebook

## Risk Mitigation
- **Email Delivery**: Test with small batches first
- **SMTP Limits**: Add rate limiting and retry logic
- **File Processing**: Validate Excel format and structure
- **User Errors**: Comprehensive input validation
- **Security**: Clear warnings about credential handling

## Future Enhancements (Post-MVP)
- Email template library
- Scheduled sending
- Email tracking and analytics
- Multiple file format support
- Advanced filtering options
- Export functionality for results