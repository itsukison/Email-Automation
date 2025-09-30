# Email Automation Tool

A user-friendly Streamlit application for sending personalized bulk emails to companies from Excel files.

## Features

- 📁 **Excel File Upload**: Upload company data with email addresses
- ⚙️ **Configurable SMTP Settings**: Support for various email providers
- ✉️ **Email Template Editor**: Customize email content with company name placeholders
- 📧 **CC/BCC Support**: Add carbon copy and blind carbon copy recipients
- 📊 **Progress Tracking**: Real-time progress during email sending
- 📈 **Results Dashboard**: Detailed success/failure reporting
- 💾 **Export Results**: Download sending results as CSV

## Installation

1. **Clone or download this repository**
2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Start the application:**
   ```bash
   streamlit run app.py
   ```

2. **Configure SMTP Settings** (in sidebar):
   - Enter your email address
   - Enter your email password or app-specific password
   - Configure SMTP server and port
   - Add CC/BCC recipients if needed

3. **Upload Excel File:**
   - File must contain columns: `entity name` and `email`
   - Supported formats: .xlsx, .xls

4. **Customize Email Template:**
   - Edit the email content in the template editor
   - Use `{company_name}` as placeholder for personalization
   - Preview your template before sending

5. **Send Emails:**
   - Review configuration summary
   - Click "Send All Emails" to start bulk sending
   - Monitor progress in real-time

## Excel File Format

Your Excel file should have the following columns:

| entity name | email |
|-------------|-------|
| Company A   | contact@companya.com |
| Company B   | info@companyb.com |

## SMTP Configuration

### Common SMTP Settings

| Provider | SMTP Server | Port | Security |
|----------|-------------|------|----------|
| Gmail | smtp.gmail.com | 587 | TLS |
| Outlook | smtp-mail.outlook.com | 587 | TLS |
| Yahoo | smtp.mail.yahoo.com | 587 | TLS |
| Lark Office | smtp.larkoffice.com | 465 | SSL |

### Security Notes

- Use app-specific passwords when available
- Never share your credentials
- Consider using environment variables for sensitive data

## Features Overview

### Email Configuration
- **Sender Email**: Your email address
- **SMTP Password**: Your email password or app-specific password
- **SMTP Server**: Email provider's SMTP server
- **SMTP Port**: Usually 465 (SSL) or 587 (TLS)
- **Subject**: Email subject line

### Recipients
- **To**: Primary recipients from Excel file
- **CC**: Carbon copy recipients (visible to all)
- **BCC**: Blind carbon copy recipients (hidden from others)

### Email Template
- Rich text editor for email content
- Company name personalization with `{company_name}` placeholder
- Preview functionality
- Default TikTok Shop business template included

### Progress Tracking
- Real-time progress bar
- Current recipient display
- Success/failure counting
- Detailed error reporting

### Results Dashboard
- Summary statistics (total, successful, failed)
- Detailed results table
- CSV export functionality
- Error message details

## Troubleshooting

### Common Issues

1. **SMTP Authentication Failed**
   - Check email and password
   - Enable "Less secure app access" or use app-specific password
   - Verify SMTP server and port settings

2. **Invalid Email Addresses**
   - Check Excel file for malformed email addresses
   - Ensure email column contains valid email formats

3. **File Upload Issues**
   - Ensure Excel file has required columns: `entity name`, `email`
   - Check file format (.xlsx or .xls)
   - Verify file is not corrupted

4. **Slow Sending**
   - Large recipient lists take time
   - SMTP servers may have rate limits
   - Network connectivity affects speed

### Tips for Success

- **Test First**: Start with a small test file (5-10 recipients)
- **Check Spam**: Verify emails aren't going to spam folders
- **Rate Limiting**: Some providers limit emails per hour/day
- **Backup**: Keep backup of your recipient lists
- **Monitoring**: Watch for bounce-back emails

## Security Best Practices

1. **Credentials**: Never hardcode passwords in the application
2. **App Passwords**: Use app-specific passwords when available
3. **Network**: Use secure networks for sending emails
4. **Data**: Don't store sensitive data in the application
5. **Testing**: Always test with small groups first

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your SMTP settings with your email provider
3. Test with a minimal Excel file first
4. Check application logs for detailed error messages

## License

This project is for internal use. Please ensure compliance with your organization's email policies and applicable laws regarding bulk email sending.# Email-Automation
