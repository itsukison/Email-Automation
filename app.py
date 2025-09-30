import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import time
import re
from typing import List, Dict, Optional
import io

# Page configuration
st.set_page_config(
    page_title="Email Automation Tool",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

def validate_email(email: str) -> bool:
    """Validate email address format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email.strip()) is not None

def parse_email_list(email_string: str) -> List[str]:
    """Parse comma-separated email list and validate each email"""
    if not email_string.strip():
        return []
    
    emails = [email.strip() for email in email_string.split(',')]
    valid_emails = []
    invalid_emails = []
    
    for email in emails:
        if email and validate_email(email):
            valid_emails.append(email)
        elif email:
            invalid_emails.append(email)
    
    if invalid_emails:
        st.error(f"Invalid email addresses: {', '.join(invalid_emails)}")
    
    return valid_emails

def process_excel_file(uploaded_file) -> Optional[pd.DataFrame]:
    """Process uploaded Excel file and validate structure"""
    try:
        df = pd.read_excel(uploaded_file)
        
        # Check for required columns
        required_columns = ['entity name', 'email']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            st.info("Required columns: 'entity name', 'email'")
            return None
        
        # Validate email addresses in the file
        invalid_emails = []
        for idx, email in enumerate(df['email']):
            if pd.isna(email) or not validate_email(str(email)):
                invalid_emails.append(f"Row {idx + 1}: {email}")
        
        if invalid_emails:
            st.warning(f"Found {len(invalid_emails)} invalid email addresses:")
            for invalid in invalid_emails[:5]:  # Show first 5
                st.write(f"- {invalid}")
            if len(invalid_emails) > 5:
                st.write(f"... and {len(invalid_emails) - 5} more")
        
        return df
        
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

def generate_email_body(template: str, company_name: str) -> str:
    """Generate personalized email body"""
    return template.replace("{company_name}", company_name)

def send_emails(df: pd.DataFrame, config: Dict, template: str) -> List[Dict]:
    """Send emails to all recipients with progress tracking"""
    results = []
    
    try:
        # Setup SMTP connection
        server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
        server.login(config['sender_email'], config['sender_password'])
        
        # Create progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_emails = len(df)
        
        for idx, row in df.iterrows():
            company_name = row['entity name']
            recipient_email = row['email']
            
            # Update progress
            progress = (idx + 1) / total_emails
            progress_bar.progress(progress)
            status_text.text(f"Sending to {company_name} ({idx + 1}/{total_emails})")
            
            try:
                # Generate personalized email content
                email_body = generate_email_body(template, company_name)
                
                # Create email message
                msg = MIMEText(email_body, 'plain', 'utf-8')
                msg['Subject'] = Header(config['subject'], 'utf-8')
                msg['From'] = config['sender_email']
                msg['To'] = recipient_email
                
                # Add CC recipients
                if config['cc_emails']:
                    msg['Cc'] = ', '.join(config['cc_emails'])
                
                # Prepare recipient list (To + CC + BCC)
                all_recipients = [recipient_email]
                if config['cc_emails']:
                    all_recipients.extend(config['cc_emails'])
                if config['bcc_emails']:
                    all_recipients.extend(config['bcc_emails'])
                
                # Send email
                server.sendmail(config['sender_email'], all_recipients, msg.as_string())
                
                results.append({
                    'company': company_name,
                    'email': recipient_email,
                    'status': 'Success',
                    'message': 'Email sent successfully'
                })
                
                # Small delay to avoid overwhelming the server
                time.sleep(0.5)
                
            except Exception as e:
                results.append({
                    'company': company_name,
                    'email': recipient_email,
                    'status': 'Failed',
                    'message': str(e)
                })
        
        server.quit()
        progress_bar.progress(1.0)
        status_text.text("Email sending completed!")
        
    except Exception as e:
        st.error(f"SMTP connection error: {str(e)}")
        results.append({
            'company': 'N/A',
            'email': 'N/A',
            'status': 'Failed',
            'message': f"SMTP Error: {str(e)}"
        })
    
    return results

def display_results(results: List[Dict]) -> None:
    """Display email sending results"""
    if not results:
        return
    
    results_df = pd.DataFrame(results)
    
    # Summary statistics
    total = len(results_df)
    success = len(results_df[results_df['status'] == 'Success'])
    failed = len(results_df[results_df['status'] == 'Failed'])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Emails", total)
    with col2:
        st.metric("Successful", success, delta=f"{success/total*100:.1f}%")
    with col3:
        st.metric("Failed", failed, delta=f"{failed/total*100:.1f}%" if failed > 0 else "0%")
    
    # Detailed results
    st.subheader("Detailed Results")
    st.dataframe(results_df, use_container_width=True)
    
    # Download results
    csv = results_df.to_csv(index=False)
    st.download_button(
        label="Download Results as CSV",
        data=csv,
        file_name="email_results.csv",
        mime="text/csv"
    )

def main():
    st.title("📧 Email Automation Tool")
    st.markdown("Send personalized bulk emails to companies from Excel files")
    
    # Initialize session state
    if 'email_data' not in st.session_state:
        st.session_state.email_data = None
    if 'results' not in st.session_state:
        st.session_state.results = []
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("📋 Configuration")
        
        # Email Configuration
        st.subheader("SMTP Settings")
        sender_email = st.text_input(
            "Sender Email",
            value="itsuki.son@bytedance.com",
            help="Your email address"
        )
        
        sender_password = st.text_input(
            "SMTP Password",
            type="password",
            help="Your email password or app-specific password"
        )
        
        smtp_server = st.text_input(
            "SMTP Server",
            value="smtp.larkoffice.com",
            help="SMTP server address"
        )
        
        smtp_port = st.number_input(
            "SMTP Port",
            value=465,
            min_value=1,
            max_value=65535,
            help="SMTP port (usually 465 for SSL or 587 for TLS)"
        )
        
        # Email Recipients
        st.subheader("Recipients")
        cc_emails_input = st.text_area(
            "CC Recipients",
            value="",
            help="Comma-separated email addresses for CC"
        )
        
        bcc_emails_input = st.text_area(
            "BCC Recipients",
            value="michinari.nakai@bytedance.com, kentaro.koi@bytedance.com",
            help="Comma-separated email addresses for BCC"
        )
        
        # Email Subject
        subject = st.text_input(
            "Email Subject",
            value="TikTok Shop出店のご案内",
            help="Subject line for all emails"
        )
    
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Upload Excel file with company data",
            type=['xlsx', 'xls'],
            help="File should contain 'entity name' and 'email' columns"
        )
        
        if uploaded_file is not None:
            with st.spinner("Processing file..."):
                df = process_excel_file(uploaded_file)
                if df is not None:
                    st.session_state.email_data = df
                    st.success(f"✅ Loaded {len(df)} companies")
                    
                    # Preview data
                    st.subheader("Data Preview")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    if len(df) > 10:
                        st.info(f"Showing first 10 rows of {len(df)} total rows")
    
    with col2:
        st.header("✉️ Email Template")
        
        # Default email template
        default_template = """{company_name} EC事業部 ご担当者様 

お世話になっております。
TikTok Japan EC事業（ByteDance株式会社）の孫と申します。

既にTikTok Shopにご出店いただいておりましたら、重ねてのご案内となり失礼いたします。 まだご参加でない場合、ぜひこの機会にご検討いただきたく、ご連絡を差し上げました。

TikTok Shopは従来の「検索型EC」とは異なり、動画とレコメンドを掛け合わせた 「発見型EC（ディスカバリーコマース）」 を実現します。 ユーザーは検索せずとも興味に合った商品に出会い、その場で購入まで完結できる新しい購買体験を楽しんでいます。

出店いただくことで、
- フォロワー数に関わらず、バズを通じた新規顧客への大規模なリーチ
- 動画やライブ配信からシームレスに購入できる 高いCVR（購入転換率）
- 低コスト・低リスクでの新しい販路開拓
 
といったメリットをご享受いただけます。
より詳しいご案内やQ&Aを含めた説明セミナーも予定しております。

まずは下記のアンケートにご回答いただき、セミナーにご参加いただければ幸いです。

▼アンケートリンク
https://bytedance.sg.larkoffice.com/share/base/form/shrlgWrMb9JHyZiMLsLdCaBs7Cd

ご検討のほど、何卒よろしくお願いいたします。
――――――――――――
 孫 逸歓
 Global E-commerce Japan
 Senior Manager of Beauty & Fashion
 ByteDance株式会社
 東京都渋谷区渋谷2-21-1 渋谷ヒカリエ 26F
 itsuki.son@bytedance.com"""
        
        email_template = st.text_area(
            "Email Template",
            value=default_template,
            height=400,
            help="Use {company_name} as placeholder for company name"
        )
        

    
    # Email sending section
    st.header("🚀 Send Emails")
    
    if st.session_state.email_data is not None:
        # Parse CC and BCC emails
        cc_emails = parse_email_list(cc_emails_input)
        bcc_emails = parse_email_list(bcc_emails_input)
        
        # Validation
        config_valid = True
        if not sender_email or not validate_email(sender_email):
            st.error("Please enter a valid sender email address")
            config_valid = False
        
        if not sender_password:
            st.error("Please enter SMTP password")
            config_valid = False
        
        if not email_template.strip():
            st.error("Please enter email template")
            config_valid = False
        
        if not subject.strip():
            st.error("Please enter email subject")
            config_valid = False
        
        # Display configuration summary
        if config_valid:
            with st.expander("📋 Configuration Summary", expanded=False):
                st.write(f"**Sender:** {sender_email}")
                st.write(f"**SMTP Server:** {smtp_server}:{smtp_port}")
                st.write(f"**Subject:** {subject}")
                st.write(f"**Recipients:** {len(st.session_state.email_data)} companies")
                if cc_emails:
                    st.write(f"**CC:** {', '.join(cc_emails)}")
                if bcc_emails:
                    st.write(f"**BCC:** {', '.join(bcc_emails)}")
        
        # Send button
        if config_valid:
            if st.button("🚀 Send All Emails", type="primary"):
                config = {
                    'sender_email': sender_email,
                    'sender_password': sender_password,
                    'smtp_server': smtp_server,
                    'smtp_port': smtp_port,
                    'subject': subject,
                    'cc_emails': cc_emails,
                    'bcc_emails': bcc_emails
                }
                
                with st.spinner("Sending emails..."):
                    results = send_emails(st.session_state.email_data, config, email_template)
                    st.session_state.results = results
    else:
        st.info("Please upload an Excel file to proceed")
    
    # Display results
    if st.session_state.results:
        st.header("📊 Results")
        display_results(st.session_state.results)
    
    # Footer
    st.markdown("---")
    st.markdown("💡 **Tips:**")
    st.markdown("- Use `{company_name}` in your template to personalize emails")
    st.markdown("- Test with a small file first before sending to large lists")
    st.markdown("- Keep your SMTP credentials secure")

if __name__ == "__main__":
    main()