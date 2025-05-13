from playwright.sync_api import sync_playwright
import json
from datetime import datetime
import os
import re
import requests
import time
import urllib.parse

# Constants
RESUMES_DIR = "resumes"

def download_from_google_drive(drive_link, destination_folder, filename):
    """Download a file from Google Drive public link."""
    try:
        # Create directory if it doesn't exist
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        
        # Extract file ID from various Google Drive link formats
        file_id = None
        
        # Format: https://drive.google.com/file/d/FILE_ID/view
        file_id_match = re.search(r'\/file\/d\/([^\/]+)', drive_link)
        if file_id_match:
            file_id = file_id_match.group(1)
        
        # Format: https://drive.google.com/open?id=FILE_ID
        if not file_id:
            parsed_url = urllib.parse.urlparse(drive_link)
            params = urllib.parse.parse_qs(parsed_url.query)
            if 'id' in params:
                file_id = params['id'][0]
        
        if not file_id:
            print(f"Could not extract file ID from Drive link: {drive_link}")
            return {"success": False, "error": "Invalid Google Drive link format"}
        
        # Direct download link for Google Drive files
        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        
        # For larger files, we need to handle the confirmation page
        session = requests.Session()
        
        print(f"Downloading file from Google Drive: {drive_link}")
        response = session.get(download_url, stream=True)
        
        # Check if there's a download warning (for large files)
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                download_url = f"{download_url}&confirm={value}"
                response = session.get(download_url, stream=True)
                break
        
        # Save the file
        file_path = os.path.join(destination_folder, filename)
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        print(f"File successfully downloaded from Google Drive to: {file_path}")
        return {"success": True, "file_path": file_path}
    
    except Exception as e:
        print(f"Error downloading from Google Drive: {str(e)}")
        return {"success": False, "error": str(e)}

def is_referral_request(message_content):
    """Analyze message content to determine if it's a referral request."""
    referral_keywords = [
        "referral", "refer me", "referring", "job application", "applying", "opportunity",
        "position", "role", "job posting", "job opening", "recommend me", "recommendation",
        "looking for a job", "job search", "openings", "hiring", "job opportunity",
        "internal referral", "company referral", "forwarding my resume", "forwarding my cv",
        "attached resume", "attached cv", "resume attached", "cv attached", "apply for"
    ]
    
    if not message_content:
        return False
    
    message_lower = message_content.lower()
    
    # Check for referral keywords
    for keyword in referral_keywords:
        if keyword in message_lower:
            return True
    
    return False

def extract_basic_conversation_info(convo):
    """Extract basic information from a conversation element."""
    sender_elem = convo.query_selector('.msg-conversation-listitem__participant-names')
    sender_name = "Unknown"
    if sender_elem:
        name_span = sender_elem.query_selector('span.truncate')
        if name_span:
            sender_name = name_span.inner_text()
        else:
            sender_name = sender_elem.inner_text()
    
    preview = convo.query_selector('.msg-conversation-card__message-snippet')
    preview_text = preview.inner_text() if preview else ""
    
    timestamp = convo.query_selector('.msg-conversation-card__time-stamp')
    time_text = timestamp.inner_text() if timestamp else ""
    
    return {
        "sender": sender_name,
        "timestamp": time_text,
        "preview": preview_text
    }

def download_file(url, browser_context, destination_folder, filename):
    """Download a file directly using requests with cookies from browser context."""
    try:
        # Create directory if it doesn't exist
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
            
        # Get cookies from browser context
        cookies = browser_context.cookies()
        cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies}
        
        # Get headers from browser context
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36',
            'Accept': '*/*',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://www.linkedin.com/'
        }
        
        # Make the request to download the file
        print(f"Downloading file from: {url}")
        response = requests.get(url, cookies=cookies_dict, headers=headers, stream=True)
        
        if response.status_code == 200:
            file_path = os.path.join(destination_folder, filename)
            
            # Save the file
            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        
            print(f"File successfully downloaded to: {file_path}")
            return {"success": True, "file_path": file_path}
        else:
            print(f"Failed to download file. Status code: {response.status_code}")
            return {"success": False, "error": f"HTTP status code: {response.status_code}"}
    except Exception as e:
        print(f"Error downloading file: {str(e)}")
        return {"success": False, "error": str(e)}

def extract_message_content(page, conversation_info, browser_context):
    """Extract message content from an opened conversation."""
    full_messages = []
    message_items = page.query_selector_all('.msg-s-message-list__event')
    
    # Create resumes directory if it doesn't exist
    if not os.path.exists(RESUMES_DIR):
        os.makedirs(RESUMES_DIR)
    
    sender_name = conversation_info.get('sender', 'Unknown').replace(' ', '_')
    is_potential_referral = False
    all_message_content = ""
    has_attachments = False
    has_drive_links = False
    drive_links = []
    
    # First pass to collect all message content for analysis
    for item in message_items:
        body = item.query_selector('.msg-s-event-listitem__body')
        if body:
            message_content = body.inner_text()
            all_message_content += message_content + " "
            
            # Collect Google Drive links
            links = re.findall(r'(https://drive\.google\.com/\S+)', message_content)
            if links:
                has_drive_links = True
                drive_links.extend(links)
                
        # Check if message has attachments
        if item.query_selector_all('.msg-s-event-listitem__attachment-item'):
            has_attachments = True
    
    # Determine if this is a referral request
    is_potential_referral = is_referral_request(all_message_content)
    
    # If this is a referral request with Google Drive links, try to download them
    resume_downloaded = False
    if is_potential_referral and drive_links:
        for i, drive_link in enumerate(drive_links):
            clean_filename = f"{sender_name}_resume.pdf"
            # If multiple drive links, add a number suffix only for the 2nd onward
            if i > 0:
                clean_filename = f"{sender_name}_resume_{i+1}.pdf"
                
            result = download_from_google_drive(drive_link, RESUMES_DIR, clean_filename)
            if result.get("success"):
                resume_downloaded = True
    
    # Process each message item
    for item_index, item in enumerate(message_items):
        message_info = {}
        
        # Get message body
        body = item.query_selector('.msg-s-event-listitem__body')
        if body:
            message_content = body.inner_text()
            message_info["content"] = message_content
            
            # Check for Google Drive links in the message content if it's a referral
            if is_potential_referral:
                links = re.findall(r'(https://drive\.google\.com/\S+)', message_content)
                if links:
                    message_info["google_drive_links"] = links
            
            # Check for emails in the message content
            email_links = item.query_selector_all('a[href^="mailto:"]')
            if email_links:
                emails = []
                for email_link in email_links:
                    email_href = email_link.get_attribute('href')
                    if email_href and email_href.startswith('mailto:'):
                        emails.append(email_href.replace('mailto:', ''))
                if emails:
                    message_info["emails"] = emails
        
        # Check for file attachments only if this looks like a referral request
        if is_potential_referral:
            attachments = item.query_selector_all('.msg-s-event-listitem__attachment-item')
            if attachments:
                message_attachments = []
                downloaded_attachments = []
                for attachment in attachments:
                    try:
                        # Look for the link with href containing "/dms/"
                        download_link = attachment.query_selector('a[href*="/dms/"]')
                        if download_link:
                            attachment_url = download_link.get_attribute('href')
                            attachment_name = download_link.get_attribute('download') or "resume.pdf"
                            
                            # If it appears to be a resume, save it with the sender's name
                            resume_keywords = ["resume", "cv", "curriculum", "vitae"]
                            is_likely_resume = any(keyword in attachment_name.lower() for keyword in resume_keywords)
                            
                            # For PDF files or files that appear to be resumes
                            if is_likely_resume or attachment_name.lower().endswith('.pdf'):
                                # Generate a clean filename with sender's name
                                clean_filename = f"{sender_name}_resume.pdf"
                                if not clean_filename.lower().endswith('.pdf'):
                                    clean_filename += '.pdf'
                                
                                # Download the resume to the resumes folder
                                download_result = download_file(
                                    url=attachment_url,
                                    browser_context=browser_context,
                                    destination_folder=RESUMES_DIR,
                                    filename=clean_filename
                                )
                                
                                if download_result.get("success", False):
                                    file_path = download_result.get("file_path")
                                    message_attachments.append({
                                        "filename": clean_filename,
                                        "saved_path": file_path,
                                        "original_url": attachment_url,
                                        "is_resume": True
                                    })
                                    downloaded_attachments.append(clean_filename)
                                    resume_downloaded = True
                            
                    except Exception as e:
                        print(f"Error processing attachment: {str(e)}")
                
                if message_attachments:
                    message_info["attachments"] = message_attachments
                    
                # Add a summary of downloaded attachments to the message
                if downloaded_attachments:
                    message_info["attachment_downloaded"] = ", ".join(downloaded_attachments)
        
        if message_info:
            full_messages.append(message_info)
    
    # Add referral status to conversation info
    conversation_info["is_potential_referral"] = is_potential_referral
    
    # Add resume download status
    conversation_info["has_resume_downloaded"] = resume_downloaded
    
    # Explicitly add Google Drive links to the conversation info
    if drive_links:
        conversation_info["google_drive_links"] = drive_links
    
    # Return combined information
    return {
        **conversation_info,
        "messages": full_messages
    }

def get_linkedin_messages(profile_path=None):
    with sync_playwright() as p:
        browser = None
        context = None
        browser_context = None
        
        try:
            # Launch arguments to avoid detection
            browser_args = [
                '--disable-blink-features=AutomationControlled',
                '--no-sandbox',
                '--disable-infobars',
                '--disable-dev-shm-usage',
                '--disable-browser-side-navigation',
                '--disable-features=IsolateOrigins,site-per-process'
            ]
            
            user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
            
            # Check if the profile path exists
            if not profile_path or not os.path.exists(profile_path):
                print(f"Warning: Profile path '{profile_path}' does not exist or was not provided.")
                print("Launching browser without a profile. You will need to login manually.")
                browser = p.chromium.launch(headless=False, args=browser_args)
                context = browser.new_context(
                    accept_downloads=True,
                    user_agent=user_agent,
                    viewport={'width': 1280, 'height': 800}
                )
                page = context.new_page()
            else:
                print(f"Launching browser with profile: {profile_path}")
                browser_context = p.chromium.launch_persistent_context(
                    user_data_dir=profile_path,
                    headless=False,
                    accept_downloads=True,
                    args=browser_args,
                    user_agent=user_agent,
                    viewport={'width': 1280, 'height': 800},
                    ignore_default_args=['--enable-automation']
                )
                page = browser_context.new_page()
            
            # Add JavaScript to avoid detection
            page.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });
                window.navigator.chrome = { runtime: {} };
            """)
            
            # Increase timeout for page loads
            page.set_default_timeout(120000)  # 2 minutes
            
            # Check login status first
            print("Opening LinkedIn...")
            page.goto('https://www.linkedin.com/')
            
            if page.url.startswith('https://www.linkedin.com/login'):
                print("Login required. Please sign in manually in the browser window.")
                print("If you need to complete 2FA, you have up to 5 minutes to complete the process.")
                page.wait_for_url('https://www.linkedin.com/feed/', timeout=300000)  # 5 minutes for login + 2FA
                print("Login successful!")
            
            # Navigate to messaging
            print("Opening LinkedIn messaging...")
            page.goto('https://www.linkedin.com/messaging/')
            
            # Wait for the messages list to load
            print("Waiting for messages to load...")
            page.wait_for_selector('.msg-conversations-container__conversations-list', timeout=120000)  # 2 minutes
            page.wait_for_timeout(5000)  # Give more time for dynamic content
            
            # Focus on unread conversations
            unread_conversations = page.query_selector_all('.msg-conversation-card__convo-item-container--unread')
            print(f"Found {len(unread_conversations)} unread conversations")
            
            unread_messages = []
            
            # Process each unread conversation
            for i, convo in enumerate(unread_conversations):
                try:
                    # Get basic conversation info
                    convo_info = extract_basic_conversation_info(convo)
                    print(f"Processing conversation {i+1} from: {convo_info['sender']}")
                    
                    # Click to open the conversation
                    clickable = convo.query_selector('.msg-conversation-listitem__link')
                    if clickable:
                        clickable.click()
                        page.wait_for_selector('.msg-s-message-list__event', timeout=30000)  # 30 seconds
                        page.wait_for_timeout(3000)
                        
                        # Extract message content and add to results
                        current_browser = browser_context if browser_context else context
                        message_data = extract_message_content(page, convo_info, current_browser)
                        unread_messages.append(message_data)
                        
                except Exception as e:
                    print(f"Error processing conversation {i+1}: {str(e)}")
                    continue
            
            # Return unread messages
            return {
                "unread_messages": unread_messages
            }
            
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            return None
        
        finally:
            # Clean up resources
            if browser_context:
                browser_context.close()
            if context:
                context.close()
            if browser:
                browser.close()
