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
        "attached resume", "attached cv", "resume attached", "cv attached", "apply for",
        "job id", "job reference", "refer", "would appreciate a referral", "requesting a referral",
        "kindly refer", "please refer", "seeking a referral", "job referral", "application process",
        "help with referral", "refer my profile", "refer my application", "refer my resume",
        "LinkedIn job", "request you to refer", "job consideration", "open position", "open role",
        "open requisition", "req id", "requisition"
    ]
    
    if not message_content:
        return False
    
    message_lower = message_content.lower()
    
    # Check for job ID patterns (common in referral requests)
    job_id_pattern = re.search(r'job\s*id\s*[:#]?\s*([a-z]\d+)', message_lower, re.IGNORECASE)
    if job_id_pattern:
        return True
    
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
    print("  🔄 Extracting message content...")
    full_messages = []
    message_items = page.query_selector_all('.msg-s-message-list__event')
    
    # Create resumes directory if it doesn't exist
    if not os.path.exists(RESUMES_DIR):
        os.makedirs(RESUMES_DIR)
    
    # Get the name from the conversation info, but also try to find the real name in the message
    display_name = conversation_info.get('sender', 'Unknown')
    sender_name = display_name.replace(' ', '_')
    
    # Check if we can find the signature or a more accurate name in the message content
    # We'll collect this during message analysis
    actual_name = None
    is_potential_referral = False
    all_message_content = ""
    has_attachments = False
    has_drive_links = False
    drive_links = []
    
    # First pass to collect all message content for analysis
    print(f"  📃 Analyzing message text to identify referral request...")
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
                print(f"  🔍 Found Google Drive link in message")
                
        # Check if message has attachments
        if item.query_selector_all('.msg-s-event-listitem__attachment-item'):
            has_attachments = True
            print(f"  🔍 Found attachment in message")
    
    # Try to extract actual name from message (look for signature patterns)
    signature_patterns = [
        r'(?:Regards|Sincerely|Thanks|Thank you|Best|Cheers|Best regards),?\s*\n*([A-Z][a-z]+(?: [A-Z][a-z]+)*)',
        r'(?:name is|this is)\s*([A-Z][a-z]+(?: [A-Z][a-z]+)*)',
        r'(?:^|\n)([A-Z][a-z]+(?: [A-Z][a-z]+)*)\s*$',  # Name at the end of message or line
        r'Dear\s+[A-Za-z\.]+,\s*\n+.*\n+.*(?:Regards|Sincerely|Thanks|Thank you|Best),?\s*\n*([A-Z][a-z]+(?: [A-Z][a-z]+)*)'  # Common email format with closing
    ]
    
    for pattern in signature_patterns:
        name_match = re.search(pattern, all_message_content)
        if name_match:
            actual_name = name_match.group(1).strip()
            if 2 <= len(actual_name.split()) <= 3:  # Likely a real name with 2-3 words
                # Check if this name is different from the display name
                if actual_name.lower() != display_name.lower():
                    print(f"  📝 Found different signature name in message: {actual_name}")
                else:
                    print(f"  📝 Found matching signature name in message: {actual_name}")
                break
    
    # If found a better name in the message, use it
    if actual_name:
        if actual_name.lower() != display_name.lower():
            sender_name = actual_name.replace(' ', '_')
            print(f"  ✏️ Using message signature name: {sender_name} (different from conversation contact)")
        else:
            sender_name = actual_name.replace(' ', '_')
            print(f"  ✏️ Using message signature name: {sender_name} (matches conversation contact)")
    else:
        print(f"  ✏️ Using conversation contact name: {display_name}")
    
    # Determine if this is a referral request
    is_potential_referral = is_referral_request(all_message_content)
    
    # If this is a referral request with Google Drive links, try to download them
    resume_downloaded = False
    if is_potential_referral and drive_links:
        print(f"  📥 Attempting to download resume from Google Drive...")
        for i, drive_link in enumerate(drive_links):
            clean_filename = f"{sender_name}_resume.pdf"
            # If multiple drive links, add a number suffix only for the 2nd onward
            if i > 0:
                clean_filename = f"{sender_name}_resume_{i+1}.pdf"
                
            result = download_from_google_drive(drive_link, RESUMES_DIR, clean_filename)
            if result.get("success"):
                resume_downloaded = True
                print(f"  ✅ Successfully downloaded resume: {clean_filename}")
            else:
                print(f"  ❌ Failed to download from Google Drive: {result.get('error', 'Unknown error')}")
    
    # Process each message item
    attachments_processed = False
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
        if is_potential_referral and not attachments_processed:
            attachments = item.query_selector_all('.msg-s-event-listitem__attachment-item')
            if attachments:
                print(f"  📎 Found {len(attachments)} attachment(s) in message")
                message_attachments = []
                downloaded_attachments = []
                for att_index, attachment in enumerate(attachments):
                    try:
                        # Look for the link with href containing "/dms/"
                        download_link = attachment.query_selector('a[href*="/dms/"]')
                        if download_link:
                            attachment_url = download_link.get_attribute('href')
                            attachment_name = download_link.get_attribute('download') or "resume.pdf"
                            
                            print(f"  📥 Attempting to download attachment {att_index+1}: {attachment_name}")
                            
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
                                    print(f"  ✅ Successfully downloaded resume: {clean_filename}")
                                else:
                                    print(f"  ❌ Failed to download attachment: {download_result.get('error', 'Unknown error')}")
                            
                    except Exception as e:
                        print(f"  ❌ Error processing attachment: {str(e)}")
                
                if message_attachments:
                    message_info["attachments"] = message_attachments
                    
                # Add a summary of downloaded attachments to the message
                if downloaded_attachments:
                    message_info["attachment_downloaded"] = ", ".join(downloaded_attachments)
                    attachments_processed = True
        
        if message_info:
            full_messages.append(message_info)
    
    # Add referral status to conversation info
    conversation_info["is_potential_referral"] = is_potential_referral
    
    # Add resume download status
    conversation_info["has_resume_downloaded"] = resume_downloaded
    
    # Add the actual name we found to the conversation info
    if actual_name:
        conversation_info["actual_name"] = actual_name
    
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
                print(f"⚠️ Profile path '{profile_path}' does not exist or was not provided.")
                print("⚠️ Starting new browser session. You will need to login manually.")
                browser = p.chromium.launch(headless=False, args=browser_args)
                context = browser.new_context(
                    accept_downloads=True,
                    user_agent=user_agent,
                    viewport={'width': 1280, 'height': 800}
                )
                page = context.new_page()
            else:
                print(f"🔐 Using existing profile: {profile_path}")
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
            print("Opening LinkedIn to check login status...")
            page.goto('https://www.linkedin.com/')
            
            if page.url.startswith('https://www.linkedin.com/login'):
                print("-----------------------------------------------------")
                print("LinkedIn login required! Please sign in manually in the browser window.")
                print("If you need to complete 2FA, you have up to 5 minutes to complete the process.")
                print("-----------------------------------------------------")
                
                # Wait for login to complete
                page.wait_for_url('https://www.linkedin.com/feed/', timeout=300000)  # 5 minutes for login + 2FA
                print("✅ Login completed successfully!")
            else:
                print("✅ Already logged in to LinkedIn")
            
            # Navigate to messaging
            print("Navigating to LinkedIn messaging...")
            page.goto('https://www.linkedin.com/messaging/')
            
            # Wait for the messages list to load
            print("⏳ Waiting for LinkedIn messages to load...")
            page.wait_for_selector('.msg-conversations-container__conversations-list', timeout=120000)  # 2 minutes
            page.wait_for_timeout(5000)  # Give more time for dynamic content
            
            # First, count visible conversations
            conversations = page.query_selector_all('.msg-conversation-listitem__link')
            print(f"📨 Initially found {len(conversations)} conversations visible")
            
            # Define target count before using it
            target_count = 15  # We want to check at least 15 messages
            
            # Scroll to load more conversations, to ensure we check at least 15 messages
            print(f"⏬ Scrolling to load at least {target_count} conversations to check for unread messages...")
            conversation_list = page.query_selector('.msg-conversations-container__conversations-list')
            max_scroll_attempts = 10
            scroll_attempts = 0
            
            # Track conversations we've already seen to ensure we're loading new ones
            current_conversation_count = len(conversations)
            last_conversation_count = 0
            
            while len(conversations) < target_count and scroll_attempts < max_scroll_attempts:
                last_conversation_count = len(conversations)
                print(f"Scroll attempt {scroll_attempts+1}: Currently have {last_conversation_count} conversations")
                
                # Scroll incrementally rather than all the way to the bottom
                if conversation_list:
                    # Simple scroll down by 300px
                    page.evaluate("element => { element.scrollTop += 300; }", conversation_list)
                else:
                    # Fallback scrolling if we can't get the conversation list element
                    page.keyboard.press('PageDown')
                
                # Wait for new items to load
                page.wait_for_timeout(2000)
                
                # Count conversations again
                conversations = page.query_selector_all('.msg-conversation-listitem__link')
                current_conversation_count = len(conversations)
                print(f"After scrolling: found {current_conversation_count} conversations")
                
                # Check if we've loaded new conversations
                if current_conversation_count <= last_conversation_count:
                    print("No new conversations loaded, waiting longer...")
                    page.wait_for_timeout(2000)  # Wait a bit longer
                    
                    # Try one more time
                    conversations = page.query_selector_all('.msg-conversation-listitem__link')
                    current_conversation_count = len(conversations)
                    
                    # If still no new conversations, we may have reached the end
                    if current_conversation_count <= last_conversation_count:
                        print("Still no new conversations. We may have reached the end of the list.")
                        # Try one final larger scroll
                        if conversation_list:
                            page.evaluate("element => { element.scrollTop += 800; }", conversation_list)
                        page.wait_for_timeout(3000)
                        conversations = page.query_selector_all('.msg-conversation-listitem__link')
                        if len(conversations) <= last_conversation_count:
                            print("No more conversations to load. Breaking out of scroll loop.")
                            break
                
                scroll_attempts += 1
            
            # Now check all loaded conversations for unread status
            unread_messages = []
            unread_conversation_elements = []
            
            # Find all conversations with unread status
            for convo in conversations:
                # Check if this conversation is unread
                is_unread = convo.query_selector('.msg-conversation-card__unread-count') is not None
                
                if is_unread:
                    unread_conversation_elements.append(convo)
            
            print(f"🔍 Found {len(unread_conversation_elements)} unread conversations out of {len(conversations)} total loaded")
            print("=" * 50)
            
            # Process each unread conversation
            if len(unread_conversation_elements) > 0:
                print("📝 Processing only the unread conversations...")
            else:
                print("ℹ️ No unread conversations found.")
            
            for i, convo in enumerate(unread_conversation_elements):
                try:
                    print("\n" + "=" * 50)
                    print(f"CONVERSATION {i+1}/{len(unread_conversation_elements)}")
                    print("=" * 50)
                    
                    # Get basic conversation info by examining the container
                    container = convo.evaluate("(element) => element.closest('.msg-conversation-card')")
                    if not container:
                        container = convo
                    
                    # Extract sender name and other info
                    sender_elem = page.query_selector_all('.msg-conversation-listitem__participant-names')[i]
                    sender_name = "Unknown"
                    if sender_elem:
                        name_span = sender_elem.query_selector('span.truncate')
                        if name_span:
                            sender_name = name_span.inner_text()
                        else:
                            sender_name = sender_elem.inner_text()
                    
                    preview = page.query_selector_all('.msg-conversation-card__message-snippet')[i] if i < len(page.query_selector_all('.msg-conversation-card__message-snippet')) else None
                    preview_text = preview.inner_text() if preview else ""
                    
                    convo_info = {
                        "sender": sender_name,
                        "preview": preview_text
                    }
                    
                    print(f"Opening conversation from LinkedIn contact: {convo_info['sender']}")
                    
                    # Click to open the conversation
                    convo.click()
                    print(f"  → Opening conversation...")
                    page.wait_for_selector('.msg-s-message-list__event', timeout=30000)  # 30 seconds
                    page.wait_for_timeout(3000)
                    
                    # Extract message content and add to results
                    current_browser = browser_context if browser_context else context
                    message_data = extract_message_content(page, convo_info, current_browser)
                    
                    # Only add to unread_messages if it's a potential referral request
                    if message_data.get("is_potential_referral", False):
                        print(f"  ✅ Identified as a REFERRAL REQUEST")
                        
                        # If we found an actual name different from the conversation name, make it clear
                        if message_data.get("actual_name") and message_data.get("actual_name") != convo_info["sender"]:
                            print(f"  ℹ️ Note: This appears to be a message from {message_data.get('actual_name')} sent through {convo_info['sender']}'s conversation")
                        
                        # Check if resume was found
                        if message_data.get("has_resume_downloaded", False):
                            print(f"  ✅ Resume found and downloaded")
                            
                            # Show if it was from Google Drive or attachment
                            if message_data.get("google_drive_links"):
                                print(f"  📥 Downloaded from Google Drive: {message_data['google_drive_links'][0]}")
                            else:
                                print(f"  📎 Downloaded from message attachment")
                            
                            print("  📊 Resume will be added to Excel report")
                        else:
                            print(f"  ❌ No resume found for this referral request")
                            
                        unread_messages.append(message_data)
                    else:
                        print(f"  ❌ Not a referral request - skipping")
                    
                    print("-" * 50)
                    
                except Exception as e:
                    print(f"  ❌ Error processing conversation {i+1}: {str(e)}")
                    print("-" * 50)
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
