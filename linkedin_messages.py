from playwright.sync_api import sync_playwright
import json
from datetime import datetime
import os
import re

def get_linkedin_messages(profile_path):
    with sync_playwright() as p:
        # 1. Launch browser with the specified profile
        browser = p.chromium.launch_persistent_context(
            user_data_dir=profile_path,
            headless=False  # Set to True if you don't want to see the browser
        )
        
        # Create a new page
        page = browser.new_page()
        
        try:
            # 2. Navigate to LinkedIn messaging
            print("Opening LinkedIn messaging...")
            page.goto('https://www.linkedin.com/messaging/')
            
            # 3. Wait for the messages list to load
            print("Waiting for messages to load...")
            page.wait_for_selector('.msg-conversations-container__conversations-list', timeout=30000)
            
            # Wait a bit for dynamic content to load
            page.wait_for_timeout(3000)
            
            # 4. Get all list items in the conversations list
            conversations = page.query_selector_all('.msg-conversations-container__conversations-list li')
            print(f"Found {len(conversations)} conversations")
            
            # Store all unread messages data
            messages_data = []
            
            # Find unread messages using the primary class indicator
            unread_conversations = page.query_selector_all('.msg-conversation-card__convo-item-container--unread')
            print(f"Found {len(unread_conversations)} unread conversations")
            
            # Process each unread conversation
            for i, convo in enumerate(unread_conversations):
                try:
                    # Get sender name
                    sender_elem = convo.query_selector('.msg-conversation-listitem__participant-names')
                    sender_name = "Unknown"
                    if sender_elem:
                        name_span = sender_elem.query_selector('span.truncate')
                        if name_span:
                            sender_name = name_span.inner_text()
                        else:
                            sender_name = sender_elem.inner_text()
                    
                    print(f"Processing unread conversation {i+1} from: {sender_name}")
                    
                    # Get preview and timestamp
                    preview = convo.query_selector('.msg-conversation-card__message-snippet')
                    preview_text = preview.inner_text() if preview else ""
                    
                    timestamp = convo.query_selector('.msg-conversation-card__time-stamp')
                    time_text = timestamp.inner_text() if timestamp else ""
                    
                    # Click on the conversation to view full content
                    clickable = convo.query_selector('.msg-conversation-listitem__link')
                    if clickable:
                        print(f"Opening conversation with {sender_name}...")
                        clickable.click()
                        
                        # Wait for message content to load
                        page.wait_for_selector('.msg-s-message-list__event', timeout=5000)
                        page.wait_for_timeout(2000)  # Wait a bit for everything to render
                        
                        # Get messages from the conversation
                        full_messages = []
                        message_items = page.query_selector_all('.msg-s-message-list__event')
                        
                        print(f"Found {len(message_items)} message events in this conversation")
                        
                        for item_index, item in enumerate(message_items):
                            # Check for sender info inside each message
                            message_info = {}
                            
                            # Get message body
                            body = item.query_selector('.msg-s-event-listitem__body')
                            if body:
                                message_content = body.inner_text()
                                message_info["content"] = message_content
                                print(f"Message {item_index+1} content: {message_content[:50]}..." if len(message_content) > 50 else message_content)
                                
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
                            
                            if message_info:
                                full_messages.append(message_info)
                        
                        # Add message data
                        messages_data.append({
                            "sender": sender_name,
                            "timestamp": time_text,
                            "preview": preview_text,
                            "messages": full_messages
                        })
                except Exception as e:
                    print(f"Error processing conversation {i+1}: {str(e)}")
                    continue
            
            # Save messages to a file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"linkedin_messages_{timestamp}.json"
            
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(messages_data, f, indent=2, ensure_ascii=False)
            
            print(f"Messages saved to {filename}")
            print(f"Found {len(messages_data)} unread conversations")
            
            if len(messages_data) == 0:
                print("No unread messages were found. This could be because:")
                print("1. You don't have any unread messages")
                print("2. The script couldn't identify the unread message indicators")
                print("3. LinkedIn's UI has changed and selectors need updating")
                
                # As a fallback, attempt to process the first few conversations regardless of read status
                print("\nAttempting to process the first 3 conversations as a fallback...")
                fallback_data = []
                
                for i in range(min(3, len(conversations))):
                    try:
                        convo = conversations[i]
                        # Similar processing as above but without unread checks
                        sender_elem = convo.query_selector('.msg-conversation-listitem__participant-names')
                        sender_name = sender_elem.inner_text() if sender_elem else "Unknown"
                        
                        clickable = convo.query_selector('.msg-conversation-listitem__link')
                        if clickable:
                            print(f"Opening conversation with {sender_name}...")
                            clickable.click()
                            
                            page.wait_for_selector('.msg-s-message-list__event', timeout=5000)
                            page.wait_for_timeout(2000)
                            
                            full_messages = []
                            message_items = page.query_selector_all('.msg-s-message-list__event')
                            
                            for item in message_items:
                                body = item.query_selector('.msg-s-event-listitem__body')
                                if body:
                                    full_messages.append({"content": body.inner_text()})
                            
                            fallback_data.append({
                                "sender": sender_name,
                                "messages": full_messages
                            })
                    except Exception as e:
                        print(f"Error processing fallback conversation {i+1}: {str(e)}")
                
                # Save fallback data
                fallback_filename = f"linkedin_messages_fallback_{timestamp}.json"
                with open(fallback_filename, 'w', encoding='utf-8') as f:
                    json.dump(fallback_data, f, indent=2, ensure_ascii=False)
                
                print(f"Fallback data saved to {fallback_filename}")
            
            return messages_data
            
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            return None
        
        finally:
            browser.close()

if __name__ == "__main__":
    # 1. Ask for Chrome profile path
    profile_path = input("Please enter your Chrome profile path: ")
    messages = get_linkedin_messages(profile_path) 