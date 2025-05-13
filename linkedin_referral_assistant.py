#!/usr/bin/env python3
"""
LinkedIn Referral Assistant

This script automates the process of:
1. Reading unread LinkedIn messages
2. Identifying referral requests
3. Processing resumes from attachments or Google Drive links
4. Analyzing resumes using Azure OpenAI
5. Saving relevant information to an Excel sheet

Usage:
python linkedin_referral_assistant.py
"""

import os
import sys
import platform
import json
from datetime import datetime
from linkedin_messages import get_linkedin_messages, RESUMES_DIR
from analyze_referrals import process_linkedin_messages

def get_default_profile_path():
    """Get the default Chrome profile path based on the operating system."""
    system = platform.system()
    home = os.path.expanduser("~")
    
    if system == "Windows":
        return os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'User Data')
    elif system == "Darwin":  # macOS
        # Create a custom Chrome profile directory specifically for automation
        custom_profile = os.path.join(home, 'Library', 'Application Support', 'LinkedIn-Assistant-Profile')
        if not os.path.exists(custom_profile):
            os.makedirs(custom_profile, exist_ok=True)
            print(f"Created custom Chrome profile directory: {custom_profile}")
        return custom_profile
    elif system == "Linux":
        return os.path.join(home, '.config', 'google-chrome')
    else:
        return None

def main():
    print("=" * 60)
    print("LinkedIn Referral Assistant")
    print("=" * 60)
    print("This tool will help you process LinkedIn referral requests automatically.")
    print("It will:")
    print("1. Read your unread LinkedIn messages")
    print("2. Identify which messages are referral requests")
    print("3. Extract and save resumes to the 'resumes' folder")
    print("4. Analyze resumes using Azure OpenAI")
    print("5. Compile key information in an Excel sheet")
    print("=" * 60)
    
    # Ask for Chrome profile path
    print("\nEnter your Chrome profile path (press Enter to use default):")
    profile_path = input().strip()
    
    # If empty, try to use default
    if not profile_path:
        default_path = get_default_profile_path()
        if default_path:
            print(f"Using default Chrome profile path: {default_path}")
            profile_path = default_path
        else:
            print("Could not determine default profile path.")
            profile_path = input("Please enter your Chrome profile path: ")
    
    # Create resumes directory if it doesn't exist
    if not os.path.exists(RESUMES_DIR):
        os.makedirs(RESUMES_DIR)
    
    # Step 1: Read unread LinkedIn messages
    print("\n--- Step 1: Reading Unread LinkedIn Messages ---")
    messages = get_linkedin_messages(profile_path)
    
    # Handle case where no messages were found or there was an error
    if not messages or not messages.get("unread_messages"):
        print("\n--- Scan Complete ---")
        return  # Exit gracefully, not as an error
    
    # Check if any referral requests were detected
    unread_messages = messages.get("unread_messages", [])
    referral_requests = [msg for msg in unread_messages if msg.get("is_potential_referral", False)]
    referral_count = len(referral_requests)
    
    print(f"\nFound {len(unread_messages)} unread messages, {referral_count} of which appear to be referral requests.")
    
    if referral_count == 0:
        print("\n--- Scan Complete ---")
        print("No referral requests were detected in your unread messages.")
        return
    
    # Save the referral messages to a file for later analysis
    # This ensures analyze_referrals.py can find them even if run separately
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        referrals_file = f"linkedin_referrals_{timestamp}.json"
        with open(referrals_file, 'w', encoding='utf-8') as f:
            json.dump(referral_requests, f, indent=2)
        print(f"Saved referral data to {referrals_file} for later analysis")
    except Exception as e:
        print(f"Warning: Could not save referral data to file: {str(e)}")
        
    # Process each referral request
    print("\n--- Processing Referral Requests ---")
    resumes_found = 0
    
    for request in referral_requests:
        sender = request.get('sender', 'Unknown')
        has_resume = request.get('has_resume_downloaded', False)
        
        # Count resumes
        if has_resume:
            resumes_found += 1
            print(f"✓ Processed referral request from {sender} - Resume found and saved")
            
            # If it has Google Drive links, mention them
            if request.get('google_drive_links'):
                links = request.get('google_drive_links')
                print(f"  - Downloaded resume from Google Drive: {links[0]}")
        else:
            print(f"! Referral request from {sender} - No resume attached")
            
            # If it has Google Drive links but couldn't download, mention the issue
            if request.get('google_drive_links'):
                links = request.get('google_drive_links')
                print(f"  - Google Drive link found but could not download: {links[0]}")
                print("    (This may be due to permission settings on the file)")
    
    # Summary
    print("\n--- Summary ---")
    if resumes_found > 0:
        print(f"✓ Found {resumes_found} resumes out of {referral_count} referral requests")
        print(f"✓ Resumes saved to the '{RESUMES_DIR}' folder")
        
        # List the files in the resumes directory
        resume_files = os.listdir(RESUMES_DIR)
        if resume_files:
            print("\nSaved resume files:")
            for file in resume_files:
                print(f"  - {file}")
        
        # Step 2: Process the resumes and create Excel
        print("\n--- Creating Excel Report ---")
        result = process_linkedin_messages()
        
        if result:
            print("✓ Excel report created successfully with extracted information")
            print("✓ You can find it in 'linkedin_referrals.xlsx'")
        else:
            print("! Could not create Excel report. Check for errors above.")
    else:
        print("No resumes were found in the referral requests.")
        print("You may need to ask these contacts to send their resumes:")
        for msg in referral_requests:
            if not msg.get('has_resume_downloaded', False):
                print(f"  - {msg.get('sender', 'Unknown')}")
    
    print("\nThank you for using the LinkedIn Referral Assistant!")

if __name__ == "__main__":
    main() 