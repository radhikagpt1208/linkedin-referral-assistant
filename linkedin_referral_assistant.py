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
    print("\n" + "=" * 60)
    print("PROCESSING REFERRAL REQUESTS")
    print("=" * 60)
    resumes_found = 0
    
    for i, request in enumerate(referral_requests):
        sender = request.get('sender', 'Unknown')
        has_resume = request.get('has_resume_downloaded', False)
        
        print(f"\n{'-' * 50}")
        print(f"REFERRAL {i+1}/{len(referral_requests)}: {sender}")
        print(f"{'-' * 50}")
        
        # Check if there's a mismatch between conversation contact and actual sender
        actual_name = request.get('actual_name')
        if actual_name and actual_name != sender:
            print(f"ℹ️ Note: This appears to be from {actual_name} sent through {sender}'s conversation")
            # Use the actual name for display
            sender = actual_name
        
        # Count resumes
        if has_resume:
            resumes_found += 1
            print(f"✓ Resume Status: Resume found and saved")
            
            # If it has Google Drive links, mention them
            if request.get('google_drive_links'):
                links = request.get('google_drive_links')
                print(f"  → Source: Google Drive link")
                print(f"  → Link: {links[0]}")
        else:
            print(f"✗ Resume Status: No resume attached")
            
            # If it has Google Drive links but couldn't download, mention the issue
            if request.get('google_drive_links'):
                links = request.get('google_drive_links')
                print(f"  → Found Google Drive link but could not download")
                print(f"  → Link: {links[0]}")
                print("  → Possible reason: Permission settings on the file")
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
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
        print("\n" + "=" * 60)
        print("CREATING EXCEL REPORT")
        print("=" * 60)
        print("Note: Only referral requests with resume files will be included in the Excel report")
        result = process_linkedin_messages()
        
        if result:
            print("\n" + "=" * 60)
            print("SUCCESS")
            print("=" * 60)
            print("✓ Excel report created successfully with extracted information")
            print("✓ You can find it in 'linkedin_referrals.xlsx'")
        else:
            print("\n" + "=" * 60)
            print("ERROR")
            print("=" * 60)
            print("! Could not create Excel report. Check for errors above.")
    else:
        print("\nNo resumes were found in the referral requests.")
        print("You may need to ask these contacts to send their resumes:")
        for msg in referral_requests:
            if not msg.get('has_resume_downloaded', False):
                print(f"  - {msg.get('sender', 'Unknown')}")
    
    print("\n" + "=" * 60)
    print("PROCESS COMPLETE")
    print("=" * 60)
    print("Thank you for using the LinkedIn Referral Assistant!")

if __name__ == "__main__":
    main() 