"""
CLI: orchestrate FreeScout email sending (reads spreadsheet via sponsor_reader, sends via freescout_automation).

Inputs: path to .xlsx / .xls / .csv — see Reachout/README.md §5 (standard vs round CSV, required columns).

Round workflow:
  1. python extract_round_leads.py "path/to/master.xlsx" [-o round.csv]
  2. python main.py round.csv [--start-from N]
"""
import argparse
import sys
import os
import re
from datetime import datetime
from sponsor_reader import SponsorReader
from freescout_automation import FreeScoutAutomation
from config import (
    FREESCOUT_URL,
    FREESCOUT_EMAIL,
    FREESCOUT_PASSWORD,
    FILTER_BY_OUTREACH,
    OUTREACH_FILTER_VALUE,
    TEMPLATE_NAME_COLUMN,
)


def _is_browser_session_lost(exc: BaseException) -> bool:
    """True if the exception indicates Chrome/WebDriver session is gone (connection refused, etc.)."""
    msg = str(exc).lower()
    return "connection refused" in msg or "max retries exceeded" in msg or "invalid session" in msg


def main():
    """Main orchestration function"""
    parser = argparse.ArgumentParser(description="Send sponsor emails via FreeScout.")
    parser.add_argument("excel_path", nargs="?", default=None, help="Path to round CSV (or Excel)")
    parser.add_argument("--start-from", type=int, default=1, metavar="N", help="Start from sponsor index N (1-based). Use to resume after browser crash.")
    args = parser.parse_args()

    # Check configuration
    if not FREESCOUT_URL or not FREESCOUT_EMAIL or not FREESCOUT_PASSWORD:
        print("Error: FreeScout configuration missing.")
        print("Please create a .env file with FREESCOUT_URL, FREESCOUT_EMAIL, and FREESCOUT_PASSWORD")
        sys.exit(1)
    if FILTER_BY_OUTREACH and not (OUTREACH_FILTER_VALUE or "").strip():
        print("Error: FILTER_BY_OUTREACH is True but OUTREACH_FILTER_VALUE is empty.")
        print("Set OUTREACH_FILTER_VALUE in .env or config.py to match Assigned Team Member in your sheet.")
        sys.exit(1)

    excel_path = args.excel_path or input("Enter path to Excel/CSV file: ").strip()
    if not excel_path:
        print("Error: No file path provided.")
        sys.exit(1)
    if not os.path.exists(excel_path):
        print(f"Error: File not found: {excel_path}")
        sys.exit(1)
    
    # Read sponsor data
    if FILTER_BY_OUTREACH:
        print(f"Filter: Only processing sponsors where 'Assigned Team Member' = '{OUTREACH_FILTER_VALUE}'")
    reader = SponsorReader(excel_path)
    try:
        reader.read_file()
        # Round CSV (from extract_round_leads.py) has "Template Name" column; skip sponsor type lookup
        if TEMPLATE_NAME_COLUMN not in reader.df.columns:
            sponsor_type_col = reader.identify_sponsor_type_column()
            if not sponsor_type_col:
                print("Warning: Could not automatically identify sponsor type column.")
                print("Available columns:", list(reader.df.columns))
                sponsor_type_col = input("Enter the column name containing sponsor types: ").strip()
                reader.sponsor_type_column = sponsor_type_col
        sponsors = reader.get_sponsors()
        start_from = max(1, int(args.start_from))
        if start_from > 1:
            sponsors = sponsors[start_from - 1:]
            print(f"Resuming from index {start_from}: {len(sponsors)} sponsors to process.")
        else:
            print(f"Found {len(sponsors)} sponsors to process.")
        if len(sponsors) == 0:
            print("No sponsors found. Exiting.")
            sys.exit(0)
    except Exception as e:
        print(f"Error reading sponsor data: {e}")
        sys.exit(1)
    
    # Confirm before starting
    print(f"\nReady to send emails to {len(sponsors)} sponsors.")
    print("This will open a browser window and automate the email sending process.")
    confirmation = input("Continue? (y/n): ").strip().lower()
    if confirmation != 'y':
        print("Cancelled.")
        sys.exit(0)
    
    # Initialize automation
    automation = FreeScoutAutomation()
    
    # Setup log file for tracking sent emails
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_filename = f"sent_emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_filepath = os.path.join(log_dir, log_filename)
    
    def parse_emails(email_str):
        """Parse multiple emails from a string. Strips leading 'CC:', 'BCC:', 'To:' from each part."""
        if not email_str:
            return []
        if ',' in email_str:
            parts = [e.strip() for e in email_str.split(',')]
        elif ';' in email_str:
            parts = [e.strip() for e in email_str.split(';')]
        else:
            parts = [email_str.strip()]
        result = []
        for part in parts:
            if not part or '@' not in part:
                continue
            for prefix in ('CC:', 'BCC:', 'To:', 'Cc:', 'Bcc:'):
                if part.upper().startswith(prefix.upper()):
                    part = part[len(prefix):].strip()
                    break
            if part and '@' in part:
                result.append(part)
        return result
    
    def log_sent_email(email_addresses, company_name, template_name, status="SUCCESS"):
        """Log sent email to file"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        emails_str = ', '.join(email_addresses)
        with open(log_filepath, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp} | {status} | {emails_str} | {company_name} | {template_name}\n")
    
    print(f"\n📝 Log file: {log_filepath}")
    
    try:
        automation.setup_browser()
        automation.login()
        # Page should already be loaded, no need to wait
        
        # Process each sponsor
        success_count = 0
        failure_count = 0
        skipped_count = 0
        
        for i, sponsor in enumerate(sponsors, 1):
            # Parse emails for display
            email_str = sponsor['email']
            email_addresses = parse_emails(email_str)
            emails_display = ', '.join(email_addresses) if email_addresses else email_str
            
            print(f"\n{'='*80}")
            print(f"[{i}/{len(sponsors)}] Processing: {emails_display} ({sponsor['company_name']})")
            print(f"Template: {sponsor['template_name']}")
            print(f"{'='*80}")
            
            try:
                result = automation.send_sponsor_email(sponsor, confirm_before_send=True)
                if result:
                    success_count += 1
                    # Log successful send
                    log_sent_email(email_addresses, sponsor['company_name'], sponsor['template_name'], "SUCCESS")
                else:
                    skipped_count += 1
                    # Log skipped
                    log_sent_email(email_addresses, sponsor['company_name'], sponsor['template_name'], "SKIPPED")
            except KeyboardInterrupt:
                print("\n\nProcess interrupted by user.")
                break
            except Exception as e:
                print(f"Error: {e}")
                failure_count += 1
                log_sent_email(email_addresses, sponsor['company_name'], sponsor['template_name'], f"FAILED: {str(e)[:100]}")
                if _is_browser_session_lost(e):
                    next_index = (args.start_from - 1) + i + 1
                    print("\n*** Browser session lost (Chrome may have closed or crashed). ***")
                    print("Restart the script and use --start-from to resume from the next sponsor.")
                    print(f"  Example: python main.py {excel_path!r} --start-from {next_index}")
                    break

            # No delay needed between emails - process immediately
        
        # Summary
        print("\n" + "="*80)
        print("SUMMARY")
        print("="*80)
        print(f"Total sponsors: {len(sponsors)}")
        print(f"Successfully sent: {success_count}")
        print(f"Skipped/Cancelled: {skipped_count}")
        print(f"Failed: {failure_count}")
        print("="*80)
        
        # Log template usage summary
        print("\n📊 Template Usage Summary:")
        template_counts = {}
        for sponsor in sponsors:
            template = sponsor.get('template_name', 'Unknown')
            template_counts[template] = template_counts.get(template, 0) + 1
        for template, count in sorted(template_counts.items()):
            print(f"  {template}: {count}")
        print("="*80)
        print(f"\n📝 Log file saved to: {log_filepath}")
        
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("\nClosing browser...")
        automation.close()
        print("Done.")


if __name__ == "__main__":
    import time
    main()

