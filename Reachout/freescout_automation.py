"""
Browser automation module for FreeScout email sending
"""
import time
import re
import html
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from typing import Dict, Optional, Tuple
from config import FREESCOUT_URL, FREESCOUT_EMAIL, FREESCOUT_PASSWORD, HEADLESS_MODE, BROWSER_WAIT_TIME


class FreeScoutAutomation:
    """Handles browser automation for FreeScout email sending"""
    
    def __init__(self):
        """Initialize the automation with browser setup"""
        self.driver = None
        self.wait = None
        
    def setup_browser(self):
        """Setup and initialize the browser"""
        chrome_options = Options()
        if HEADLESS_MODE:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, BROWSER_WAIT_TIME)
        
    def login(self):
        """Login to FreeScout"""
        if not FREESCOUT_URL:
            raise ValueError("FREESCOUT_URL not configured. Please set it in .env file.")
        
        self.driver.get(FREESCOUT_URL)
        time.sleep(2)  # Wait for page load
        
        # Wait for login form and fill credentials
        try:
            # Try to find email/username field (common selectors)
            email_selectors = [
                "input[type='email']",
                "input[name='email']",
                "input[id='email']",
                "input[type='text']",
                "#email",
                ".email-input"
            ]
            
            email_field = None
            for selector in email_selectors:
                try:
                    email_field = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    break
                except:
                    continue
            
            if not email_field:
                raise Exception("Could not find email field. Please check FreeScout login page structure.")
            
            email_field.clear()
            email_field.send_keys(FREESCOUT_EMAIL)
            time.sleep(0.5)
            
            # Find password field
            password_selectors = [
                "input[type='password']",
                "input[name='password']",
                "#password",
                ".password-input"
            ]
            
            password_field = None
            for selector in password_selectors:
                try:
                    password_field = self.driver.find_element(By.CSS_SELECTOR, selector)
                    break
                except:
                    continue
            
            if not password_field:
                raise Exception("Could not find password field.")
            
            password_field.clear()
            password_field.send_keys(FREESCOUT_PASSWORD)
            time.sleep(0.5)
            
            # Find and click login button
            login_selectors = [
                "button[type='submit']",
                "input[type='submit']",
                "button:contains('Login')",
                "button:contains('Sign in')",
                ".login-button",
                "#login-button"
            ]
            
            login_button = None
            for selector in login_selectors:
                try:
                    if "contains" in selector:
                        # XPath for text contains
                        login_button = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Login') or contains(text(), 'Sign in')]")
                    else:
                        login_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    break
                except:
                    continue
            
            if not login_button:
                raise Exception("Could not find login button.")
            
            login_button.click()
            # Wait for page to load after login (use WebDriverWait instead of fixed sleep)
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
            except:
                pass  # Page is likely already loaded
            
            # Verify login success (check if we're redirected away from login page)
            if "login" in self.driver.current_url.lower():
                raise Exception("Login may have failed. Please check credentials.")
            
            print("Successfully logged in to FreeScout")
            
        except Exception as e:
            raise Exception(f"Login failed: {e}")
    
    def click_new_conversation(self):
        """Click the mail icon/new conversation button"""
        try:
            print("Looking for 'New Conversation' button...")
            print(f"Current URL: {self.driver.current_url}")
            
            # Wait for page to fully load after login (use WebDriverWait)
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
            except:
                pass  # Page is likely already loaded
            
            # Common selectors for new conversation/mail icon
            # Priority: Most specific FreeScout selectors first
            new_conversation_selectors = [
                "a[href*='new-ticket']",  # FreeScout specific: href contains new-ticket
                "a[aria-label='New Conversation']",  # FreeScout specific: exact aria-label
                "a.btn.btn-trans[aria-label='New Conversation']",  # FreeScout specific: combined
                "a[href*='/new-ticket']",  # Alternative path format
                "button[aria-label*='New']",
                "button[title*='New']",
                ".new-conversation",
                "#new-conversation",
                "a[href*='new']",
                "button.new-mail",
                ".mail-icon",
                "[data-action='new-conversation']",
                "button[data-testid*='new']",
                "a[data-testid*='new']",
                "button.btn-new",
                ".btn-new-conversation"
            ]
            
            new_button = None
            for selector in new_conversation_selectors:
                try:
                    print(f"  Trying selector: {selector[:60]}...")
                    if selector.startswith("//"):
                        # XPath selector
                        new_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        # CSS selector
                        new_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    if new_button:
                        print(f"  ✓ Found button with selector: {selector[:60]}")
                        break
                except Exception as e:
                    print(f"  ✗ Not found: {str(e)[:60]}")
                    continue
            
            if not new_button:
                # Try finding by icon (FreeScout uses glyphicon-envelope)
                print("  Trying to find by icon (glyphicon-envelope)...")
                try:
                    # Find the icon and get its parent <a> tag
                    icon = self.driver.find_element(By.CSS_SELECTOR, "i.glyphicon.glyphicon-envelope")
                    new_button = icon.find_element(By.XPATH, "./ancestor::a")
                    print("  ✓ Found by icon")
                except:
                    print("  ✗ Not found by icon")
                    pass
            
            if not new_button:
                print("\n⚠️  Could not find new conversation button automatically.")
                print("Please help identify the button:")
                print("1. Look at the FreeScout page in the browser")
                print("2. Find the button/link to create a new conversation")
                print("3. Right-click it and select 'Inspect'")
                print("4. Look for attributes like id, class, or data-* attributes")
                print("\nCurrent page HTML (first 2000 chars):")
                print(self.driver.page_source[:2000])
                print("\n" + "="*80)
                input("Press Enter after you've inspected the page (or Ctrl+C to cancel)...")
                
                # Try a few more common patterns
                manual_selectors = [
                    "//button[contains(@class, 'new')]",
                    "//a[contains(@class, 'new')]",
                    "//*[@id='new-conversation']",
                    "//*[contains(@class, 'compose')]",
                    "//*[contains(@aria-label, 'new') or contains(@aria-label, 'New')]"
                ]
                
                for selector in manual_selectors:
                    try:
                        new_button = self.driver.find_element(By.XPATH, selector)
                        print(f"✓ Found with manual selector: {selector}")
                        break
                    except:
                        continue
                
                if not new_button:
                    raise Exception("Could not find new conversation button. Please check FreeScout interface or update selectors.")
            
            print("Clicking new conversation button...")
            new_button.click()
            # Wait for compose window to open (use WebDriverWait instead of fixed sleep)
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field, #subject, div.note-editable")))
            except:
                pass  # Form is likely already loaded
            print("✓ New conversation window should be open")
            
        except Exception as e:
            raise Exception(f"Failed to click new conversation: {e}")
    
    def fill_to_field(self, email: str):
        """
        Fill the 'To' field with email address(es) (Select2 dropdown)
        Supports multiple emails separated by comma, semicolon, or space
        """
        import re
        
        try:
            # Parse multiple emails from the email string
            # Split by comma, semicolon, or whitespace (but preserve email format)
            # Common patterns: "email1@example.com, email2@example.com" or "email1@example.com; email2@example.com"
            email_string = str(email).strip()
            
            # Split by comma or semicolon first (most common)
            if ',' in email_string:
                emails = [e.strip() for e in email_string.split(',')]
            elif ';' in email_string:
                emails = [e.strip() for e in email_string.split(';')]
            else:
                # Try splitting by whitespace, but be careful not to split email addresses
                # Only split if there are multiple @ symbols (indicating multiple emails)
                if email_string.count('@') > 1:
                    # Split by whitespace, but keep email addresses together
                    emails = re.split(r'\s+', email_string)
                    # Filter out empty strings and validate basic email format
                    emails = [e.strip() for e in emails if e.strip() and '@' in e]
                else:
                    # Single email
                    emails = [email_string]
            
            # Clean and validate emails
            valid_emails = []
            for e in emails:
                e = e.strip()
                if e and '@' in e:  # Basic validation
                    valid_emails.append(e)
            
            if not valid_emails:
                raise Exception(f"No valid email addresses found in: {email_string}")
            
            print(f"Found {len(valid_emails)} email address(es): {', '.join(valid_emails)}")
            
            # Find Select2 container and input field (only need to do this once)
            print("Looking for 'To' field (Select2)...")
            
            # FreeScout uses Select2 for the To field
            # First, find the Select2 container
            select2_selectors = [
                ".select2-container",
                "span.select2-container",
                ".select2-selection",
                "span.select2-selection"
            ]
            
            select2_container = None
            for selector in select2_selectors:
                try:
                    select2_container = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if select2_container:
                        print(f"  ✓ Found Select2 container")
                        break
                except:
                    continue
            
            # Find the search input field
            to_field = None
            to_selectors = [
                "input.select2-search__field",  # FreeScout specific
                ".select2-search__field",
                "input.select2-search input",
                "input[role='textbox'][aria-autocomplete='list']"
            ]
            
            for selector in to_selectors:
                try:
                    to_field = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if to_field:
                        print(f"  ✓ Found To field")
                        break
                except:
                    continue
            
            if not to_field:
                raise Exception("Could not find 'To' field. Please check FreeScout interface.")
            
            # Add each email address
            for i, email_addr in enumerate(valid_emails, 1):
                print(f"  Adding email {i}/{len(valid_emails)}: {email_addr}")
                
                # For each email (including first), click Select2 container to open/focus the field
                # This is needed because after adding an email, Select2 closes and we need to reopen it
                # Re-find the container for each email in case DOM changed
                select2_container_current = None
                for selector in select2_selectors:
                    try:
                        select2_container_current = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                        if select2_container_current:
                            break
                    except:
                        continue
                
                if select2_container_current:
                    try:
                        select2_container_current.click()
                        time.sleep(0.2)  # Small wait for dropdown to open
                    except:
                        pass
                
                # Find the input field again (it might have been recreated after adding previous email)
                to_field = None
                for selector in to_selectors:
                    try:
                        to_field = WebDriverWait(self.driver, 1).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                        if to_field:
                            break
                    except:
                        continue
                
                if not to_field:
                    raise Exception("Could not find 'To' field after opening Select2.")
                
                # Type the email (don't clear - the field should be empty when opened)
                to_field.send_keys(email_addr)
                
                # Wait for Select2 to process
                try:
                    WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option, .select2-selection__choice"))
                    )
                except:
                    pass  # Continue anyway
                
                # Try to select from dropdown or press Enter
                try:
                    option = WebDriverWait(self.driver, 1).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-results__option--highlighted, .select2-results__option"))
                    )
                    option.click()
                    print(f"    ✓ Selected {email_addr} from dropdown")
                except:
                    # If no dropdown option, just press Enter to confirm
                    to_field.send_keys(Keys.RETURN)
                    print(f"    ✓ Confirmed {email_addr} with Enter")
                
                # Small wait between emails to ensure Select2 processes each one
                if i < len(valid_emails):
                    time.sleep(0.3)
            
            # After adding all emails, ensure Select2 dropdown is closed
            # Press Escape to close any open Select2 dropdown
            try:
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.2)
            except:
                pass
            
            print(f"✓ To field filled with {len(valid_emails)} email address(es)")
            
        except Exception as e:
            raise Exception(f"Failed to fill 'To' field: {e}")
    
    def add_tag(self, tag_name: str):
        """Add a tag to the conversation"""
        try:
            print(f"Adding tag: {tag_name}...")
            
            # Find and click the tag icon
            tag_icon_selectors = [
                "i.conv-new-tag.conv-add-tags",
                "i.glyphicon-tag[data-toggle='dropdown']",
                ".conv-new-tag",
                "i[class*='conv-add-tags']"
            ]
            
            tag_icon = None
            for selector in tag_icon_selectors:
                try:
                    tag_icon = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if tag_icon:
                        print(f"  ✓ Found tag icon")
                        break
                except:
                    continue
            
            if not tag_icon:
                print("  ⚠️  Could not find tag icon, skipping tag addition")
                return
            
            # Click the tag icon to open dropdown
            tag_icon.click()
            time.sleep(0.3)  # Wait for dropdown to open
            
            # Find the tag dropdown container first to scope all searches
            tag_dropdown = None
            try:
                tag_dropdown = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#add-tag-wrap"))
                )
                print(f"  ✓ Found tag dropdown container")
            except:
                print("  ⚠️  Could not find tag dropdown container, skipping tag addition")
                return
            
            # Wait a bit for Select2 to initialize
            time.sleep(0.5)
            
            # Try to find and click the Select2 container or selection area to open/focus
            # Try multiple approaches to open the Select2
            select2_opened = False
            try:
                # Approach 1: Find Select2 container within tag dropdown
                tag_select2_container = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-container, span.select2-container")
                tag_select2_container.click()
                print(f"  ✓ Clicked Select2 container")
                select2_opened = True
                time.sleep(0.3)
            except:
                try:
                    # Approach 2: Find Select2 selection area and click it
                    tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                    tag_select2_selection.click()
                    print(f"  ✓ Clicked Select2 selection area")
                    select2_opened = True
                    time.sleep(0.3)
                except:
                    try:
                        # Approach 3: Find the hidden select element and trigger Select2
                        hidden_select = tag_dropdown.find_element(By.CSS_SELECTOR, "select.tag-input")
                        self.driver.execute_script("arguments[0].dispatchEvent(new Event('select2:open'));", hidden_select)
                        print(f"  ✓ Triggered Select2 open via JavaScript")
                        select2_opened = True
                        time.sleep(0.3)
                    except:
                        pass
            
            # Find the Select2 input field - MUST be within tag dropdown only
            # Try multiple selectors and wait a bit longer
            tag_input = None
            input_selectors = [
                "#add-tag-wrap input.select2-search__field",
                "#add-tag-wrap .select2-search__field",
                ".conv-new-tag-dd input.select2-search__field",
                "input.select2-search__field"
            ]
            
            for selector in input_selectors:
                try:
                    tag_input = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    # Verify it's actually in the tag dropdown
                    try:
                        parent = tag_input.find_element(By.XPATH, "./ancestor::*[@id='add-tag-wrap' or contains(@class, 'conv-new-tag-dd')]")
                        print(f"  ✓ Found tag input field (scoped to tag dropdown)")
                        break
                    except:
                        # If not in tag dropdown, continue to next selector
                        tag_input = None
                        continue
                except:
                    continue
            
            if not tag_input:
                print("  ⚠️  Could not find tag input field in dropdown, skipping tag addition")
                return
            
            # Click the input to ensure it's focused
            try:
                tag_input.click()
                time.sleep(0.2)
            except:
                pass
            
            # Type the tag name
            tag_input.clear()
            tag_input.send_keys(tag_name)
            time.sleep(0.5)  # Wait for Select2 to process and show results
            
            # Wait for Select2 results to appear
            try:
                WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option"))
                )
            except:
                print("  ⚠️  Select2 results did not appear, but continuing...")
            
            # Try to select from dropdown or press Enter
            # IMPORTANT: Only look for options that contain our tag name to avoid matching "To" field
            tag_selected_in_select2 = False
            try:
                # Look for dropdown option that contains the tag name - scoped to tag dropdown context
                # Select2 results appear outside the dropdown, but we can verify by text content
                options = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".select2-results__option"))
                )
                
                # Find the option that contains our tag name
                option = None
                for opt in options:
                    if tag_name.lower() in opt.text.lower():
                        option = opt
                        break
                
                if option:
                    # Scroll option into view and click
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option)
                    time.sleep(0.2)
                    option.click()
                    print(f"  ✓ Selected {tag_name} from dropdown")
                    time.sleep(0.5)  # Wait for selection to be processed
                else:
                    raise Exception("Tag option not found in dropdown")
                
                # Verify tag was selected in Select2 - check within tag dropdown's Select2 only
                try:
                    # Find the Select2 selection within the tag dropdown
                    tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                    tag_chip = tag_select2_selection.find_element(By.CSS_SELECTOR, ".select2-selection__choice")
                    if tag_name in tag_chip.get_attribute('title') or tag_name in tag_chip.text:
                        tag_selected_in_select2 = True
                        print(f"  ✓ Tag '{tag_name}' confirmed in Select2 selection")
                except:
                    # Alternative: check if tag name appears in selection text
                    try:
                        tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                        if tag_name in tag_select2_selection.text:
                            tag_selected_in_select2 = True
                            print(f"  ✓ Tag '{tag_name}' confirmed in Select2 (by text)")
                    except:
                        pass
            except Exception as e:
                # If no dropdown option found, just press Enter
                print(f"  ℹ️  No matching dropdown option found, using Enter key")
                tag_input.send_keys(Keys.RETURN)
                print(f"  ✓ Confirmed {tag_name} with Enter")
                time.sleep(0.5)  # Wait for Enter to be processed
                
                # Check if tag was selected - scoped to tag dropdown
                try:
                    tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                    tag_chip = tag_select2_selection.find_element(By.CSS_SELECTOR, ".select2-selection__choice")
                    if tag_name in tag_chip.get_attribute('title') or tag_name in tag_chip.text:
                        tag_selected_in_select2 = True
                except:
                    pass
            
            if not tag_selected_in_select2:
                print(f"  ⚠️  Warning: Tag may not be selected in Select2 before clicking OK")
            
            time.sleep(0.3)
            
            # Find and click the tick/ok button - MUST be within tag dropdown
            ok_button = None
            
            # Try multiple approaches to find the button (prioritize direct selector)
            try:
                # Approach 1: Find button directly (most reliable)
                ok_button = WebDriverWait(self.driver, 2).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#add-tag-wrap button.btn-default"))
                )
                print(f"  ✓ Found OK button directly")
            except:
                try:
                    # Approach 2: Find by icon and get parent button
                    icon = WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#add-tag-wrap i.glyphicon-ok"))
                    )
                    ok_button = icon.find_element(By.XPATH, "./parent::button")
                    print(f"  ✓ Found OK button via icon")
                except:
                    try:
                        # Approach 3: Find any button in the dropdown
                        ok_button = self.driver.find_element(By.CSS_SELECTOR, ".conv-new-tag-dd button.btn-default")
                        print(f"  ✓ Found OK button via dropdown")
                    except:
                        pass
            
            if ok_button:
                # Scroll into view
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", ok_button)
                    time.sleep(0.3)
                except:
                    pass
                
                # Try regular click first, then JavaScript as fallback
                try:
                    ok_button.click()
                    print(f"  ✓ Clicked OK button")
                    time.sleep(0.5)  # Wait for tag to be saved
                except Exception as e:
                    # Fallback to JavaScript click
                    try:
                        self.driver.execute_script("arguments[0].click();", ok_button)
                        print(f"  ✓ Clicked OK button via JavaScript")
                        time.sleep(0.5)
                    except Exception as js_e:
                        print(f"  ⚠️  Could not click OK button: {js_e}")
                
                # Verify tag was actually added by checking if it appears in the conversation tags
                time.sleep(0.5)  # Wait for tag to be saved
                tag_added = False
                try:
                    # Check if tag appears in the conversation tags area
                    # Tags usually appear as spans or badges with the tag name
                    tag_selectors = [
                        ".conv-tag",
                        ".tag-item",
                        ".tag",
                        f"[data-tag='{tag_name}']",
                        f"span[title='{tag_name}']",
                        ".conv-tags .tag",
                        ".conv-tags span"
                    ]
                    tag_elements = []
                    for selector in tag_selectors:
                        try:
                            elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                            tag_elements.extend(elements)
                        except:
                            pass
                    
                    for elem in tag_elements:
                        elem_text = elem.text.strip()
                        elem_title = elem.get_attribute('title') or ''
                        elem_data_tag = elem.get_attribute('data-tag') or ''
                        if (tag_name.lower() in elem_text.lower() or 
                            tag_name.lower() in elem_title.lower() or 
                            tag_name.lower() in elem_data_tag.lower()):
                            tag_added = True
                            print(f"  ✓ Tag '{tag_name}' confirmed in conversation tags")
                            break
                    
                    if not tag_added:
                        # Also check if dropdown closed (alternative verification)
                        try:
                            dropdown = self.driver.find_element(By.CSS_SELECTOR, "#add-tag-wrap")
                            if not dropdown.is_displayed() or "display: none" in (dropdown.get_attribute("style") or ""):
                                print(f"  ✓ Dropdown closed - tag likely saved")
                                tag_added = True
                        except:
                            pass
                    
                    if not tag_added:
                        print(f"  ⚠️  Warning: Could not verify tag was added. It may not have been saved.")
                except Exception as verify_e:
                    print(f"  ⚠️  Could not verify tag: {verify_e}")
            else:
                print("  ⚠️  Could not find OK button, tag may not be saved")
            
            if tag_added:
                print(f"✓ Tag '{tag_name}' added successfully")
            else:
                print(f"⚠️  Tag '{tag_name}' may not have been added - please verify manually")
            
        except Exception as e:
            print(f"  ⚠️  Warning: Could not add tag: {e}")
            import traceback
            traceback.print_exc()
            # Don't raise exception - tag addition is not critical
    
    def select_template(self, template_name: str):
        """Select template from dropdown (FreeScout Saved Replies)"""
        try:
            print(f"Looking for template selector (searching for: '{template_name}')...")
            
            # FreeScout uses a button with dropdown for Saved Replies
            template_button_selectors = [
                "button[aria-label='Saved Replies']",  # FreeScout specific
                "button.dropdown-toggle[data-toggle='dropdown']",
                "button.note-btn.dropdown-toggle",
                "button[title='Saved Replies']",
                ".dropdown-toggle[data-toggle='dropdown']"
            ]
            
            template_button = None
            for selector in template_button_selectors:
                try:
                    print(f"  Trying template button selector: {selector[:50]}...")
                    template_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if template_button:
                        print(f"  ✓ Found template button with: {selector[:50]}")
                        break
                except Exception as e:
                    print(f"  ✗ Not found: {str(e)[:50]}")
                    continue
            
            if not template_button:
                print("\n⚠️  Could not find template button automatically.")
                raise Exception("Could not find template button.")
            
            # Click the button to open the dropdown
            print("  Clicking template button to open dropdown...")
            template_button.click()
            # Wait for dropdown to open - use shorter timeout and visibility check
            try:
                WebDriverWait(self.driver, 1).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".dropdown-menu.dropdown-saved-replies")))
            except:
                # If visibility check fails, try presence with small wait
                try:
                    WebDriverWait(self.driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".dropdown-menu.dropdown-saved-replies")))
                    time.sleep(0.1)  # Tiny wait for dropdown animation
                except:
                    pass  # Dropdown might already be open
            
            template_link_selectors = [
                f"//div[@class='dropdown-menu']//a[contains(text(), '{template_name}')]",  # XPath with exact match
                f"//li/a[contains(text(), '{template_name}')]",  # Alternative XPath
                f"a[data-value]:contains('{template_name}')"  # CSS (won't work, but keeping for reference)
            ]
            
            template_link = None
            # Use shorter timeout for template link selection (2 seconds instead of 10)
            short_wait = WebDriverWait(self.driver, 2)
            for selector in template_link_selectors:
                try:
                    print(f"  Trying template link selector: {selector[:60]}...")
                    if selector.startswith("//"):
                        template_link = short_wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        template_link = short_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    if template_link:
                        print(f"  ✓ Found template link with: {selector[:60]}")
                        break
                except Exception as e:
                    # Don't print error for first selector - it's expected to try multiple
                    if selector != template_link_selectors[0]:
                        print(f"  ✗ Not found: {str(e)[:60]}")
                    continue
            
            if not template_link:
                # Try a more flexible search - look for any link containing the template name
                try:
                    print(f"  Trying flexible search for template: '{template_name}'...")
                    # Search in the dropdown menu
                    dropdown = self.driver.find_element(By.CSS_SELECTOR, ".dropdown-menu.dropdown-saved-replies")
                    template_link = dropdown.find_element(By.XPATH, f".//a[contains(text(), '{template_name}')]")
                    print("  ✓ Found template link with flexible search")
                except Exception as e:
                    print(f"  ✗ Flexible search failed: {str(e)[:60]}")
                    raise Exception(f"Could not find template '{template_name}' in dropdown.")
            
            # Ensure template link is visible and clickable
            print(f"  Clicking template: '{template_name}'...")
            # Wait for element to be clickable
            try:
                template_link = WebDriverWait(self.driver, 2).until(
                    EC.element_to_be_clickable(template_link)
                )
            except:
                # If not clickable, try scrolling into view
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", template_link)
                    time.sleep(0.3)
                except:
                    pass
            
            # Try clicking - use JavaScript click as it's more reliable for dropdown items
            try:
                # First try regular click
                template_link.click()
            except Exception as e:
                # If regular click fails, use JavaScript click (more reliable for dropdown items)
                print(f"  Regular click failed ({str(e)[:50]}), using JavaScript click...")
                try:
                    self.driver.execute_script("arguments[0].click();", template_link)
                except Exception as js_e:
                    raise Exception(f"Could not click template link. Regular click: {e}. JavaScript click: {js_e}")
            # Wait for template to load into editor (use WebDriverWait instead of fixed sleep)
            body_selectors = [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable"
            ]
            for selector in body_selectors:
                try:
                    body_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    # Wait for content to appear (not just empty) - use WebDriverWait
                    try:
                        WebDriverWait(self.driver, 2).until(
                            lambda d: body_element.text and len(body_element.text.strip()) > 50
                        )
                        body_text = body_element.text if hasattr(body_element, 'text') else body_element.get_attribute('textContent') or ''
                        print(f"  ✓ Template content loaded ({len(body_text)} chars)")
                    except:
                        print("  ⚠️  Warning: Template content may not have loaded fully")
                    break
                except:
                    continue
            
            print(f"✓ Template selected and loaded: '{template_name}'")
            
        except Exception as e:
            raise Exception(f"Failed to select template '{template_name}': {e}")
    
    def extract_template_content(self) -> Tuple[str, str]:
        """
        Extract template content and subject
        
        Returns:
            Tuple of (subject, body) - subject may be empty if not found
        """
        try:
            print("Extracting template content from body...")
            # Find the email body/editor - FreeScout uses contenteditable div
            body_selectors = [
                "div.note-editable[contenteditable='true']",  # FreeScout specific
                "div[contenteditable='true']",
                ".note-editable",
                "textarea[name*='body']",
                "textarea[id*='body']",
                ".email-body",
                "#email-body",
                "iframe[title*='editor']",
                ".editor-body"
            ]
            
            body_element = None
            for selector in body_selectors:
                try:
                    print(f"  Trying body selector: {selector[:50]}...")
                    body_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if body_element:
                        print(f"  ✓ Found body editor with: {selector[:50]}")
                        break
                except:
                    continue
            
            if not body_element:
                raise Exception("Could not find email body editor.")
            
            # Get content - handle iframe or contenteditable div
            if body_element.tag_name == 'iframe':
                self.driver.switch_to.frame(body_element)
                body_text = self.driver.find_element(By.TAG_NAME, "body").text
                self.driver.switch_to.default_content()
            else:
                # For contenteditable div, get HTML to preserve formatting and links
                # Simple approach: just get innerHTML - it should already be clean content
                body_html = body_element.get_attribute('innerHTML') or ""
                
                # Also get text version for subject extraction
                body_text = body_element.text if hasattr(body_element, 'text') else body_element.get_attribute('innerText') or body_element.get_attribute('textContent') or ""
            
            print(f"  Raw template HTML length: {len(body_html)} chars")
            print(f"  Raw template text length: {len(body_text)} chars")
            # Debug: Check if placeholder exists in HTML
            if "[Prospective Sponsor's Name]" in body_html or "[prospective sponsor's name]" in body_html.lower():
                print(f"  ✓ Found '[Prospective Sponsor's Name]' in template HTML")
                # Show snippet
                idx = body_html.lower().find("[prospective sponsor's name]")
                if idx >= 0:
                    snippet = body_html[max(0, idx-30):idx+80]
                    print(f"  Snippet: ...{snippet}...")
            
            # Extract subject line from text version (line starting with "Subject -" or "Subject:")
            subject = ""
            body_lines = body_text.split('\n')
            subject_found = False
            
            for line in body_lines:
                line_stripped = line.strip()
                # Check for subject line patterns (case insensitive)
                if not subject_found and (line_stripped.lower().startswith("subject -") or line_stripped.lower().startswith("subject:")):
                    # Extract subject
                    if "Subject -" in line_stripped or "subject -" in line_stripped:
                        subject = line_stripped.split("Subject -", 1)[-1].split("subject -", 1)[-1].strip()
                    elif "Subject:" in line_stripped or "subject:" in line_stripped:
                        subject = line_stripped.split("Subject:", 1)[-1].split("subject:", 1)[-1].strip()
                    print(f"  Found subject: '{subject}'")
                    subject_found = True
                    break
            
            # Remove subject line from HTML if found
            if subject_found:
                # Remove subject line from HTML (handle various formats)
                # First, try to remove it as a complete HTML element
                body_html = re.sub(r'<[^>]*>Subject\s*[-:]\s*[^<]*</[^>]*>', '', body_html, flags=re.IGNORECASE)
                # Remove subject line that might span multiple tags
                body_html = re.sub(r'<[^>]*>Subject\s*[-:]\s*[^<]*', '', body_html, flags=re.IGNORECASE)
                # Remove subject line from text content (not in tags)
                body_html = re.sub(r'Subject\s*[-:]\s*[^<\n\r]*', '', body_html, flags=re.IGNORECASE)
                # Also remove any leading/trailing whitespace and newlines after removal
                body_html = re.sub(r'^(\s|\n|\r|<br\s*/?>)+', '', body_html, flags=re.IGNORECASE | re.MULTILINE)
            
            # Clean up HTML: remove only truly empty/duplicate elements, preserve ALL formatting including line breaks
            # IMPORTANT: Don't remove <div><br></div> or <span><br></span> - these create line breaks!
            # Only remove empty divs with nested empty spans (like <div><span></span></div> with no content)
            body_html = re.sub(r'<div[^>]*>\s*<span[^>]*>\s*</span>\s*</div>', '', body_html, flags=re.IGNORECASE)
            # Remove completely empty divs (no content, no <br>, no spans) - but preserve <div><br></div>
            body_html = re.sub(r'<div[^>]*>\s*</div>', '', body_html, flags=re.IGNORECASE)
            # DON'T remove trailing <br> tags - they might be intentional line breaks
            # DON'T reduce <br> tags - preserve the structure as-is
            # Remove only completely empty paragraphs (no content, no <br>)
            body_html = re.sub(r'<p[^>]*>\s*</p>', '', body_html, flags=re.IGNORECASE)
            # Preserve ALL whitespace and newlines - only trim leading/trailing whitespace from entire string
            body_html = body_html.strip()
            
            print(f"  Extracted subject: '{subject}'")
            print(f"  Body HTML length after cleanup: {len(body_html)} chars")
            
            return subject, body_html  # Return HTML to preserve formatting
            
        except Exception as e:
            raise Exception(f"Failed to extract template content: {e}")
    
    def fill_placeholders(self, body: str, sponsor_data: Dict) -> str:
        """
        Fill placeholders in template body (works with both text and HTML)
        
        Args:
            body: Template body text or HTML
            sponsor_data: Dictionary with sponsor information
            
        Returns:
            Body text/HTML with placeholders filled
        """
        from config import PLACEHOLDER_MAPPINGS
        
        filled_body = body
        contact_person = sponsor_data.get("contact_person", "")
        company_name = sponsor_data.get("company_name", "")
        
        # Extract first name from contact person (everything before first space)
        contact_first_name = contact_person.split()[0] if contact_person and contact_person.strip() else contact_person
        
        # Get product name - priority: Company Name first, then General Info
        product_name = company_name  # Use Company Name first (always prefer this)
        row_data = sponsor_data.get("row_data", {})
        # Only use General Info if Company Name is empty or not available
        if not product_name or product_name.strip() == "":
            if "General Info" in row_data:
                general_info = str(row_data["General Info"]).strip() if not pd.isna(row_data.get("General Info")) else ""
                if general_info:
                    product_name = general_info
        
        print(f"  Filling placeholders:")
        print(f"    Contact Person (full): '{contact_person}'")
        print(f"    Contact Person (first name): '{contact_first_name}'")
        print(f"    Company Name: '{company_name}'")
        
        # Check if placeholders exist in body (check both plain text and HTML)
        # The apostrophe might be HTML encoded as &#39; or &apos;
        placeholder_variations = [
            "[Prospective Sponsor's Name]",  # Plain text
            "[Prospective Sponsor&#39;s Name]",  # HTML entity &#39;
            "[Prospective Sponsor&apos;s Name]",  # HTML entity &apos;
            "[Prospective Sponsor&rsquo;s Name]",  # HTML entity &rsquo;
        ]
        
        placeholder_found = False
        for variant in placeholder_variations:
            if variant in filled_body:
                print(f"    ✓ Found placeholder variant: '{variant}' in body")
                placeholder_found = True
                break
        
        # Also check case-insensitive
        if not placeholder_found:
            for variant in placeholder_variations:
                if variant.lower() in filled_body.lower():
                    print(f"    ✓ Found placeholder variant (case variation): '{variant}'")
                    placeholder_found = True
                    break
        
        # Check with regex for split across tags
        if not placeholder_found:
            if re.search(r'\[Prospective\s+Sponsor[&#\w;]*s\s+Name\]', filled_body, re.IGNORECASE):
                print(f"    ✓ Found '[Prospective Sponsor's Name]' in body (regex match)")
                placeholder_found = True
        
        if "[Company Name]" in filled_body:
            print(f"    ✓ Found '[Company Name]' in body")
        
        # Fill placeholders in order (longest first to avoid partial replacements)
        # Handle both plain text and HTML (placeholders might be in HTML tags or as text)
        
        # Combined placeholder first
        # Note: [Prospective Sponsor's Name] should use first name only
        # Handle HTML entity encoded apostrophes: &#39;, &apos;, &rsquo;
        # Also handle Unicode curly quotes: ' (U+2019), ' (U+2018)
        placeholder_patterns = [
            ("[Customer Company POC][Prospective Sponsor's Name]", contact_person),  # Combined uses full name
            ("[Customer Company POC Name]", contact_person),  # Alternative format - uses full name
            ("[Customer Company's product name]", product_name),  # Product name - straight apostrophe (U+0027)
            ("[Customer Company's product name]", product_name),  # Product name - curly quote U+2019 (right single quotation mark)
            ("[Customer Company's product name]", product_name),  # Product name - curly quote U+2018 (left single quotation mark)
            ("[Customer Company&#39;s product name]", product_name),  # HTML entity &#39;
            ("[Customer Company&apos;s product name]", product_name),  # HTML entity &apos;
            ("[Customer Company&rsquo;s product name]", product_name),  # HTML entity &rsquo;
            ("[Customer Company Name]", company_name),  # Company name - alternative format
            ("[Prospective Sponsor's Name]", contact_first_name),  # Individual uses first name only - plain apostrophe (U+0027)
            ("[Prospective Sponsor's Name]", contact_first_name),  # Curly quote U+2019 (right single quotation mark)
            ("[Prospective Sponsor's Name]", contact_first_name),  # Curly quote U+2018 (left single quotation mark)
            ("[Prospective Sponsor&#39;s Name]", contact_first_name),  # HTML entity &#39;
            ("[Prospective Sponsor&apos;s Name]", contact_first_name),  # HTML entity &apos;
            ("[Prospective Sponsor&rsquo;s Name]", contact_first_name),  # HTML entity &rsquo;
            ("[Customer Company POC]", contact_person),  # Uses full name
            ("[Company Name]", company_name),
        ]
        
        # Also try replacing with regex that handles any apostrophe character
        # This catches Unicode variations we might have missed
        # Handle various apostrophe characters: ' (U+0027), ' (U+2019), ' (U+2018)
        apostrophe_chars = ["'", "'", "'"]  # Regular, right curly, left curly
        for apostrophe_char in apostrophe_chars:
            # Handle [Prospective Sponsor's Name]
            placeholder_with_apos = f"[Prospective Sponsor{apostrophe_char}s Name]"
            if placeholder_with_apos in filled_body:
                count = filled_body.count(placeholder_with_apos)
                filled_body = filled_body.replace(placeholder_with_apos, contact_first_name)
                print(f"    ✓ Replaced '{placeholder_with_apos}' ({count} occurrences) - Unicode apostrophe")
            
            # Handle [Customer Company's product name]
            product_placeholder = f"[Customer Company{apostrophe_char}s product name]"
            if product_placeholder in filled_body:
                count = filled_body.count(product_placeholder)
                filled_body = filled_body.replace(product_placeholder, product_name)
                print(f"    ✓ Replaced '{product_placeholder}' ({count} occurrences) - Unicode apostrophe")
        
        # Also try regex pattern for any apostrophe variation
        apostrophe_pattern = r'\[Prospective Sponsor[\'\"\u2018\u2019]s Name\]'
        if re.search(apostrophe_pattern, filled_body, re.IGNORECASE):
            matches = re.findall(apostrophe_pattern, filled_body, re.IGNORECASE)
            filled_body = re.sub(apostrophe_pattern, contact_first_name, filled_body, flags=re.IGNORECASE)
            print(f"    ✓ Replaced {len(matches)} occurrence(s) with Unicode apostrophe regex")
        
        # Also try regex pattern for [Customer Company's product name] with any apostrophe
        product_apostrophe_pattern = r'\[Customer Company[\'\"\u2018\u2019]s product name\]'
        if re.search(product_apostrophe_pattern, filled_body, re.IGNORECASE):
            matches = re.findall(product_apostrophe_pattern, filled_body, re.IGNORECASE)
            filled_body = re.sub(product_apostrophe_pattern, product_name, filled_body, flags=re.IGNORECASE)
            print(f"    ✓ Replaced {len(matches)} occurrence(s) of product name placeholder with Unicode apostrophe regex")
        
        for placeholder, replacement in placeholder_patterns:
            # Simple string replacement - replace ALL occurrences
            count_before = filled_body.count(placeholder)
            if count_before > 0:
                filled_body = filled_body.replace(placeholder, replacement)
                print(f"    ✓ Replaced '{placeholder}' ({count_before} occurrences) with '{replacement}'")
            
            # Also replace HTML entity encoded versions
            # Brackets encoded
            placeholder_html_brackets = placeholder.replace("[", "&#91;").replace("]", "&#93;")
            count_html_brackets = filled_body.count(placeholder_html_brackets)
            if count_html_brackets > 0:
                filled_body = filled_body.replace(placeholder_html_brackets, replacement)
                print(f"    ✓ Replaced HTML brackets version ({count_html_brackets} occurrences)")
            
            # Apostrophe encoded as &#39;
            placeholder_html_39 = placeholder.replace("'", "&#39;")
            count_html_39 = filled_body.count(placeholder_html_39)
            if count_html_39 > 0:
                filled_body = filled_body.replace(placeholder_html_39, replacement)
                print(f"    ✓ Replaced HTML &#39; version ({count_html_39} occurrences)")
            
            # Apostrophe encoded as &apos;
            placeholder_html_apos = placeholder.replace("'", "&apos;")
            count_html_apos = filled_body.count(placeholder_html_apos)
            if count_html_apos > 0:
                filled_body = filled_body.replace(placeholder_html_apos, replacement)
                print(f"    ✓ Replaced HTML &apos; version ({count_html_apos} occurrences)")
            
            # Both brackets and apostrophe encoded
            placeholder_html_full = placeholder.replace("[", "&#91;").replace("]", "&#93;").replace("'", "&#39;")
            count_html_full = filled_body.count(placeholder_html_full)
            if count_html_full > 0:
                filled_body = filled_body.replace(placeholder_html_full, replacement)
                print(f"    ✓ Replaced fully encoded version ({count_html_full} occurrences)")
            
            # Verify replacement worked
            count_after = filled_body.count(placeholder)
            if count_after > 0:
                print(f"      ⚠️  Warning: {count_after} occurrences of '{placeholder}' still remain!")
                idx = filled_body.find(placeholder)
                if idx >= 0:
                    snippet = filled_body[max(0, idx-50):idx+len(placeholder)+50]
                    print(f"      Snippet: ...{snippet}...")
        
        # Look for any other [placeholder] patterns and try to match from row_data
        placeholder_pattern = r'\[([^\]]+)\]'
        placeholders = re.findall(placeholder_pattern, filled_body)
        
        for placeholder in placeholders:
            if placeholder not in ["Company Name", "Customer Company POC", "Prospective Sponsor's Name", 
                                 "company name", "customer company poc", "prospective sponsor's name"]:
                # Try to find in row_data
                row_data = sponsor_data.get("row_data", {})
                if placeholder in row_data:
                    filled_body = filled_body.replace(f"[{placeholder}]", str(row_data[placeholder]))
                    print(f"    ✓ Replaced '[{placeholder}]' from row_data")
        
        # Final aggressive pass for [Prospective Sponsor's Name] - use while loop to replace ALL
        # This ensures we catch every occurrence
        # Handle all apostrophe variations: regular ', curly ' (U+2019), curly ' (U+2018)
        # Use actual Unicode characters
        regular_apos = "'"  # U+0027
        curly_right = "'"  # U+2019
        curly_left = "'"  # U+2018
        
        placeholder_final_variations = [
            f"[Prospective Sponsor{regular_apos}s Name]",  # Regular apostrophe
            f"[Prospective Sponsor{curly_right}s Name]",  # Curly quote U+2019
            f"[Prospective Sponsor{curly_left}s Name]",  # Curly quote U+2018
            "[Prospective Sponsor&#39;s Name]",
            "[Prospective Sponsor&apos;s Name]",
            "[Prospective Sponsor&rsquo;s Name]",
            "[Prospective Sponsor&#x27;s Name]",
        ]
        
        for placeholder_var in placeholder_final_variations:
            count = 0
            while placeholder_var in filled_body:
                filled_body = filled_body.replace(placeholder_var, contact_first_name)
                count += 1
            if count > 0:
                print(f"    ✓ Final pass: Replaced '{placeholder_var}' ({count} occurrences) with '{contact_first_name}'")
        
        # Also try regex for any remaining variations (spaces, different encoding, any apostrophe)
        # Match any apostrophe character: ' ' ' or HTML entities
        pattern = r'\[Prospective\s+Sponsor[\'\"\u2018\u2019&#\w;]*s\s+Name\]'
        matches = re.findall(pattern, filled_body, re.IGNORECASE)
        if matches:
            filled_body = re.sub(pattern, contact_first_name, filled_body, flags=re.IGNORECASE)
            print(f"    ✓ Final pass: Replaced {len(matches)} occurrence(s) with regex pattern (any apostrophe)")
        
        # Final verification - check for any variation
        still_found = False
        for placeholder_var in placeholder_final_variations:
            if placeholder_var in filled_body:
                still_found = True
                idx = filled_body.find(placeholder_var)
                if idx >= 0:
                    snippet = filled_body[max(0, idx-100):idx+150]
                    print(f"    ⚠️  Warning: '{placeholder_var}' still found after all replacements!")
                    print(f"    Location snippet: ...{snippet}...")
        
        # Check for any variation with "[Prospective Sponsor"
        if "[Prospective Sponsor" in filled_body:
            idx = filled_body.find("[Prospective Sponsor")
            if idx >= 0:
                end_idx = filled_body.find("]", idx)
                if end_idx > idx:
                    remaining = filled_body[idx:end_idx+1]
                    # Check if it's one we already tried
                    if remaining not in placeholder_final_variations:
                        print(f"    ⚠️  Warning: Found variation '{remaining}' that wasn't in our list!")
                        # Try to replace it anyway
                        filled_body = filled_body.replace(remaining, contact_first_name)
                        print(f"    ✓ Attempted to replace variation: '{remaining}'")
        
        # Minimal cleanup: remove only truly empty elements, preserve ALL formatting including line breaks
        # IMPORTANT: Don't remove <div><br></div> or <span><br></span> - these create line breaks!
        # Only remove empty divs with nested empty spans (like <div><span></span></div> with no content)
        filled_body = re.sub(r'<div[^>]*>\s*<span[^>]*>\s*</span>\s*</div>', '', filled_body, flags=re.IGNORECASE)
        # Remove completely empty divs (no content, no <br>, no spans) - but preserve <div><br></div>
        filled_body = re.sub(r'<div[^>]*>\s*</div>', '', filled_body, flags=re.IGNORECASE)
        # DON'T remove trailing <br> tags - they might be intentional line breaks
        # DON'T reduce <br> tags - preserve the structure as-is
        # Remove only completely empty paragraphs (no content, no <br>)
        filled_body = re.sub(r'<p[^>]*>\s*</p>', '', filled_body, flags=re.IGNORECASE)
        # Preserve ALL whitespace and line breaks - only trim trailing whitespace from entire string
        filled_body = filled_body.rstrip()
        
        # Validate HTML structure is preserved after replacement
        # Check for common HTML tags that should be preserved
        html_tags_present = bool(re.search(r'<(a|strong|em|b|i|span|p|div|h[1-6]|ul|ol|li)', filled_body, re.IGNORECASE))
        if html_tags_present:
            print(f"    ✓ HTML structure preserved (tags found)")
        else:
            print(f"    ⚠️  Warning: No HTML tags found - content may be plain text")
        
        # Check for hyperlinks specifically
        links_present = bool(re.search(r'<a\s+[^>]*href', filled_body, re.IGNORECASE))
        if links_present:
            link_count = len(re.findall(r'<a\s+[^>]*href', filled_body, re.IGNORECASE))
            print(f"    ✓ Hyperlinks preserved ({link_count} link(s) found)")
        else:
            # Check if original had links
            if '<a' in body.lower():
                print(f"    ⚠️  Warning: Original had links but they may be lost after replacement")
        
        return filled_body
    
    def fill_email_body(self, body: str, sponsor_data: Dict = None):
        """Fill the email body with content"""
        try:
            print("Filling email body...")
            
            # Get replacement values if sponsor_data is provided
            contact_first_name = ""
            company_name = ""
            contact_person = ""
            product_name = ""
            if sponsor_data:
                contact_person = sponsor_data.get("contact_person", "")
                contact_first_name = contact_person.split()[0] if contact_person and contact_person.strip() else contact_person
                company_name = sponsor_data.get("company_name", "")
                # Get product name - priority: Company Name first, then General Info
                product_name = company_name  # Use Company Name first (always prefer this)
                row_data = sponsor_data.get("row_data", {})
                # Only use General Info if Company Name is empty or not available
                if not product_name or product_name.strip() == "":
                    if "General Info" in row_data:
                        general_info = str(row_data["General Info"]).strip() if not pd.isna(row_data.get("General Info")) else ""
                        if general_info:
                            product_name = general_info
            # Find the email body/editor - FreeScout uses contenteditable div
            body_selectors = [
                "div.note-editable[contenteditable='true']",  # FreeScout specific
                "div[contenteditable='true']",
                ".note-editable",
                "textarea[name*='body']",
                "textarea[id*='body']",
                ".email-body",
                "#email-body",
                "iframe[title*='editor']",
                ".editor-body"
            ]
            
            body_element = None
            for selector in body_selectors:
                try:
                    print(f"  Trying body selector: {selector[:50]}...")
                    body_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if body_element:
                        print(f"  ✓ Found body editor with: {selector[:50]}")
                        break
                except:
                    continue
            
            if not body_element:
                raise Exception("Could not find email body editor.")
            
            # Clear existing content and fill
            if body_element.tag_name == 'iframe':
                self.driver.switch_to.frame(body_element)
                editor_body = self.driver.find_element(By.TAG_NAME, "body")
                editor_body.clear()
                editor_body.send_keys(body)
                self.driver.switch_to.default_content()
            elif body_element.tag_name == 'textarea':
                body_element.clear()
                body_element.send_keys(body)
            else:
                # Contenteditable div - use innerHTML to preserve formatting and links
                print("  Setting HTML content to preserve formatting...")
                
                # Use JavaScript to set innerHTML and replace placeholders in one go
                try:
                    # Combined script to clear, set HTML, and replace placeholders
                    # The placeholder is inside HTML tags, so we need to replace it in the HTML string
                    # Use proper regex escaping for JavaScript
                    set_and_replace_script = """
                        var element = arguments[0];
                        var html = arguments[1];
                        var firstName = arguments[2];
                        var companyName = arguments[3];
                        var contactPerson = arguments[4];
                        var productName = arguments[5];
                        
                        // Replace placeholders in the HTML string BEFORE setting it
                        // This ensures we catch placeholders inside HTML tags like <span>[Prospective Sponsor's Name]</span>
                        
                        // Replace [Prospective Sponsor's Name] with first name (all variations)
                        // Use simple string replace (most reliable, no regex escaping issues)
                        // IMPORTANT: Use double quotes or properly escape single quotes
                        var placeholderVariations = [
                            "[Prospective Sponsor's Name]",  // Use double quotes to avoid escaping issues
                            "[Prospective Sponsor&#39;s Name]",
                            "[Prospective Sponsor&apos;s Name]",
                            "[Prospective Sponsor&rsquo;s Name]",
                            "[Prospective Sponsor&#x27;s Name]"
                        ];
                        
                        for (var i = 0; i < placeholderVariations.length; i++) {
                            while (html.indexOf(placeholderVariations[i]) !== -1) {
                                html = html.replace(placeholderVariations[i], firstName);
                            }
                        }
                        
                        // Replace [Company Name]
                        while (html.indexOf('[Company Name]') !== -1) {
                            html = html.replace('[Company Name]', companyName);
                        }
                        
                        // Replace [Customer Company Name] with company name
                        while (html.indexOf('[Customer Company Name]') !== -1) {
                            html = html.replace('[Customer Company Name]', companyName);
                        }
                        
                        // Replace [Customer Company's product name] with product name (handle all apostrophe variations)
                        // Straight apostrophe (U+0027)
                        while (html.indexOf("[Customer Company's product name]") !== -1) {
                            html = html.replace("[Customer Company's product name]", productName);
                        }
                        // Curly quote U+2019 (right single quotation mark) - this is what's in the template
                        while (html.indexOf("[Customer Company's product name]") !== -1) {
                            html = html.replace("[Customer Company's product name]", productName);
                        }
                        // Curly quote U+2018 (left single quotation mark)
                        while (html.indexOf("[Customer Company's product name]") !== -1) {
                            html = html.replace("[Customer Company's product name]", productName);
                        }
                        // HTML entity versions
                        while (html.indexOf("[Customer Company&#39;s product name]") !== -1) {
                            html = html.replace("[Customer Company&#39;s product name]", productName);
                        }
                        while (html.indexOf("[Customer Company&apos;s product name]") !== -1) {
                            html = html.replace("[Customer Company&apos;s product name]", productName);
                        }
                        while (html.indexOf("[Customer Company&rsquo;s product name]") !== -1) {
                            html = html.replace("[Customer Company&rsquo;s product name]", productName);
                        }
                        // Also try regex pattern for any apostrophe variation (straight quote, curly quotes)
                        // Handle U+0027 (straight), U+2019 (right curly), U+2018 (left curly)
                        html = html.replace(/\[Customer Company[''\u2018\u2019]s product name\]/g, productName);
                        
                        // Replace [Customer Company POC] with full contact person name
                        while (html.indexOf('[Customer Company POC]') !== -1) {
                            html = html.replace('[Customer Company POC]', contactPerson);
                        }
                        
                        // Replace [Customer Company POC Name] with full contact person name
                        while (html.indexOf('[Customer Company POC Name]') !== -1) {
                            html = html.replace('[Customer Company POC Name]', contactPerson);
                        }
                        
                        // Replace combined placeholder
                        var combinedPlaceholder = "[Customer Company POC][Prospective Sponsor's Name]";  // Use double quotes
                        while (html.indexOf(combinedPlaceholder) !== -1) {
                            html = html.replace(combinedPlaceholder, contactPerson);
                        }
                        
                        // Clear existing content
                        element.innerHTML = '';
                        element.textContent = '';
                        
                        // Clean up HTML: remove only truly empty/duplicate elements, preserve formatting
                        // IMPORTANT: Don't remove <div><br></div> or <span><br></span> - these create line breaks!
                        // Only remove empty divs with nested empty spans (like <div><span></span></div> with no content)
                        html = html.replace(new RegExp('<div[^>]*>\\s*<span[^>]*>\\s*</span>\\s*</div>', 'gi'), '');
                        // Remove completely empty divs (no content, no <br>, no spans) - but preserve <div><br></div>
                        html = html.replace(new RegExp('<div[^>]*>\\s*</div>', 'gi'), '');
                        // DON'T remove trailing <br> tags - they might be intentional line breaks
                        // DON'T reduce <br> tags - preserve the structure as-is
                        // Remove only completely empty paragraphs (no content, no <br>)
                        html = html.replace(new RegExp('<p[^>]*>\\s*</p>', 'gi'), '');
                        // Preserve ALL whitespace and newlines - only trim leading/trailing whitespace
                        html = html.trim();
                        
                        // Debug: Log HTML length before setting
                        console.log('Setting HTML content, length: ' + html.length);
                        console.log('HTML contains <br> tags: ' + (html.indexOf('<br') !== -1));
                        console.log('HTML contains <p> tags: ' + (html.indexOf('<p') !== -1));
                        
                        // Clear the element first
                        element.innerHTML = '';
                        element.textContent = '';
                        
                        // Set the HTML content with placeholders already replaced
                        // Use direct innerHTML (this was working before and preserved formatting)
                        // Don't use Summernote API - it's causing empty content and nested editors
                        element.innerHTML = html;
                        var methodUsed = 'innerhtml';
                        console.log('Content set via direct innerHTML');
                        
                        // Verify the HTML was set correctly by checking for key elements
                        var verifyHtml = element.innerHTML;
                        console.log('HTML after setting, length: ' + verifyHtml.length);
                        console.log('HTML after setting contains <br> tags: ' + (verifyHtml.indexOf('<br') !== -1));
                        console.log('HTML after setting contains <p> tags: ' + (verifyHtml.indexOf('<p') !== -1));
                        
                        // Use the verified HTML (browser may have normalized it slightly, but structure should be preserved)
                        var finalHtml = verifyHtml;
                        
                        // Update hidden textarea with the original HTML (before browser normalization)
                        var editorContainer = element.closest ? element.closest('.note-editor') : null;
                        if (editorContainer) {
                            var hiddenTextarea = editorContainer.querySelector('textarea.note-codable');
                            if (hiddenTextarea) {
                                hiddenTextarea.value = finalHtml;
                                console.log('Hidden textarea updated');
                            }
                        }
                        
                        // Trigger events to notify the editor that content has changed
                        var inputEvent = new Event('input', { bubbles: true, cancelable: true });
                        element.dispatchEvent(inputEvent);
                        var changeEvent = new Event('change', { bubbles: true, cancelable: true });
                        element.dispatchEvent(changeEvent);
                        
                        // Also trigger Summernote-specific events if available
                        if (typeof jQuery !== 'undefined') {
                            try {
                                jQuery(element).trigger('summernote.change');
                            } catch (e) {
                                // Ignore if event doesn't exist
                            }
                        }
                        
                        // Debug: Check if content is empty after setting
                        var isEmpty = !finalHtml || finalHtml.trim().length === 0;
                        if (isEmpty) {
                            console.error('Error: Content appears to be empty after setting!');
                        } else {
                            console.log('Content set successfully, final length: ' + finalHtml.length);
                        }
                        
                        // Check if placeholder was replaced
                        var replacementSuccess = finalHtml.indexOf("[Prospective Sponsor") === -1;
                        
                        // Return simple result
                        return {
                            replacementSuccess: replacementSuccess,
                            methodUsed: methodUsed,
                            isEmpty: isEmpty,
                            htmlLength: finalHtml.length
                        };
                    """
                    result = self.driver.execute_script(set_and_replace_script, body_element, body, contact_first_name, company_name, contact_person, product_name)
                    # DOM should update immediately, minimal wait if needed
                    time.sleep(0.1)
                    
                    # Update hidden textarea to ensure form submission uses correct content
                    print("  Updating hidden textarea...")
                    self.driver.execute_script("""
                        var element = arguments[0];
                        var container = element.closest('.note-editor');
                        if (container) {
                            var textarea = container.querySelector('textarea.note-codable');
                            if (textarea) {
                                textarea.value = element.innerHTML;
                                var event = new Event('change', { bubbles: true });
                                textarea.dispatchEvent(event);
                            }
                        }
                    """, body_element)
                    print("  ✓ Hidden textarea updated")
                    
                    # Handle result and log debug info
                    if isinstance(result, dict):
                        replacement_success = result.get('replacementSuccess', False)
                        method_used = result.get('methodUsed', 'unknown')
                        is_empty = result.get('isEmpty', False)
                        html_length = result.get('htmlLength', 0)
                        
                        # Debug logging
                        print(f"  Method used: {method_used}")
                        print(f"  HTML length: {html_length}")
                        if is_empty:
                            print("  ⚠️  Warning: Content appears to be empty!")
                        else:
                            print("  ✓ Content is not empty")
                        
                        if replacement_success:
                            print("  ✓ Placeholders replaced successfully")
                        else:
                            print("  ⚠️  Warning: Placeholder might still be in DOM")
                    else:
                        # Backward compatibility
                        replacement_success = result if isinstance(result, bool) else False
                        print("  ⚠️  Unexpected result format")
                except Exception as e:
                    print(f"  Warning: Could not set HTML directly: {e}")
                    import traceback
                    traceback.print_exc()
                    # Fallback: Clear and use send_keys
                    print("  Falling back to text input...")
                    body_element.click()
                    time.sleep(0.2)
                    body_element.send_keys(Keys.CONTROL + "a")
                    time.sleep(0.2)
                    body_element.send_keys(Keys.DELETE)
                    time.sleep(0.3)
                    # Convert HTML to text for fallback
                    # Simple HTML to text conversion
                    text_body = re.sub(r'<[^>]+>', '', body)  # Remove HTML tags
                    text_body = html.unescape(text_body)  # Decode HTML entities
                    body_element.send_keys(text_body)
                    time.sleep(0.5)
            
            print("✓ Email body filled")
            
        except Exception as e:
            raise Exception(f"Failed to fill email body: {e}")
    
    def fill_subject_field(self, subject: str):
        """Fill the subject field"""
        try:
            print("Looking for subject field...")
            # FreeScout uses id="subject"
            subject_selectors = [
                "#subject",  # FreeScout specific
                "input[name='subject']",
                "input[id='subject']",
                "input[name*='subject']",
                "input[id*='subject']",
                ".subject-field",
                "input[placeholder*='Subject']"
            ]
            
            subject_field = None
            for selector in subject_selectors:
                try:
                    print(f"  Trying subject selector: {selector[:50]}...")
                    subject_field = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if subject_field:
                        print(f"  ✓ Found subject field with: {selector[:50]}")
                        break
                except Exception as e:
                    print(f"  ✗ Not found: {str(e)[:50]}")
                    continue
            
            if not subject_field:
                raise Exception("Could not find subject field.")
            
            print(f"Filling subject field: {subject[:50]}...")
            subject_field.clear()
            subject_field.send_keys(subject)
            # No need to wait - field is filled immediately
            print("✓ Subject field filled")
            
        except Exception as e:
            raise Exception(f"Failed to fill subject field: {e}")
    
    def get_email_preview(self) -> Dict[str, str]:
        """Get current email preview (To, Subject, Body) for confirmation"""
        try:
            preview = {}
            
            # Get To field
            to_selectors = [
                "input[placeholder*='To']",
                "input[name='to']",
                "#to",
                ".to-field"
            ]
            for selector in to_selectors:
                try:
                    to_field = self.driver.find_element(By.CSS_SELECTOR, selector)
                    preview['to'] = to_field.get_attribute('value') or ''
                    break
                except:
                    continue
            
            # Get Subject field
            subject_selectors = [
                "input[name*='subject']",
                "#subject",
                ".subject-field"
            ]
            for selector in subject_selectors:
                try:
                    subject_field = self.driver.find_element(By.CSS_SELECTOR, selector)
                    preview['subject'] = subject_field.get_attribute('value') or ''
                    break
                except:
                    continue
            
            # Get Body - FreeScout uses contenteditable div
            body_selectors = [
                "div.note-editable[contenteditable='true']",  # FreeScout specific
                "div[contenteditable='true']",
                ".note-editable",
                "textarea[name*='body']",
                ".email-body"
            ]
            for selector in body_selectors:
                try:
                    body_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if body_element.tag_name == 'iframe':
                        self.driver.switch_to.frame(body_element)
                        preview['body'] = self.driver.find_element(By.TAG_NAME, "body").text
                        self.driver.switch_to.default_content()
                    else:
                        # For contenteditable div, get text content but preserve line breaks
                        # Use innerText which preserves line breaks from <br> and <p> tags
                        text_content = body_element.get_attribute('innerText') or body_element.text or body_element.get_attribute('textContent') or ''
                        # If innerText doesn't preserve breaks, try to convert HTML to text with breaks
                        if not text_content or len(text_content) < 50:
                            html_content = body_element.get_attribute('innerHTML') or ''
                            # Convert <br> and </p><p> to newlines for preview
                            import re
                            text_content = re.sub(r'<br\s*/?>', '\n', html_content, flags=re.IGNORECASE)
                            text_content = re.sub(r'</p>\s*<p[^>]*>', '\n\n', text_content, flags=re.IGNORECASE)
                            text_content = re.sub(r'<[^>]+>', '', text_content)  # Remove remaining HTML tags
                            text_content = html.unescape(text_content)  # Decode HTML entities
                        preview['body'] = text_content
                        
                        # Debug: Check if placeholder is in the preview
                        if "[Prospective Sponsor" in text_content:
                            print(f"  ⚠️  DEBUG: Placeholder still found in preview text!")
                            # Get HTML to see what's actually there
                            html_content = body_element.get_attribute('innerHTML') or ''
                            if "[Prospective Sponsor" in html_content:
                                idx = html_content.find("[Prospective Sponsor")
                                snippet = html_content[max(0, idx-50):idx+80]
                                print(f"  DEBUG: HTML snippet: ...{snippet}...")
                    break
                except:
                    continue
            
            return preview
            
        except Exception as e:
            print(f"Warning: Could not get full email preview: {e}")
            return {}
    
    def send_email(self):
        """Click the Send button"""
        try:
            # Before sending, ensure hidden textarea is synchronized
            print("Preparing to send email...")
            body_selectors = [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable"
            ]
            
            for selector in body_selectors:
                try:
                    body_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    # Update hidden textarea to ensure form submission uses correct content
                    editor_container = self.driver.execute_script("""
                        var element = arguments[0];
                        return element.closest ? element.closest('.note-editor') : null;
                    """, body_element)
                    
                    if editor_container:
                        current_html = body_element.get_attribute('innerHTML') or ''
                        self.driver.execute_script("""
                            var container = arguments[0];
                            var html = arguments[1];
                            var textarea = container.querySelector('textarea.note-codable');
                            if (textarea) {
                                textarea.value = html;
                                var event = new Event('change', { bubbles: true });
                                textarea.dispatchEvent(event);
                            }
                        """, editor_container, current_html)
                        print("  ✓ Hidden textarea synchronized")
                    
                    break
                except:
                    continue
            
            # Find send button
            send_selectors = [
                "button.btn-send-text",  # FreeScout specific - most specific first
                "button.btn-reply-submit.btn-send-text",  # Full class match
                "button[type='submit']",
                "button[aria-label*='Send']",
                ".send-button",
                "#send-button",
                "button.send",
                "//button[contains(text(), 'Send')]"  # XPath fallback
            ]
            
            send_button = None
            for selector in send_selectors:
                try:
                    if selector.startswith("//"):
                        send_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        send_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if send_button:
                        print(f"  ✓ Found Send button with selector: {selector}")
                        break
                except:
                    continue
            
            if not send_button:
                raise Exception("Could not find Send button.")
            
            # Send the email
            print("  Clicking Send button...")
            send_button.click()
            print("✓ Email sent successfully")
            
        except Exception as e:
            raise Exception(f"Failed to send email: {e}")
    
    def send_sponsor_email(self, sponsor_data: Dict, confirm_before_send: bool = True) -> bool:
        """
        Complete workflow to send email for a sponsor
        
        Args:
            sponsor_data: Dictionary with sponsor information
            confirm_before_send: Whether to show preview and wait for confirmation
            
        Returns:
            True if email was sent successfully, False otherwise
        """
        try:
            template_name = sponsor_data.get('template_name', 'Unknown')
            # Parse and display all emails
            email_str = sponsor_data.get('email', 'N/A')
            if email_str != 'N/A' and (',' in email_str or ';' in email_str):
                if ',' in email_str:
                    emails = [e.strip() for e in email_str.split(',')]
                else:
                    emails = [e.strip() for e in email_str.split(';')]
                emails = [e for e in emails if e and '@' in e]
                emails_display = ', '.join(emails) if emails else email_str
            else:
                emails_display = email_str
            
            print(f"\n📧 Processing email for: {emails_display}")
            print(f"📋 Template: {template_name}")
            
            # Step 1: Click new conversation
            self.click_new_conversation()
            
            # Step 2: Fill To field
            self.fill_to_field(sponsor_data['email'])
            
            # Step 3: Select template
            self.select_template(template_name)
            
            # Step 4: Extract template content
            # IMPORTANT: Don't clear before extracting - we need the template content first
            subject, body = self.extract_template_content()
            
            # IMPORTANT: Clear the editor NOW after extraction, before we set the replaced content
            # This prevents duplicate content - the template is still in the editor after extraction
            print("  Clearing editor after extraction to prevent duplicates...")
            body_selectors = [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable"
            ]
            for selector in body_selectors:
                try:
                    body_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    # Clear completely - this removes the original template content
                    self.driver.execute_script("arguments[0].innerHTML = ''; arguments[0].textContent = '';", body_element)
                    # Clear happens immediately, no need to wait
                    print("  ✓ Editor cleared")
                    break
                except:
                    continue
            
            # Debug: Check what we extracted
            print(f"\n  DEBUG: Checking extracted body for placeholders...")
            if "[Prospective Sponsor's Name]" in body:
                print(f"    ✓ Found '[Prospective Sponsor's Name]' in extracted body")
            elif "[Prospective Sponsor" in body:
                print(f"    ⚠️  Found partial match '[Prospective Sponsor' - might be split or encoded")
                # Show snippet
                idx = body.find("[Prospective Sponsor")
                if idx >= 0:
                    snippet = body[max(0, idx-20):idx+80]
                    print(f"    Snippet: ...{snippet}...")
            
            # Step 5: Fill placeholders
            filled_body = self.fill_placeholders(body, sponsor_data)
            
            # Debug: Verify after replacement
            print(f"\n  DEBUG: Checking filled body after replacement...")
            # Check for exact match
            if "[Prospective Sponsor's Name]" in filled_body:
                count_remaining = filled_body.count("[Prospective Sponsor's Name]")
                print(f"    ⚠️  '[Prospective Sponsor's Name]' still found ({count_remaining} times) after Python replacement!")
                idx = filled_body.find("[Prospective Sponsor's Name]")
                if idx >= 0:
                    snippet = filled_body[max(0, idx-50):idx+80]
                    print(f"    Snippet: ...{snippet}...")
                    # Show the actual characters around it to check for encoding issues
                    print(f"    Character codes around placeholder:")
                    for i in range(max(0, idx-5), min(len(filled_body), idx+len("[Prospective Sponsor's Name]")+5)):
                        char = filled_body[i]
                        print(f"      [{i}]: '{char}' (ord: {ord(char)})")
            else:
                print(f"    ✓ '[Prospective Sponsor's Name]' successfully replaced in Python string")
            
            # Also check for any variation
            if "[Prospective Sponsor" in filled_body:
                idx = filled_body.find("[Prospective Sponsor")
                if idx >= 0:
                    end_idx = filled_body.find("]", idx)
                    if end_idx > idx:
                        variation = filled_body[idx:end_idx+1]
                        print(f"    ⚠️  Found variation: '{variation}' - checking character codes...")
                        for char in variation:
                            print(f"      '{char}' (ord: {ord(char)})")
            
            # Step 6: Remove subject line from body (already done in extract_template_content)
            # Step 7: Fill email body (pass sponsor_data for placeholder replacement)
            self.fill_email_body(filled_body, sponsor_data)
            
            # Step 8: Fill subject field
            if subject:
                self.fill_subject_field(subject)
            
            # Step 9: Show preview and get confirmation
            if confirm_before_send:
                preview = self.get_email_preview()
                # Parse multiple emails for display
                email_str = sponsor_data['email']
                if ',' in email_str or ';' in email_str:
                    emails = [e.strip() for e in re.split(r'[,;]', email_str) if e.strip()]
                    to_display = ', '.join(emails)
                else:
                    to_display = email_str
                
                print("\n" + "="*80)
                print("EMAIL PREVIEW:")
                print("="*80)
                print(f"To: {preview.get('to', to_display)}")
                print(f"Subject: {preview.get('subject', subject)}")
                print(f"\nBody:\n{preview.get('body', filled_body[:500])}")
                if len(filled_body) > 500:
                    print("... (truncated)")
                print("="*80)
                
                confirmation = input("\nSend this email? (y/n/skip): ").strip().lower()
                if confirmation == 'n':
                    print("Email cancelled by user.")
                    return False
                elif confirmation == 'skip':
                    print("Skipping this email.")
                    return False
                # 'y' or any other input will proceed
            
            # Step 10: Send email
            self.send_email()
            # Show all emails in success message
            email_str = sponsor_data['email']
            if ',' in email_str or ';' in email_str:
                emails = [e.strip() for e in re.split(r'[,;]', email_str) if e.strip() and '@' in e]
                email_display = ', '.join(emails) if emails else email_str
            else:
                email_display = email_str
            print(f"✓ Email sent successfully to: {email_display}")
            print(f"  Template used: {template_name}")
            return True
            
        except Exception as e:
            # Show all emails in error message
            email_str = sponsor_data['email']
            if ',' in email_str or ';' in email_str:
                emails = [e.strip() for e in re.split(r'[,;]', email_str) if e.strip() and '@' in e]
                email_display = ', '.join(emails) if emails else email_str
            else:
                email_display = email_str
            print(f"✗ Error sending email to {email_display}: {e}")
            return False
    
    def close(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()

