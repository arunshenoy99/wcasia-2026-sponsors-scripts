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
import urllib.parse
from typing import Dict, List, Optional, Tuple

from config import (
    FREESCOUT_URL,
    FREESCOUT_EMAIL,
    FREESCOUT_PASSWORD,
    HEADLESS_MODE,
    BROWSER_WAIT_TIME,
    BROWSER_DELAY_SCALE,
    REPLY_TO_THREAD_TEMPLATE_PATTERN,
)


def _delay(sec: float) -> None:
    """Sleep for sec * BROWSER_DELAY_SCALE (min 0.05s). Use for UI-settle waits so runs can be sped up via .env."""
    time.sleep(max(0.05, sec * BROWSER_DELAY_SCALE))


def _name_from_email(email_str: str) -> Tuple[str, str]:
    """Derive display name from email: local part before @, split by . _ -, capitalize. Returns (full_name, first_name)."""
    if not email_str or "@" not in email_str:
        return "", ""
    local = email_str.split("@")[0].strip()
    if not local:
        return "", ""
    parts = re.split(r"[._-]+", local)
    parts = [p.capitalize() for p in parts if p]
    if not parts:
        return "", ""
    first = parts[0]
    full = " ".join(parts)
    return full, first


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
        _delay(2)  # Wait for page load
        
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
            _delay(0.5)
            
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
            _delay(0.5)
            
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
            
            
        except Exception as e:
            raise Exception(f"Login failed: {e}")
    
    def _parse_emails(self, email_str: str) -> List[str]:
        """Parse email string (comma/semicolon separated) into list of valid email addresses.
        Strips leading 'CC:', 'BCC:', 'To:' from each part so values like
        'a@x.com, CC: b@y.com' parse to ['a@x.com', 'b@y.com'].
        """
        if not email_str or email_str == 'N/A':
            return []
        s = str(email_str).strip()
        if ',' in s:
            parts = [e.strip() for e in s.split(',')]
        elif ';' in s:
            parts = [e.strip() for e in s.split(';')]
        else:
            parts = [s]
        result = []
        for part in parts:
            if not part or '@' not in part:
                continue
            # Strip CC:/BCC:/To: prefix (case-insensitive)
            for prefix in ('CC:', 'BCC:', 'To:', 'Cc:', 'Bcc:'):
                if part.upper().startswith(prefix.upper()):
                    part = part[len(prefix):].strip()
                    break
            if part and '@' in part:
                result.append(part)
        return result
    
    def _open_search_tabs_for_emails(self, emails: List[str]) -> List[str]:
        """Open FreeScout search URL for each email in a new tab. Return list of new tab handles in same order as emails. Switches back to original window."""
        if not emails or not FREESCOUT_URL:
            return []
        base = FREESCOUT_URL.rstrip('/')
        original_handle = self.driver.current_window_handle
        new_handles = []
        for email in emails:
            handles_before = set(self.driver.window_handles)
            url = f"{base}/search?q={urllib.parse.quote(email, safe='')}"
            self.driver.execute_script("window.open(arguments[0], '_blank');", url)
            _delay(0.4)
            added = [h for h in self.driver.window_handles if h not in handles_before]
            if added:
                new_handles.append(added[-1])
        self.driver.switch_to.window(original_handle)
        _delay(0.5)
        return new_handles
    
    def _close_search_tabs(self, tab_handles: List[str]) -> None:
        """Close the given tab handles and switch back to the main window."""
        if not tab_handles:
            return
        main_handle = self.driver.current_window_handle
        for h in tab_handles:
            try:
                if h in self.driver.window_handles:
                    self.driver.switch_to.window(h)
                    self.driver.close()
            except Exception:
                pass
        try:
            if main_handle in self.driver.window_handles:
                self.driver.switch_to.window(main_handle)
            elif self.driver.window_handles:
                self.driver.switch_to.window(self.driver.window_handles[0])
            # Return to base FreeScout URL so next sponsor starts from a consistent state (avoids stuck on mailbox view)
            if FREESCOUT_URL:
                base = FREESCOUT_URL.rstrip('/')
                self.driver.get(base)
                _delay(1.5)
        except Exception:
            pass

    def _is_conversation_link(self, href: str) -> bool:
        """True if href looks like a conversation (not search, new-ticket, or mailbox root)."""
        if not href or "search" in href or "new-ticket" in href or "new" in href.lower():
            return False
        path = href.split("?", 1)[0].rstrip("/")
        segments = [s for s in path.split("/") if s]
        if segments and segments[-1] == "mailbox":
            return False
        return True

    def _switch_to_thread_tab_if_opened(self, handles_before_click: set) -> None:
        """If a new tab was opened (e.g. conversation link with target=_blank), switch driver to it."""
        for _ in range(6):
            _delay(0.3)
            now = self.driver.window_handles
            new_handles = [h for h in now if h not in handles_before_click]
            if new_handles:
                self.driver.switch_to.window(new_handles[-1])
                _delay(0.5)
                return

    def _open_first_thread_from_current_page(self) -> bool:
        """On current page (must be search results), find first conversation link, click it, switch to new tab if opened. Returns True if thread opened."""
        first_conversation_selectors = [
            "table tbody tr a[href*='/conversations/']",
            "table tbody tr a[href*='/ticket/']",
            ".conv-list a[href*='/conversations/']",
            ".conversation-list a[href*='/conversations/']",
            "a[href*='/conversations/']",
            "a[href*='/ticket/']",
            ".conv-list a",
            ".conversation-list a",
            "table tbody tr a",
            "a[href*='conversation']",
            "a[href*='/mailbox/']",
        ]
        first_conv_href = None
        handles_before = set(self.driver.window_handles)
        for selector in first_conversation_selectors:
            try:
                links = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for link in links:
                    href = (link.get_attribute("href") or "").strip()
                    if not self._is_conversation_link(href):
                        continue
                    first_conv_href = href
                    link.click()
                    _delay(1)
                    self._switch_to_thread_tab_if_opened(handles_before)
                    if "search" not in self.driver.current_url:
                        self._wait_for_conversation_view()
                        return True
                    break
                if first_conv_href:
                    break
            except Exception:
                continue
        if first_conv_href and "search" in self.driver.current_url:
            self.driver.get(first_conv_href)
            _delay(1)
            self._wait_for_conversation_view()
            return True
        try:
            first_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'/conversations/') or contains(@href,'/ticket/')][1]"))
            )
            first_conv_href = first_link.get_attribute("href")
            if first_conv_href and self._is_conversation_link(first_conv_href):
                first_link.click()
                _delay(1)
                self._switch_to_thread_tab_if_opened(handles_before)
                if "search" in self.driver.current_url:
                    self.driver.get(first_conv_href)
                    _delay(1)
                self._wait_for_conversation_view()
                return True
        except Exception:
            pass
        return False

    def open_thread_for_email(self, email: str, already_on_search_page: bool = False) -> bool:
        """Navigate to search for email (unless already_on_search_page), open the first conversation. FreeScout may open in new tab; we switch to it."""
        if not already_on_search_page:
            if not FREESCOUT_URL or not email or "@" not in email:
                return False
            base = FREESCOUT_URL.rstrip("/")
            url = f"{base}/search?q={urllib.parse.quote(email, safe='')}"
            self.driver.get(url)
            _delay(1)
        return self._open_first_thread_from_current_page()

    def _wait_for_conversation_view(self) -> None:
        """Wait for the conversation (thread) view to load so reply editor/buttons are available."""
        for _ in range(15):
            try:
                url = self.driver.current_url.lower()
                if "conversation" in url or "ticket" in url:
                    break
            except Exception:
                pass
            _delay(0.3)
        # Wait for something that exists only on conversation view (reply icon or reply form)
        try:
            self.wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span.conv-reply, .note-editable, [aria-label='Reply']")
            ))
        except Exception:
            pass
        _delay(0.5)

    def open_reply_form(self) -> None:
        """Click the Reply control to open/expand the reply form so the reply editor is visible."""
        reply_btn_selectors = [
            "span.conv-reply.conv-action[aria-label='Reply']",
            "span.conv-reply[aria-label='Reply']",
            "[aria-label='Reply']",
            "span.conv-reply.conv-action.glyphicon-share-alt",
        ]
        for selector in reply_btn_selectors:
            try:
                btn = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                _delay(0.3)
                try:
                    btn.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", btn)
                _delay(1)
                return
            except Exception:
                continue
        raise Exception("Could not find Reply button to open reply form.")

    def focus_reply_editor_and_fill(self, body: str) -> None:
        """In the open conversation view, find the visible reply editor (after Reply was clicked), clear it, and set body."""
        reply_editor_selectors = [
            "div.note-editable[contenteditable='true']",
            "div[contenteditable='true']",
            ".note-editable",
            ".reply-form div[contenteditable='true']",
            "#reply-body",
            "textarea[name*='body']",
            "textarea.reply-body",
        ]
        for selector in reply_editor_selectors:
            try:
                try:
                    el = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selector)))
                except Exception:
                    els = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    el = next((e for e in els if e.is_displayed()), els[0] if els else None)
                if not el:
                    continue
                el.click()
                _delay(0.3)
                if el.tag_name == "textarea":
                    el.clear()
                    el.send_keys(html.unescape(re.sub(r"<[^>]+>", "", body)))
                else:
                    self.driver.execute_script("arguments[0].innerHTML = arguments[1];", el, body)
                return
            except Exception:
                continue
        raise Exception("Could not find reply editor in conversation view.")

    def click_send_reply(self) -> None:
        """Find and click the visible Send button (btn-send-text sends the reply; multiple may exist, one visible)."""
        send_selectors = [
            "button.btn-send-text",
            "button.btn-reply-submit.btn-send-text",
            "button[type='submit']",
            "button.btn-reply-submit",
            "button.btn-primary",
        ]
        for selector in send_selectors:
            try:
                els = self.driver.find_elements(By.CSS_SELECTOR, selector)
                btn = next((e for e in els if e.is_displayed()), None)
                if not btn:
                    continue
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                _delay(0.3)
                try:
                    btn.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", btn)
                _delay(1)
                return
            except Exception:
                continue
        try:
            buttons = self.driver.find_elements(
                By.XPATH,
                "//button[contains(.,'Send') or contains(.,'Reply')] | //input[@type='submit' and (@value='Send' or @value='Reply' or contains(@value,'Send'))]",
            )
            btn = next((b for b in buttons if b.is_displayed()), None)
            if btn:
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                _delay(0.2)
                try:
                    btn.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", btn)
                _delay(1)
                return
        except Exception:
            pass
        raise Exception("Could not find Send/Reply button in conversation view.")
    
    def click_new_conversation(self):
        """Click the mail icon/new conversation button"""
        try:
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
                    if selector.startswith("//"):
                        new_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        new_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if new_button:
                        break
                except Exception:
                    continue
            
            if not new_button:
                try:
                    icon = self.driver.find_element(By.CSS_SELECTOR, "i.glyphicon.glyphicon-envelope")
                    new_button = icon.find_element(By.XPATH, "./ancestor::a")
                except Exception:
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
                        break
                    except Exception:
                        continue
                
                if not new_button:
                    raise Exception("Could not find new conversation button. Please check FreeScout interface or update selectors.")
            
            new_button.click()
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field, #subject, div.note-editable")))
            except Exception:
                pass
            # Dismiss any assignee dropdown so we don't interact with it when filling To
            try:
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                _delay(0.5)
            except Exception:
                pass
            
        except Exception as e:
            raise Exception(f"Failed to click new conversation: {e}")
    
    def _get_to_field_select2_container(self):
        """Find the Select2 container for the To/recipient field (not assignee). Returns (container, input) or (None, None)."""
        # FreeScout new conversation form has Assignee first, then To. Target To specifically.
        # 1) Try: find the row that has "To" or "Recipient" label, then get Select2 inside it
        try:
            row = self.driver.find_element(
                By.XPATH,
                "//label[contains(.,'To') or contains(.,'Recipient') or contains(.,'Customer')]/ancestor::div[contains(@class,'form-group') or contains(@class,'row')][1]"
            )
            container = row.find_element(By.CSS_SELECTOR, ".select2-container, span.select2-container")
            inp = row.find_element(By.CSS_SELECTOR, "input.select2-search__field")
            if container and inp:
                return container, inp
        except Exception:
            pass
        # 2) Try: find select that is for To/recipient (name often 'to' or 'customer_id'), then its Select2 wrapper
        for name in ("to", "customer_id", "customer"):
            try:
                sel = self.driver.find_element(By.CSS_SELECTOR, f"select[name='{name}'], select[id*='{name}']")
                container = self.driver.execute_script(
                    "var s = arguments[0]; return s.nextElementSibling && s.nextElementSibling.classList && s.nextElementSibling.classList.contains('select2-container') ? s.nextElementSibling : null;",
                    sel,
                )
                if container:
                    inp = container.find_element(By.CSS_SELECTOR, "input.select2-search__field")
                    return container, inp
            except Exception:
                continue
        # 3) Fallback: use the second Select2 on the page (first is often assignee)
        try:
            containers = self.driver.find_elements(By.CSS_SELECTOR, ".select2-container, span.select2-container")
            if len(containers) >= 2:
                container = containers[1]
                container.click()
                _delay(0.2)
                inp = self.driver.find_element(By.CSS_SELECTOR, "input.select2-search__field")
                return container, inp
        except Exception:
            pass
        return None, None

    def fill_to_field(self, email: str):
        """
        Fill the 'To' field with email address(es) (Select2 dropdown)
        Supports multiple emails separated by comma, semicolon, or space.
        Targets the To/recipient field specifically (not the assignee dropdown).
        """
        try:
            # Parse multiple emails (reuse _parse_emails so CC:/BCC:/To: are stripped)
            valid_emails = self._parse_emails(email)
            if not valid_emails:
                raise Exception(f"No valid email addresses found in: {email}")
            
            # Find To field specifically (avoid assignee Select2)
            select2_container, to_field = self._get_to_field_select2_container()
            if not to_field:
                # Legacy fallback: prefer second Select2 (first is often assignee in new-conversation form)
                all_containers = self.driver.find_elements(By.CSS_SELECTOR, ".select2-container, span.select2-container")
                if len(all_containers) >= 2:
                    select2_container = all_containers[1]
                else:
                    select2_container = None
                    for sel in [".select2-container", "span.select2-container", ".select2-selection", "span.select2-selection"]:
                        try:
                            select2_container = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                            break
                        except Exception:
                            continue
                for sel in ["input.select2-search__field", ".select2-search__field", "input.select2-search input"]:
                    try:
                        to_field = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                        break
                    except Exception:
                        continue
            
            if not to_field:
                raise Exception("Could not find 'To' field. Please check FreeScout interface.")
            
            to_selectors = ["input.select2-search__field", ".select2-search__field", "input.select2-search input"]
            select2_selectors = [".select2-container", "span.select2-container", ".select2-selection", "span.select2-selection"]
            
            # Add each email address
            for i, email_addr in enumerate(valid_emails, 1):
                # Use the To field's container (avoid clicking assignee)
                container_to_use = select2_container
                if not container_to_use:
                    for selector in select2_selectors:
                        try:
                            container_to_use = WebDriverWait(self.driver, 1).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                            )
                            break
                        except Exception:
                            continue
                if container_to_use:
                    try:
                        container_to_use.click()
                        _delay(0.2)
                    except Exception:
                        pass
                
                # Find the input field again (recreated after adding previous email)
                to_field = None
                if select2_container:
                    try:
                        to_field = select2_container.find_element(By.CSS_SELECTOR, "input.select2-search__field")
                    except Exception:
                        pass
                if not to_field:
                    for selector in to_selectors:
                        try:
                            to_field = WebDriverWait(self.driver, 1).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                            )
                            break
                        except Exception:
                            continue
                if not to_field:
                    raise Exception("Could not find 'To' field after opening Select2.")
                
                to_field.send_keys(email_addr)
                try:
                    WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option, .select2-selection__choice"))
                    )
                except Exception:
                    pass
                try:
                    option = WebDriverWait(self.driver, 1).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-results__option--highlighted, .select2-results__option"))
                    )
                    option.click()
                except Exception:
                    to_field.send_keys(Keys.RETURN)
                if i < len(valid_emails):
                    _delay(0.3)
            
            try:
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                _delay(0.2)
            except Exception:
                pass
            
        except Exception as e:
            raise Exception(f"Failed to fill 'To' field: {e}")
    
    def add_tag(self, tag_name: str):
        """Add a tag to the conversation"""
        try:
            
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
                        break
                except:
                    continue
            
            if not tag_icon:
                return
            
            # Click the tag icon to open dropdown
            tag_icon.click()
            _delay(0.3)  # Wait for dropdown to open
            
            # Find the tag dropdown container first to scope all searches
            tag_dropdown = None
            try:
                tag_dropdown = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#add-tag-wrap"))
                )
            except Exception:
                return
            
            # Wait a bit for Select2 to initialize
            _delay(0.5)
            
            # Try to find and click the Select2 container or selection area to open/focus
            # Try multiple approaches to open the Select2
            select2_opened = False
            try:
                # Approach 1: Find Select2 container within tag dropdown
                tag_select2_container = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-container, span.select2-container")
                tag_select2_container.click()
                select2_opened = True
                _delay(0.3)
            except:
                try:
                    # Approach 2: Find Select2 selection area and click it
                    tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                    tag_select2_selection.click()
                    select2_opened = True
                    _delay(0.3)
                except:
                    try:
                        # Approach 3: Find the hidden select element and trigger Select2
                        hidden_select = tag_dropdown.find_element(By.CSS_SELECTOR, "select.tag-input")
                        self.driver.execute_script("arguments[0].dispatchEvent(new Event('select2:open'));", hidden_select)
                        select2_opened = True
                        _delay(0.3)
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
                        tag_input.find_element(By.XPATH, "./ancestor::*[@id='add-tag-wrap' or contains(@class, 'conv-new-tag-dd')]")
                        break
                    except:
                        # If not in tag dropdown, continue to next selector
                        tag_input = None
                        continue
                except:
                    continue
            
            if not tag_input:
                return
            
            # Click the input to ensure it's focused
            try:
                tag_input.click()
                _delay(0.2)
            except:
                pass
            
            # Type the tag name
            tag_input.clear()
            tag_input.send_keys(tag_name)
            _delay(0.5)  # Wait for Select2 to process and show results
            
            # Wait for Select2 results to appear
            try:
                WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option"))
                )
            except Exception:
                pass
            
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
                    _delay(0.2)
                    option.click()
                    _delay(0.5)
                else:
                    raise Exception("Tag option not found in dropdown")
                
                # Verify tag was selected in Select2 - check within tag dropdown's Select2 only
                try:
                    # Find the Select2 selection within the tag dropdown
                    tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                    tag_chip = tag_select2_selection.find_element(By.CSS_SELECTOR, ".select2-selection__choice")
                    if tag_name in tag_chip.get_attribute('title') or tag_name in tag_chip.text:
                        tag_selected_in_select2 = True
                except:
                    # Alternative: check if tag name appears in selection text
                    try:
                        tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                        if tag_name in tag_select2_selection.text:
                            tag_selected_in_select2 = True
                    except:
                        pass
            except Exception:
                tag_input.send_keys(Keys.RETURN)
                _delay(0.5)  # Wait for Enter to be processed
                
                # Check if tag was selected - scoped to tag dropdown
                try:
                    tag_select2_selection = tag_dropdown.find_element(By.CSS_SELECTOR, ".select2-selection")
                    tag_chip = tag_select2_selection.find_element(By.CSS_SELECTOR, ".select2-selection__choice")
                    if tag_name in tag_chip.get_attribute('title') or tag_name in tag_chip.text:
                        tag_selected_in_select2 = True
                except:
                    pass
            
            _delay(0.3)
            
            # Find and click the tick/ok button - MUST be within tag dropdown
            ok_button = None
            
            # Try multiple approaches to find the button (prioritize direct selector)
            try:
                # Approach 1: Find button directly (most reliable)
                ok_button = WebDriverWait(self.driver, 2).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#add-tag-wrap button.btn-default"))
                )
            except:
                try:
                    # Approach 2: Find by icon and get parent button
                    icon = WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#add-tag-wrap i.glyphicon-ok"))
                    )
                    ok_button = icon.find_element(By.XPATH, "./parent::button")
                except:
                    try:
                        # Approach 3: Find any button in the dropdown
                        ok_button = self.driver.find_element(By.CSS_SELECTOR, ".conv-new-tag-dd button.btn-default")
                    except:
                        pass
            
            if ok_button:
                # Scroll into view
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", ok_button)
                    _delay(0.3)
                except:
                    pass
                
                # Try regular click first, then JavaScript as fallback
                try:
                    ok_button.click()
                    _delay(0.5)
                except Exception:
                    try:
                        self.driver.execute_script("arguments[0].click();", ok_button)
                        _delay(0.5)
                    except Exception:
                        pass
                
                # Verify tag was actually added by checking if it appears in the conversation tags
                _delay(0.5)  # Wait for tag to be saved
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
                            break
                    
                    if not tag_added:
                        # Also check if dropdown closed (alternative verification)
                        try:
                            dropdown = self.driver.find_element(By.CSS_SELECTOR, "#add-tag-wrap")
                            if not dropdown.is_displayed() or "display: none" in (dropdown.get_attribute("style") or ""):
                                tag_added = True
                        except Exception:
                            pass
                except Exception:
                    pass
            # Don't raise - tag addition is not critical
        except Exception:
            pass
    
    def select_template(self, template_name: str):
        """Select template from dropdown (FreeScout Saved Replies)"""
        try:
            # Use short wait for template UI (button + dropdown) so failed selectors don't block 10s each
            quick_wait = WebDriverWait(self.driver, 2)
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
                    template_button = quick_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if template_button:
                        break
                except Exception:
                    continue
            
            if not template_button:
                raise Exception("Could not find template button.")
            
            template_button.click()
            # Wait for dropdown to open
            try:
                WebDriverWait(self.driver, 0.8).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".dropdown-menu.dropdown-saved-replies")))
            except Exception:
                try:
                    WebDriverWait(self.driver, 0.5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".dropdown-menu.dropdown-saved-replies")))
                    _delay(0.05)
                except Exception:
                    pass
            
            template_link_selectors = [
                f"//div[@class='dropdown-menu']//a[contains(text(), '{template_name}')]",
                f"//li/a[contains(text(), '{template_name}')]",
                f"a[data-value]:contains('{template_name}')"
            ]
            
            template_link = None
            short_wait = WebDriverWait(self.driver, 1.5)
            for selector in template_link_selectors:
                try:
                    if selector.startswith("//"):
                        template_link = short_wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        template_link = short_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    if template_link:
                        break
                except Exception:
                    continue
            
            if not template_link:
                try:
                    dropdown = self.driver.find_element(By.CSS_SELECTOR, ".dropdown-menu.dropdown-saved-replies")
                    template_link = dropdown.find_element(By.XPATH, f".//a[contains(text(), '{template_name}')]")
                except Exception:
                    raise Exception(f"Could not find template '{template_name}' in dropdown.")
            try:
                template_link = WebDriverWait(self.driver, 1).until(
                    EC.element_to_be_clickable(template_link)
                )
            except Exception:
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", template_link)
                    _delay(0.2)
                except Exception:
                    pass
            
            try:
                template_link.click()
            except Exception:
                try:
                    self.driver.execute_script("arguments[0].click();", template_link)
                except Exception as js_e:
                    raise Exception(f"Could not click template link: {js_e}")
            # Wait for template to load into editor (short timeout; returns as soon as content appears)
            body_selectors = [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable"
            ]
            body_wait = WebDriverWait(self.driver, 1.2)
            for selector in body_selectors:
                try:
                    body_element = body_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    try:
                        WebDriverWait(self.driver, 1).until(
                            lambda d: body_element.text and len(body_element.text.strip()) > 50
                        )
                    except Exception:
                        pass
                    break
                except Exception:
                    continue
            
        except Exception as e:
            raise Exception(f"Failed to select template '{template_name}': {e}")
    
    def extract_template_content(self) -> Tuple[str, str]:
        """
        Extract template content and subject
        
        Returns:
            Tuple of (subject, body) - subject may be empty if not found
        """
        try:
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
                    body_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if body_element:
                        break
                except Exception:
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
        contact_person = (sponsor_data.get("contact_person") or "").strip()
        company_name = sponsor_data.get("company_name", "")
        email_str = sponsor_data.get("email", "")
        if contact_person:
            contact_first_name = (contact_person.split()[0] or "").strip().capitalize()
        else:
            full_from_email, contact_first_name = _name_from_email(
                email_str.split(",")[0].strip() if email_str else ""
            )
            if full_from_email:
                contact_person = full_from_email

        # Collapse block-only placeholder so "Hello" and name stay on one line (no extra space below Hello)
        for ph in ("{%customer.firstName%}", "{% customer.firstName %}"):
            filled_body = re.sub(r'</p>\s*<p[^>]*>\s*' + re.escape(ph) + r'\s*</p>', ' ' + contact_first_name, filled_body, flags=re.IGNORECASE)
            filled_body = re.sub(r'</div>\s*<div[^>]*>\s*' + re.escape(ph) + r'\s*</div>', ' ' + contact_first_name, filled_body, flags=re.IGNORECASE)
        # FreeScout-style placeholder: {%customer.firstName%}
        filled_body = filled_body.replace("{%customer.firstName%}", contact_first_name)
        if "{% customer.firstName %}" in filled_body:
            filled_body = filled_body.replace("{% customer.firstName %}", contact_first_name)
        
        # Get product name - priority: Company Name first, then General Info
        product_name = company_name  # Use Company Name first (always prefer this)
        row_data = sponsor_data.get("row_data", {})
        # Only use General Info if Company Name is empty or not available
        if not product_name or product_name.strip() == "":
            if "General Info" in row_data:
                general_info = str(row_data["General Info"]).strip() if not pd.isna(row_data.get("General Info")) else ""
                if general_info:
                    product_name = general_info
        
        # Check if placeholders exist in body (check both plain text and HTML)
        # The apostrophe might be HTML encoded as &#39; or &apos;
        placeholder_variations = [
            "[Prospective Sponsor's Name]",  # Plain text
            "[Prospective Sponsor&#39;s Name]",  # HTML entity &#39;
            "[Prospective Sponsor&apos;s Name]",  # HTML entity &apos;
            "[Prospective Sponsor&rsquo;s Name]",  # HTML entity &rsquo;
        ]
        
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
                filled_body = filled_body.replace(placeholder_with_apos, contact_first_name)
            
            product_placeholder = f"[Customer Company{apostrophe_char}s product name]"
            if product_placeholder in filled_body:
                filled_body = filled_body.replace(product_placeholder, product_name)
        
        # Also try regex pattern for any apostrophe variation
        apostrophe_pattern = r'\[Prospective Sponsor[\'\"\u2018\u2019]s Name\]'
        if re.search(apostrophe_pattern, filled_body, re.IGNORECASE):
            filled_body = re.sub(apostrophe_pattern, contact_first_name, filled_body, flags=re.IGNORECASE)
        
        # Also try regex pattern for [Customer Company's product name] with any apostrophe
        product_apostrophe_pattern = r'\[Customer Company[\'\"\u2018\u2019]s product name\]'
        if re.search(product_apostrophe_pattern, filled_body, re.IGNORECASE):
            filled_body = re.sub(product_apostrophe_pattern, product_name, filled_body, flags=re.IGNORECASE)
        
        for placeholder, replacement in placeholder_patterns:
            # Simple string replacement - replace ALL occurrences
            if filled_body.count(placeholder) > 0:
                filled_body = filled_body.replace(placeholder, replacement)
            
            # Also replace HTML entity encoded versions
            # Brackets encoded
            placeholder_html_brackets = placeholder.replace("[", "&#91;").replace("]", "&#93;")
            count_html_brackets = filled_body.count(placeholder_html_brackets)
            if count_html_brackets > 0:
                filled_body = filled_body.replace(placeholder_html_brackets, replacement)
            
            # Apostrophe encoded as &#39;
            placeholder_html_39 = placeholder.replace("'", "&#39;")
            count_html_39 = filled_body.count(placeholder_html_39)
            if count_html_39 > 0:
                filled_body = filled_body.replace(placeholder_html_39, replacement)
            
            # Apostrophe encoded as &apos;
            placeholder_html_apos = placeholder.replace("'", "&apos;")
            count_html_apos = filled_body.count(placeholder_html_apos)
            if count_html_apos > 0:
                filled_body = filled_body.replace(placeholder_html_apos, replacement)
            
            # Both brackets and apostrophe encoded
            placeholder_html_full = placeholder.replace("[", "&#91;").replace("]", "&#93;").replace("'", "&#39;")
            count_html_full = filled_body.count(placeholder_html_full)
            if count_html_full > 0:
                filled_body = filled_body.replace(placeholder_html_full, replacement)
        
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
            while placeholder_var in filled_body:
                filled_body = filled_body.replace(placeholder_var, contact_first_name)
        
        # Also try regex for any remaining variations (spaces, different encoding, any apostrophe)
        # Match any apostrophe character: ' ' ' or HTML entities
        pattern = r'\[Prospective\s+Sponsor[\'\"\u2018\u2019&#\w;]*s\s+Name\]'
        matches = re.findall(pattern, filled_body, re.IGNORECASE)
        if matches:
            filled_body = re.sub(pattern, contact_first_name, filled_body, flags=re.IGNORECASE)
        
        if "[Prospective Sponsor" in filled_body:
            idx = filled_body.find("[Prospective Sponsor")
            if idx >= 0:
                end_idx = filled_body.find("]", idx)
                if end_idx > idx:
                    remaining = filled_body[idx:end_idx+1]
                    if remaining not in placeholder_final_variations:
                        filled_body = filled_body.replace(remaining, contact_first_name)
        
        # Minimal cleanup: remove only truly empty elements, preserve ALL formatting including line breaks
        # IMPORTANT: Don't remove <div><br></div> or <span><br></span> - these create line breaks!
        # Only remove empty divs with nested empty spans (like <div><span></span></div> with no content)
        filled_body = re.sub(r'<div[^>]*>\s*<span[^>]*>\s*</span>\s*</div>', '', filled_body, flags=re.IGNORECASE)
        # Remove completely empty divs (no content, no <br>, no spans) - but preserve <div><br></div>
        filled_body = re.sub(r'<div[^>]*>\s*</div>', '', filled_body, flags=re.IGNORECASE)
        # Remove empty paragraphs and paragraphs that only contain a single <br> (extra blank line below Hello)
        filled_body = re.sub(r'<p[^>]*>\s*</p>', '', filled_body, flags=re.IGNORECASE)
        filled_body = re.sub(r'<p[^>]*>\s*<br\s*/?>\s*</p>', '', filled_body, flags=re.IGNORECASE)
        # Preserve ALL whitespace and line breaks - only trim trailing whitespace from entire string
        filled_body = filled_body.rstrip()
        
        return filled_body
    
    def fill_email_body(self, body: str, sponsor_data: Dict = None):
        """Fill the email body with content"""
        try:
            # Get replacement values if sponsor_data is provided
            contact_first_name = ""
            company_name = ""
            contact_person = ""
            product_name = ""
            if sponsor_data:
                contact_person = (sponsor_data.get("contact_person") or "").strip()
                email_str = sponsor_data.get("email", "")
                if contact_person:
                    contact_first_name = (contact_person.split()[0] or "").strip().capitalize()
                else:
                    full_from_email, contact_first_name = _name_from_email(
                        email_str.split(",")[0].strip() if email_str else ""
                    )
                    if full_from_email:
                        contact_person = full_from_email
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
                    body_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if body_element:
                        break
                except Exception:
                    continue
            
            if not body_element:
                raise Exception("Could not find email body editor.")
            
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
                        html = html.replace(/\\[Customer Company[''\\u2018\\u2019]s product name\\]/g, productName);
                        
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
                    _delay(0.1)
                    
                    # Update hidden textarea to ensure form submission uses correct content
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
                    
                    # Handle result and log debug info
                    if isinstance(result, dict):
                        replacement_success = result.get('replacementSuccess', False)
                    else:
                        replacement_success = result if isinstance(result, bool) else False
                except Exception:
                    # Fallback: Clear and use send_keys
                    body_element.click()
                    _delay(0.2)
                    body_element.send_keys(Keys.CONTROL + "a")
                    _delay(0.2)
                    body_element.send_keys(Keys.DELETE)
                    _delay(0.3)
                    # Convert HTML to text for fallback
                    # Simple HTML to text conversion
                    text_body = re.sub(r'<[^>]+>', '', body)  # Remove HTML tags
                    text_body = html.unescape(text_body)  # Decode HTML entities
                    body_element.send_keys(text_body)
                    _delay(0.5)
            
        except Exception as e:
            raise Exception(f"Failed to fill email body: {e}")
    
    def fill_subject_field(self, subject: str):
        """Fill the subject field"""
        try:
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
                    subject_field = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    if subject_field:
                        break
                except Exception:
                    continue
            
            if not subject_field:
                raise Exception("Could not find subject field.")
            
            subject_field.clear()
            subject_field.send_keys(subject)
            
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
                            # Get HTML to see what's actually there
                            html_content = body_element.get_attribute('innerHTML') or ''
                            if "[Prospective Sponsor" in html_content:
                                idx = html_content.find("[Prospective Sponsor")
                                snippet = html_content[max(0, idx-50):idx+80]
                    break
                except:
                    continue
            
            return preview
            
        except Exception as e:
            return {}
    
    def send_email(self):
        """Click the Send button"""
        try:
            # Before sending, ensure hidden textarea is synchronized
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
                        break
                except:
                    continue
            
            if not send_button:
                raise Exception("Could not find Send button.")
            
            # Send the email
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
        search_tab_handles = None
        template_name = sponsor_data.get('template_name', 'Unknown')
        valid_emails = self._parse_emails(sponsor_data.get('email', 'N/A'))
        emails_display = ', '.join(valid_emails) if valid_emails else sponsor_data.get('email', 'N/A')
        use_reply_flow = (
            REPLY_TO_THREAD_TEMPLATE_PATTERN
            and REPLY_TO_THREAD_TEMPLATE_PATTERN.lower() in template_name.lower()
        )
        if len(valid_emails) > 1:
            search_tab_handles = self._open_search_tabs_for_emails(valid_emails)
        try:
            print(f"\n📧 {emails_display} — {template_name}")
            if use_reply_flow:
                if not valid_emails:
                    print("No email to search for; skipping.")
                    return False
                any_sent = False
                main_handle = self.driver.current_window_handle
                for i, email in enumerate(valid_emails):
                    print(f"\n  → Thread for {email}")
                    # When we pre-opened search tabs (multiple emails), use that tab so we don't navigate main window to search
                    if search_tab_handles is not None and i < len(search_tab_handles):
                        if search_tab_handles[i] not in self.driver.window_handles:
                            print(f"  Search tab for {email} no longer available; skipping.")
                            continue
                        self.driver.switch_to.window(search_tab_handles[i])
                        _delay(0.3)
                        # Ensure this tab shows this email's search (tab may still be loading or show mailbox)
                        if FREESCOUT_URL:
                            base = FREESCOUT_URL.rstrip("/")
                            want_url = f"{base}/search?q={urllib.parse.quote(email, safe='')}"
                            try:
                                current = self.driver.current_url or ""
                                encoded_q = urllib.parse.quote(email, safe="")
                                if "search" not in current or encoded_q not in current:
                                    self.driver.get(want_url)
                                    _delay(1)
                            except Exception:
                                pass
                        thread_opened = self.open_thread_for_email(email, already_on_search_page=True)
                    else:
                        thread_opened = self.open_thread_for_email(email)
                    if not thread_opened:
                        print(f"  No thread found for {email}; skipping.")
                        continue
                    self.open_reply_form()
                    try:
                        self.select_template(template_name)
                        subject, body = self.extract_template_content()
                        body_selectors = [
                            "div.note-editable[contenteditable='true']",
                            "div[contenteditable='true']",
                            ".note-editable",
                        ]
                        for selector in body_selectors:
                            try:
                                body_el = self.driver.find_element(By.CSS_SELECTOR, selector)
                                self.driver.execute_script(
                                    "arguments[0].innerHTML = ''; arguments[0].textContent = '';", body_el
                                )
                                break
                            except Exception:
                                continue
                        filled_body = self.fill_placeholders(body, sponsor_data)
                        self.focus_reply_editor_and_fill(filled_body)
                    except Exception as e:
                        print(f"  Reply form / template failed for {email}: {e}")
                        continue
                    if confirm_before_send:
                        print("\n" + "="*80)
                        print(f"REPLY PREVIEW (thread: {email}):")
                        print("="*80)
                        print(f"Body:\n{filled_body[:500]}")
                        if len(filled_body) > 500:
                            print("... (truncated)")
                        print("="*80)
                        confirmation = input("\nSend this reply? (y/n/skip): ").strip().lower()
                        if confirmation == "n":
                            print("Reply cancelled by user.")
                            break
                        if confirmation == "skip":
                            print("Skipping this reply.")
                            continue
                    self.click_send_reply()
                    print(f"  ✓ Reply sent to thread: {email}")
                    any_sent = True
                    conv_handle = self.driver.current_window_handle
                    if conv_handle != main_handle and main_handle in self.driver.window_handles:
                        self.driver.switch_to.window(main_handle)
                    if conv_handle in self.driver.window_handles and conv_handle != main_handle:
                        self.driver.switch_to.window(conv_handle)
                        self.driver.close()
                    if main_handle in self.driver.window_handles:
                        self.driver.switch_to.window(main_handle)
                print(f"  Template used: {template_name}")
                return any_sent
            # New conversation flow
            self.click_new_conversation()
            self.fill_to_field(sponsor_data["email"])
            self.select_template(template_name)
            subject, body = self.extract_template_content()
            body_selectors = [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable",
            ]
            for selector in body_selectors:
                try:
                    body_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    self.driver.execute_script(
                        "arguments[0].innerHTML = ''; arguments[0].textContent = '';", body_element
                    )
                    break
                except Exception:
                    continue
            filled_body = self.fill_placeholders(body, sponsor_data)
            self.fill_email_body(filled_body, sponsor_data)
            if subject:
                self.fill_subject_field(subject)
            if confirm_before_send:
                preview = self.get_email_preview()
                parsed = self._parse_emails(sponsor_data.get("email", ""))
                to_display = ", ".join(parsed) if parsed else sponsor_data.get("email", "")
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
                if confirmation == "n":
                    print("Email cancelled by user.")
                    return False
                if confirmation == "skip":
                    print("Skipping this email.")
                    return False
            self.send_email()
            email_display = ", ".join(valid_emails) if valid_emails else sponsor_data.get("email", "")
            print(f"✓ Email sent successfully to: {email_display}")
            print(f"  Template used: {template_name}")
            return True
            
        except Exception as e:
            email_display = ", ".join(self._parse_emails(sponsor_data.get("email", ""))) or sponsor_data.get("email", "")
            print(f"✗ Error sending email to {email_display}: {e}")
            return False
        finally:
            if search_tab_handles is not None:
                self._close_search_tabs(search_tab_handles)
    
    def close(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()

