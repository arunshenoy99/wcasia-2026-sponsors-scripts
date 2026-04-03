#!/usr/bin/env python3
"""
CLI: explore FreeScout conversation page — print element info for Reply, editor, Send (selector debugging).

Standalone; uses freescout_automation login. Full context: Reachout/README.md (FreeScout selector section).

Usage:
  python explore_freescout_selectors.py [search_email]
  If search_email is omitted, you'll be prompted. Example: user@example.com

Requires .env with FREESCOUT_URL, FREESCOUT_EMAIL, FREESCOUT_PASSWORD.
Browser runs visible (not headless). Press Enter at the end to close.
"""
import sys
import time

# Force non-headless for exploration
import os
os.environ.setdefault("HEADLESS_MODE", "false")

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from config import FREESCOUT_URL, FREESCOUT_EMAIL, FREESCOUT_PASSWORD

if not FREESCOUT_URL or not FREESCOUT_EMAIL or not FREESCOUT_PASSWORD:
    print("Missing config. Set FREESCOUT_URL, FREESCOUT_EMAIL, FREESCOUT_PASSWORD in .env")
    sys.exit(1)

from freescout_automation import FreeScoutAutomation
import urllib.parse


def describe(el):
    try:
        tag = el.tag_name
        c = (el.get_attribute("class") or "").strip()
        aid = el.get_attribute("id") or ""
        aria = el.get_attribute("aria-label") or ""
        role = el.get_attribute("role") or ""
        name = el.get_attribute("name") or ""
        visible = el.is_displayed()
        return f"<{tag}> class={c!r} id={aid!r} aria-label={aria!r} role={role!r} name={name!r} visible={visible}"
    except Exception as e:
        return f"(error: {e})"


def find_and_report(driver, name, selectors, by=By.CSS_SELECTOR):
    print(f"\n--- {name} ---")
    for sel in selectors:
        try:
            els = driver.find_elements(by, sel)
            if els:
                print(f"  Selector: {sel!r} -> {len(els)} element(s)")
                for i, el in enumerate(els[:5]):
                    print(f"    [{i}] {describe(el)}")
            else:
                print(f"  Selector: {sel!r} -> 0 elements")
        except Exception as e:
            print(f"  Selector: {sel!r} -> error: {e}")


def main():
    search_email = (sys.argv[1] if len(sys.argv) > 1 else "").strip()
    if not search_email or "@" not in search_email:
        search_email = input("Enter an email to search for (e.g. from a conversation): ").strip()
    if not search_email or "@" not in search_email:
        print("No valid email given. Exiting.")
        sys.exit(1)

    print("Starting browser (visible)...")
    bot = FreeScoutAutomation()
    bot.setup_browser()
    bot.wait = WebDriverWait(bot.driver, 15)

    try:
        print("Logging in...")
        bot.login()
        time.sleep(2)

        base = FREESCOUT_URL.rstrip("/")
        url = f"{base}/search?q={urllib.parse.quote(search_email, safe='')}"
        print(f"Opening search: {url}")
        bot.driver.get(url)
        time.sleep(2)

        # Open first conversation: same widened selectors as automation; click (FreeScout opens new tab); switch to it
        first_conv_href = None
        first_link_el = None
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
        for selector in first_conversation_selectors:
            try:
                links = bot.driver.find_elements(By.CSS_SELECTOR, selector)
                for link in links:
                    href = (link.get_attribute("href") or "").strip()
                    if not bot._is_conversation_link(href):
                        continue
                    first_conv_href = href
                    first_link_el = link
                    break
                if first_link_el is not None:
                    break
            except Exception:
                continue
        if not first_link_el or not first_conv_href:
            print("  No conversation link found on search page.")
            input("Press Enter to close browser...")
            bot.close()
            sys.exit(1)
        print(f"  First conversation URL: {first_conv_href}")
        handles_before = set(bot.driver.window_handles)
        try:
            first_link_el.click()
        except Exception as e:
            print(f"  Click failed: {e}; opening via driver.get")
            bot.driver.get(first_conv_href)
        time.sleep(2)
        bot._switch_to_thread_tab_if_opened(handles_before)
        if "search" in bot.driver.current_url:
            print("  Still on search page; opening conversation via driver.get(...)")
            bot.driver.get(first_conv_href)
            time.sleep(2)

        print(f"  Current URL: {bot.driver.current_url}")

        # Report on conversation page (reply form may be closed)
        find_and_report(
            bot.driver,
            "Reply button (opens reply form)",
            [
                "span.conv-reply.conv-action[aria-label='Reply']",
                "span.conv-reply[aria-label='Reply']",
                "[aria-label='Reply']",
                "span.conv-reply.conv-action.glyphicon-share-alt",
                ".conv-reply",
                "span.glyphicon-share-alt",
            ],
        )
        find_and_report(
            bot.driver,
            "Reply editor (contenteditable / textarea)",
            [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable",
                ".reply-form div[contenteditable='true']",
                "#reply-body",
                "textarea[name*='body']",
                "textarea.reply-body",
            ],
        )
        find_and_report(
            bot.driver,
            "Send / Submit button (sends reply)",
            [
                "button[type='submit']",
                "button.btn-send-text",
                "button.btn-reply-submit",
                "button.btn-primary",
                ".reply-form button[type='submit']",
                ".conv-reply-form button[type='submit']",
                "button.reply-btn",
            ],
        )

        # Click Reply to open form, then report again
        print("\n--- Clicking Reply to open reply form... ---")
        for selector in ["span.conv-reply[aria-label='Reply']", "[aria-label='Reply']", "span.conv-reply"]:
            try:
                btn = bot.driver.find_element(By.CSS_SELECTOR, selector)
                bot.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(0.3)
                try:
                    btn.click()
                except Exception:
                    bot.driver.execute_script("arguments[0].click();", btn)
                time.sleep(2)
                print(f"  Clicked: {selector}")
                break
            except Exception as e:
                print(f"  Failed {selector}: {e}")
        else:
            print("  Could not click Reply.")

        # Report again after opening reply form
        find_and_report(
            bot.driver,
            "Reply editor AFTER opening form",
            [
                "div.note-editable[contenteditable='true']",
                "div[contenteditable='true']",
                ".note-editable",
                ".reply-form div[contenteditable='true']",
                "#reply-body",
                "textarea[name*='body']",
            ],
        )
        find_and_report(
            bot.driver,
            "Send / Submit button AFTER opening form",
            [
                "button[type='submit']",
                "button.btn-send-text",
                "button.btn-reply-submit",
                ".reply-form button[type='submit']",
                ".conv-reply-form button[type='submit']",
            ],
        )

        print("\n" + "=" * 60)
        print("Done. Browser left open for inspection. Press Enter to close.")
        input()
    finally:
        bot.close()


if __name__ == "__main__":
    main()
