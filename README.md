# SEO Backlink Validator & Automation Tool

A comprehensive dual-solution tool designed to automate the verification of SEO backlinks, anchor texts, and HTTP response statuses (404, 403, etc.) directly from Google Sheets. 

The project evolved from a lightweight Apps Script checker to a robust C# application capable of bypassing modern bot-protection systems (Cloudflare, DDoS-Guard).

## 📁 Project Structure

### [v1_Fast_Http_Checker_JS](./v1_Fast_Http_Checker_JS)
A Google Apps Script (JavaScript) solution meant to run directly inside Google Sheets.
- **Pros:** Extremely fast, zero local setup, native Google Sheets UI integration.
- **Cons:** Blocked by aggressive anti-bot protection (returns 403/Captcha pages instead of verifying the link).

### [v2_AntiBot_Playwright_CSharp](./v2_AntiBot_Playwright_CSharp)
An advanced desktop application built with C# and Microsoft Playwright.
- Connects to **Octo Browser** via CDP (Chrome DevTools Protocol) to simulate a real user fingerprint.
- Bypasses Cloudflare and other bot-checks.
- Uses **Google Sheets API** to read targets and write verification results back to the cloud in real-time.

### [v3_Link_Collector_GAS](./v3_Link_Collector_GAS)
A utility Google Apps Script designed to aggregate and transpose scattered backlink placement data from date-based columns into a clean, structured report sheet.

## 🛠 Tech Stack
- **Languages:** C# (.NET), Vanilla JavaScript
- **Automation:** Microsoft Playwright, Octo Browser CDP
- **APIs:** Google Sheets API, Google Apps Script
