# Apps-Script-Gmail-Automation-Nov-2025
Featured on Medium

# 🧠 Gmail Assistant — Google Apps Script + ChatGPT

**Automate your Gmail cleanup and reclaim focus using a simple, no-cost automation built with Google Apps Script and ChatGPT.**  
This project shows you how to create a personal Gmail assistant that automatically marks certain emails as read, archives them, and logs actions to Google Sheets — all without API keys, subscriptions, or third-party services.

Note: This project does not use API keys or require a subscription. It does not read your emails for sentiment or any "fancy LLM stuff" ; it's just using a simple string CONTAINS match on Email Body and Email Author.

---

## 🚀 Features

- ✅ Auto-archive and mark-as-read emails based on your custom rules  
- ✅ Google Sheets "Rule Engine" — edit filters visually without coding  
- ✅ Automatic execution via timed triggers (every 30 min or 1 hr)  
- ✅ Built-in logging for debugging and transparency  
- ✅ 100 % free — uses your existing Google account  
- ✅ Expandable: add daily summaries, sentiment analysis, SMS alerts, etc.

---

## 🧩 How It Works

1. **Google Sheet as Backend**
   - Three tabs:
     - `Rules` — define what emails to archive (by sender, subject, or text)
     - `Email Actions` — logs of what the script did and why
     - `Saved Prompt` — original ChatGPT or configuration prompt

2. **Apps Script as Automation Engine**
   - The main function `autoArchiveLast7Days()` scans your Gmail inbox
   - It checks each message against your rules
   - If a match is found → marks as read and archives the thread
   - Logs the action into the Sheet

3. **Trigger for Continuous Operation**
   - Uses Google Apps Script’s built-in triggers (`Clock` icon)
   - Runs on your chosen schedule, keeping your inbox clean

---

## 🧰 Setup Guide

1. Create a new **Google Sheet** and name it something like `Gmail Assistant`.
2. Copy the example structure:
   - `Rules` — add a few test rows (e.g. “unsubscribe”, “LinkedIn Job Alerts”)
   - `Email Actions` — leave blank
3. Open **Extensions → Apps Script**, paste in the code from `/src/Code.gs`, and save.
4. Authorize the script when prompted.
5. Test the function manually:
   ```javascript
   autoArchiveLast7Days();
