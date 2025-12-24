# Windows 11 Notification to Slack Bot

This bot monitors Windows 11 desktop notifications and forwards them to a Slack workspace.

## Features

- Reads Windows 11 notifications directly from the notification database
- Sends notifications to Slack workspace
- Configurable via environment variables
- No UI required - runs as a background script

## Prerequisites

- Windows 11
- Python 3.8 or higher
- A Slack workspace with a bot token

## Setup Guide

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Create a Slack Bot

1. Go to https://api.slack.com/apps
2. Click **"Create New App"** → **"From scratch"**
3. Name your app (e.g., "Windows Notification Bot")
4. Select your workspace
5. Click **"Create App"**
6. Go to **"OAuth & Permissions"** in the left sidebar
7. Under **"Bot Token Scopes"**, click **"Add an OAuth Scope"**
8. Add the scope: `chat:write`
9. Scroll up and click **"Install to Workspace"**
10. Click **"Allow"** to authorize
11. Copy the **"Bot User OAuth Token"** (starts with `xoxb-`)

### 3. Configure Environment Variables

Edit the `.env` file:

```
SLACK_BOT_TOKEN=xoxb-your-token-here
SLACK_CHANNEL=#notifications
```

**To find your Slack Channel ID:**
- Right-click the channel name in Slack → **"View channel details"**
- The Channel ID is in the URL after `/archives/` (starts with `C`)
- Or use the channel name with `#` prefix: `#channel-name`

**Important:** 
- Make sure your bot is **invited to the channel** (type `/invite @YourBotName` in the channel)
- Or add the `chat:write.public` scope in Slack app settings to post to channels without being invited
- If you get `channel_not_found` error, the channel ID is wrong or the bot isn't in that channel

**Optional: Set Machine Name**
- Add `MACHINE_NAME=Your-PC-Name` to `.env` to identify which PC sent the notification
- If not set, it defaults to the computer's hostname
- Useful when running on multiple machines (see below)

### 4. Run the Bot

```bash
python app.py
```

The bot will:
- Check for new notifications every 5 seconds
- Send them to your Slack channel
- Run until you press Ctrl+C

## Running on Multiple Machines

**Yes! You can run this bot on multiple PCs and have all notifications sent to the same Slack channel.**

### Setup for Multiple Machines:

1. **On each PC:**
   - Copy the entire project folder to each machine
   - Install dependencies: `pip install -r requirements.txt`
   - Edit `.env` file on each PC:
     - Use the **same** `SLACK_BOT_TOKEN` (same bot, same workspace)
     - Use the **same** `SLACK_CHANNEL` (all PCs send to one channel)
     - Set a **unique** `MACHINE_NAME` for each PC (optional but recommended):
       ```
       # On PC 1:
       MACHINE_NAME=Office-PC
       
       # On PC 2:
       MACHINE_NAME=Home-Laptop
       
       # On PC 3:
       MACHINE_NAME=Server-01
       ```
     - If `MACHINE_NAME` is not set, it defaults to the computer's hostname

2. **Run the bot on each PC:**
   ```bash
   python app.py
   ```

3. **All notifications will appear in the same Slack channel** with the PC name included, so you can identify which machine sent each notification.

**Example notification format:**
```
*New Message*
From: John

🖥️ PC: Office-PC
App: Slack
Time: 2024-12-24T10:30:00
```

**Benefits:**
- Centralized notification monitoring from all your PCs
- Easy to identify which machine sent which notification
- Each PC runs independently - if one goes offline, others keep working
- All notifications in one place for easy monitoring

## How It Works

The bot reads notifications directly from the Windows notification database (`wpndatabase.db`). It:
1. Connects to the SQLite database where Windows stores notifications
2. Queries for toast notifications (not tiles)
3. Extracts title and body from the XML payload
4. Sends new notifications to Slack
5. Tracks processed notifications to avoid duplicates

## Troubleshooting Guide

### Bot doesn't capture notifications

**Check 1: Windows Notification Settings**
- Press `Win + I` → **System** → **Notifications**
- Make sure **"Get notifications from apps and other senders"** is ON
- Ensure notifications are enabled for the apps you want to monitor

**Check 2: Database Access**
- The bot reads from: `%LOCALAPPDATA%\Microsoft\Windows\Notifications\wpndatabase.db`
- If the database is locked (Windows is using it), the bot falls back to PowerShell method
- This is normal and the bot will still work

**Check 3: Slack Configuration**
- Verify your bot token is correct in `.env`
- Make sure the bot is invited to the channel
- Check that the channel name/ID is correct

**Check 4: Test Slack Connection**
- The bot will show errors if Slack connection fails
- Check your internet connection
- Verify the bot token hasn't expired

### Notifications appear but aren't captured

- Windows may not store all notifications in the database immediately
- Wait a few seconds - the bot checks every 5 seconds
- Some notifications (like system notifications) may not be stored
- The bot only captures "toast" notifications, not tiles or badges

### Database is locked error

- This is normal - Windows locks the database when it's in use
- The bot automatically falls back to PowerShell method
- Both methods should work, but database method is preferred

## Notes

- The bot checks for new notifications every 5 seconds
- Notifications are tracked by their unique ID to prevent duplicates
- The bot runs continuously until stopped with Ctrl+C
- Only toast notifications are captured (not tiles or badges)

