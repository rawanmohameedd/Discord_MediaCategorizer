﻿---
This bot collects and organizes all links shared in both regular text channels and forum threads, grouping them into weekly sessions and exporting the data to an Excel file.

---

## 🚀 Features

- Supports both **text channels** and **forum threads**
- Automatically organizes links by **weekly sessions** (starting from **February 1, 2025**) 
- Exports data to Excel including:
  - Weekly **session number**
  - Shared **URL**
  - **Thread title** (if forum) or "N/A" (if text channel)
  - **Author** of the message
  - **Date** of sharing

---

## 🛠️ Setup

1. **Install dependencies**:
   ```bash
   pip install discord.py pandas openpyxl
   ```

2. **Create a `.env` file** with your bot token:
   ```
   TOKEN=your_bot_token_here
   ```

3. **Enable necessary permissions**:
   - Go to [Discord Developer Portal](https://discord.com/developers/applications/)
   - Select your bot → **Bot** tab
   - Enable **Message Content Intent**
   - Generate an invite link with these permissions:
     - Read Messages / View Channels
     - Read Message History
     - View Forum Channels

---

## ▶️ Usage

1. Run the bot:
   ```bash
   python main.py
   ```

2. When prompted, **enter the Channel ID** you want to scan:
   - To get a Channel ID:
     - Enable **Developer Mode** in Discord (User Settings → Advanced)
     - Right-click a channel → **Copy ID**

3. The bot will:
   - Scan all messages (no limit)
   - Extract links and metadata
   - Export to `discord_links.xlsx`
   - If the file is locked or in use, it creates a **timestamped alternative**

---

## 📄 Output Format

The exported Excel file contains:

| Column         | Description                                  |
|----------------|----------------------------------------------|
| `session`      | Weekly session number (starting 2025-02-01)  |
| `url`          | The collected link                           |
| `thread_title` | Name of forum thread or "N/A" for text channels |
| `author`       | Username who posted the link                 |
| `date`         | Timestamp of the message                     |

---

## ⚠️ Error Handling

- **Permission errors** are reported and handled gracefully
- **Excel file errors** (e.g., file locked) result in a fallback file with a timestamp
- **Message and thread errors** are logged but don’t stop the bot

---

## 📝 Notes

- All accessible messages are scanned (no date or message count limit)
- Forum threads include both **active** and **archived** threads
- Sessions are calculated based on message dates and grouped by week

