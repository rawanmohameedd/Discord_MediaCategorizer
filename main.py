import os
import datetime
import pandas as pd
import discord
from discord.ext import commands
import asyncio
import re

# Discord bot setup
TOKEN = os.getenv('TOKEN') 
intents = discord.Intents.default()
intents.message_content = True

# Note: You need to enable "Message Content Intent" in Discord Developer Portal
# Go to https://discord.com/developers/applications/ -> Your Bot -> Bot -> Message Content Intent
bot = commands.Bot(command_prefix='!', intents=intents)

# Where to save the Excel file
EXCEL_FILE = 'discord_links.xlsx'

# URL pattern to match
URL_PATTERN = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'

# Dictionary to store link data
link_data = {
    'url': [],
    'channel': [],
    'author': [],
    'date': [],
    'message_content': []
}

@bot.event
async def on_ready():
    print(f'Bot is ready: {bot.user.name}')
    
    # Choose which channel to scan
    target_channel_id = int(input("Enter the channel ID to scan: "))
    target_channel = bot.get_channel(target_channel_id)
    
    if not target_channel:
        print(f"Couldn't find channel with ID {target_channel_id}")
        await bot.close()
        return
    
    print(f"Scanning channel: {target_channel.name}")
    
    # Get messages from the channel
    async for message in target_channel.history(limit=1000):  # Adjust limit as needed
        # Find URLs in message content
        urls = re.findall(URL_PATTERN, message.content)
        for url in urls:
            # Store link information
            link_data['url'].append(url)
            link_data['channel'].append(target_channel.name)
            link_data['author'].append(f"{message.author.name}#{message.author.discriminator}")
            link_data['date'].append(message.created_at)
            link_data['message_content'].append(message.content)
    
    # Convert to DataFrame and sort by date
    df = pd.DataFrame(link_data)
    df = df.sort_values(by='date')
    
    # Create Excel file
    df.to_excel(EXCEL_FILE, index=False)
    print(f"Excel file created: {EXCEL_FILE}")
    
    await bot.close()

# Run the bot
try:
    asyncio.run(bot.start(TOKEN))
except KeyboardInterrupt:
    print("Bot stopped manually")