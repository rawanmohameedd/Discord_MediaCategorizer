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

bot = commands.Bot(command_prefix='!', intents=intents)

# Where to save the Excel file
EXCEL_FILE = 'discord_links.xlsx'

# URL pattern to match
URL_PATTERN = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'

# Dictionary to store link data
link_data = {
    'session': [],
    'url': [],
    'thread_title': [],  
    'author': [],
    'date': [],
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
    
    # Check if channel is a forum channel
    if isinstance(target_channel, discord.ForumChannel):
        print("Detected forum channel. Processing forum threads...")
        await process_forum_channel(target_channel)
    else:
        # Check bot permissions for regular text channels
        if not target_channel.permissions_for(bot.guilds[0].me).read_messages:
            print("Bot doesn't have permission to read messages in this channel")
            await bot.close()
            return
            
        if not target_channel.permissions_for(bot.guilds[0].me).read_message_history:
            print("Bot doesn't have permission to read message history in this channel")
            await bot.close()
            return
        
        # Process regular text channel
        await process_text_channel(target_channel)
    
    # Convert to DataFrame
    df = pd.DataFrame(link_data)
    
    # Convert timezone-aware datetime objects to timezone-naive
    if 'date' in df.columns and not df.empty:
        df['date'] = df['date'].apply(lambda dt: dt.replace(tzinfo=None) if dt is not None else dt)
    
    # Sort by date
    df = df.sort_values(by='date')
    
    # Fix session numbers based on weekly grouping
    if not df.empty:
        df['date_only'] = df['date'].dt.date
        
        # Start date for sessions
        start_date = datetime.date(2025, 2, 1)
        
        # Calculate session number per week
        df['session'] = df['date_only'].apply(
            lambda d: ((d - start_date).days // 7) + 1 if d >= start_date else 0
        )
        
        # Remove temporary column
        df = df.drop('date_only', axis=1)
    
    # Create Excel file
    try:
        df.to_excel(EXCEL_FILE, index=False)
        print(f"Excel file created: {EXCEL_FILE}")
    except PermissionError:
        timestamp = int(datetime.datetime.now().timestamp())
        new_filename = f"discord_links_{timestamp}.xlsx"
        df.to_excel(new_filename, index=False)
        print(f"Excel file created with alternative name: {new_filename}")
    
    await bot.close()

async def process_forum_channel(forum_channel):
    try:
        print(f"Fetching threads from forum channel: {forum_channel.name}")
        
        active_threads = forum_channel.threads
        print(f"Found {len(active_threads)} active threads")
        
        archived_threads = []
        try:
            async for thread in forum_channel.archived_threads(limit=None):
                archived_threads.append(thread)
            print(f"Found {len(archived_threads)} archived threads")
        except Exception as e:
            print(f"Error fetching archived threads: {str(e)}")
        
        all_threads = list(active_threads) + archived_threads
        print(f"Total threads to process: {len(all_threads)}")
        
        for thread in all_threads:
            print(f"Processing thread: {thread.name}")
            try:
                async for message in thread.history(limit=None):
                    print(f"Message from: {message.author} (ID: {message.author.id})")
                    urls = re.findall(URL_PATTERN, message.content)
                    for url in urls:
                        print(f"Found URL: {url} from author: {message.author}")
                        link_data['session'].append(0)  # Placeholder
                        link_data['url'].append(url)
                        link_data['thread_title'].append(thread.name)
                        link_data['author'].append(str(message.author))
                        link_data['date'].append(message.created_at)
            except Exception as e:
                print(f"Error processing thread {thread.name}: {str(e)}")
                continue
    except Exception as e:
        print(f"Error while processing forum channel: {str(e)}")

async def process_text_channel(text_channel):
    try:
        print("Starting to fetch message history...")
        async for message in text_channel.history(limit=None):
            print(f"Message from: {message.author} (ID: {message.author.id})")
            urls = re.findall(URL_PATTERN, message.content)
            for url in urls:
                print(f"Found URL: {url} from author: {message.author}")
                link_data['session'].append(0)  # Placeholder
                link_data['url'].append(url)
                link_data['thread_title'].append("N/A")
                link_data['author'].append(str(message.author))
                link_data['date'].append(message.created_at)
    except Exception as e:
        print(f"Error while fetching messages: {str(e)}")

# Run the bot
try:
    bot.run(TOKEN)
except KeyboardInterrupt:
    print("Bot stopped manually")
except Exception as e:
    print(f"Error running bot: {str(e)}")
