# Steps to achieve
### Step 1: Create the Telegram Bot (On your phone)
Open the Telegram app and search for @BotFather (it has a blue verification tick).

Start a chat and send the command /newbot.

Give it a name (e.g., "Sara's Study Manager").

Give it a username ending in bot (e.g., sara_study_bot).

BotFather will give you an HTTP API Token (a long string of letters and numbers). Copy this token — you need it for the code.

### Step 2: Set up the Python Environment
On your computer, open your terminal/command prompt and install the required library. (Make sure you are using Python 3.7+).

Bash
pip install python-telegram-bot

### Step 3: Create the JSON State File
In your project folder, create a file named study_state.json. This will hold your alternating schedules and the running backlog.

Paste this exactly into the file:
```
Bash
{
    "current_cycle_day": 1,
    "backlog": [],
    "templates": {
        "1": ["Leetcode", "React", "Java", "SQL Practice", "Read Book", "Weekly Task/ LinkedIn Post"],
        "2": ["DBMS", "Open Source", "Python", "Sub 8", "Artwork/ Phone Cleanup", "Improve Project"]
    },
    "history": []
}
```

### Step 4: The Python Bot Code
Create a file named bot.py in the same folder.

```
Bash
import json
import os
import pandas as pd
from io import BytesIO
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
STATE_FILE = "study_state.json"

# --- HELPER FUNCTIONS ---
def load_state():
    with open(STATE_FILE, 'r') as f:
        return json.load(f)

def save_state(data):
    with open(STATE_FILE, 'w') as f:
        json.dump(data, f, indent=4)

# Global temporary storage
completed_today = []
current_daily_list = []

# --- BOT COMMANDS ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Test Mode Active!\n"
        "Send /morning to get your checklist.\n"
        "Send /evening to complete a day and auto-increment the Day counter.\n"
        "Send /calendar to get your updated multi-row Excel report."
    )

async def morning_checklist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global completed_today, current_daily_list
    completed_today = [] # Reset completions
    
    data = load_state()
    backlog = data['backlog']
    cycle = str(data['current_cycle_day'])
    template = data['templates'][cycle]
    
    current_daily_list = list(dict.fromkeys(backlog + template))
    
    keyboard = []
    for task in current_daily_list:
        keyboard.append([InlineKeyboardButton(f"⬜ {task}", callback_data=task)])
        
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.message.reply_text(
        f"🌅 **Morning! Today is Cycle Day {cycle}.**\nHere is your list:", 
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )

async def handle_task_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer() 
    
    task_clicked = query.data
    global completed_today
    
    if task_clicked in completed_today:
        completed_today.remove(task_clicked)
    else:
        completed_today.append(task_clicked)
    
    keyboard = []
    for task in current_daily_list:
        if task in completed_today:
            keyboard.append([InlineKeyboardButton(f"✅ {task}", callback_data=task)])
        else:
            keyboard.append([InlineKeyboardButton(f"⬜ {task}", callback_data=task)])
            
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_reply_markup(reply_markup=reply_markup)

async def evening_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global completed_today, current_daily_list
    data = load_state()
    
    # Make sure history is a list
    if "history" not in data or not isinstance(data["history"], list):
        data["history"] = []
        
    # Calculate the next Day counter sequence number
    next_day_num = len(data["history"]) + 1
    day_label = f"Day {next_day_num}"
    
    # Identify missed tasks
    missed_tasks = [task for task in current_daily_list if task not in completed_today]
    today_date = datetime.now().strftime("%Y-%m-%d")
    
    # Append the new log entry as an independent row record
    data["history"].append({
        "day_label": day_label,
        "date": today_date,
        "tasks": completed_today.copy()
    })
    
    # Update alternate cycle state and backlog
    data['backlog'] = missed_tasks
    old_cycle = data['current_cycle_day']
    data['current_cycle_day'] = 2 if old_cycle == 1 else 1
    
    save_state(data)
    
    await update.message.reply_text(
        f"🌙 **{day_label} Review complete!**\n"
        f"📅 Logged under date: {today_date}\n"
        f"✅ Completed: {len(completed_today)}\n"
        f"⏭️ Shifted to tomorrow's backlog: {len(missed_tasks)}\n"
        f"Next setup will be Cycle Day {data['current_cycle_day']}."
    )

async def send_calendar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_state()
    history = data.get("history", [])
    
    if not history:
        await update.message.reply_text("Your history is empty! Complete at least one day using /evening first.")
        return
        
    # Build clean list rows for the dataframe
    formatted_rows = []
    for entry in history:
        tasks_str = ", ".join(entry["tasks"]) if entry["tasks"] else "No tasks completed"
        formatted_rows.append({
            "Day": entry["day_label"],
            "Date": entry["date"],
            "Completed Tasks": tasks_str
        })
        
    # Create DataFrame
    df = pd.DataFrame(formatted_rows)
    
    # Build the spreadsheet output in memory
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Study Progress')
        
        # Apply layout adjustments
        worksheet = writer.sheets['Study Progress']
        worksheet.column_dimensions['A'].width = 12  # Day column
        worksheet.column_dimensions['B'].width = 15  # Date column
        worksheet.column_dimensions['C'].width = 60  # Tasks column
        
    buffer.seek(0)
    buffer.name = "Study_Progress_Report.xlsx"
    
    await update.message.reply_document(
        document=buffer, 
        filename="Study_Progress_Report.xlsx", 
        caption="📅 Here is your testing history spreadsheet with sequential days!"
    )

# --- MAIN EXECUTION ---
if __name__ == '__main__':
    print("Bot is starting...")
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("morning", morning_checklist))
    app.add_handler(CommandHandler("evening", evening_review))
    app.add_handler(CommandHandler("calendar", send_calendar))
    app.add_handler(CallbackQueryHandler(handle_task_click))

    print("Bot is running! Go to Telegram to use it.")
    app.run_polling()
```

### Step 5: How to Use It
Run the script on your computer (python bot.py). Leave the terminal open.

Open Telegram on your phone and go to your bot's chat.

Type /morning. The bot will instantly reply with a list of buttons for Sub 1-4 and Task 1-2.

Throughout the day, tap the buttons on your phone. They will magically change from ⬜ to ✅.

Before bed, type /evening. The bot will look at what is still ⬜, save it to your study_state.json backlog, and set tomorrow to Cycle Day 2.

The next day, type /morning, and you will see your new backlog mixed with Sub 5-8!

/calendar to download excel file of records