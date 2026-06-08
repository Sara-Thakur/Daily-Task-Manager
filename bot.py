import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
import time
import urllib.request
import pandas as pd
from io import BytesIO
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes

# --- ENVIRONMENT CONFIGURATION ---
TOKEN = os.environ.get("TELEGRAM_TOKEN")
STATE_FILE = "study_state.json"

# Check immediately on startup if the token is present
if not TOKEN:
    print("❌ ERROR: 'TELEGRAM_TOKEN' environment variable is missing!")
    print("Please set it in your Render environment variables panel.")

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
        "👋 Test Mode Active!\n\n"
        "🌅 /morning - Get your current checklist\n"
        "🌙 /evening - Complete the day, save to excel, and progress\n"
        "⬅️ /back - Revert the last /evening log to modify your choices\n"
        "📅 /calendar - Download your Excel report"
    )

async def morning_checklist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global completed_today, current_daily_list
    
    data = load_state()
    cycle = str(data['current_cycle_day'])
    template = data['templates'][cycle]
    backlog = data.get('backlog', [])
    
    # Generate current working list combining backlog and template
    current_daily_list = list(dict.fromkeys(backlog + template))
    
    keyboard = []
    for task in current_daily_list:
        if task in completed_today:
            keyboard.append([InlineKeyboardButton(f"✅ {task}", callback_data=task)])
        else:
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
    
    if "history" not in data or not isinstance(data["history"], list):
        data["history"] = []
        
    next_day_num = len(data["history"]) + 1
    day_label = f"Day {next_day_num}"
    
    missed_tasks = [task for task in current_daily_list if task not in completed_today]
    today_date = datetime.now().strftime("%Y-%m-%d")
    
    # Log entry
    data["history"].append({
        "day_label": day_label,
        "date": today_date,
        "tasks": completed_today.copy(),
        "logged_cycle_day": data['current_cycle_day']  # Saved to cleanly restore on /back
    })
    
    data['backlog'] = missed_tasks
    old_cycle = data['current_cycle_day']
    data['current_cycle_day'] = 2 if old_cycle == 1 else 1
    
    save_state(data)
    
    await update.message.reply_text(
        f"🌙 **{day_label} Review complete!**\n"
        f"📅 Logged under date: {today_date}\n"
        f"✅ Completed: {len(completed_today)}\n"
        f"⏭️ Shifted to tomorrow's backlog: {len(missed_tasks)}\n"
        f"Next setup will be Cycle Day {data['current_cycle_day']}.\n\n"
        f"💡 *Made a mistake? Send /back to undo this entry and fix it.*",
        parse_mode="Markdown"
    )

async def back_to_previous_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global completed_today, current_daily_list
    data = load_state()
    history = data.get("history", [])
    
    if not history:
        await update.message.reply_text("❌ No logged history found to jump back to!")
        return
        
    # Remove last entry to roll back spreadsheet progress
    last_entry = history.pop()
    
    # Restore choices back into memory variables
    completed_today = last_entry["tasks"]
    
    # Reconstruct original template context
    restored_cycle = last_entry.get("logged_cycle_day", 1)
    data['current_cycle_day'] = restored_cycle
    
    # Roll back data structure updates
    template = data['templates'][str(restored_cycle)]
    backlog = data.get('backlog', [])
    current_daily_list = list(dict.fromkeys(backlog + template))
    
    # Save modified data back to json structure
    data["history"] = history
    save_state(data)
    
    # Construct interactive keyboard containing saved data states
    keyboard = []
    for task in current_daily_list:
        if task in completed_today:
            keyboard.append([InlineKeyboardButton(f"✅ {task}", callback_data=task)])
        else:
            keyboard.append([InlineKeyboardButton(f"⬜ {task}", callback_data=task)])
            
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.message.reply_text(
        f"🔄 **Reverted back to {last_entry['day_label']}!**\n"
        f"The spreadsheet row record has been cleared. Make your modifications below and use /evening when done.",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )

async def send_calendar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_state()
    history = data.get("history", [])
    
    if not history:
        await update.message.reply_text("Your history is empty! Complete at least one day using /evening first.")
        return
        
    formatted_rows = []
    for entry in history:
        tasks_str = ", ".join(entry["tasks"]) if entry["tasks"] else "No tasks completed"
        formatted_rows.append({
            "Day": entry["day_label"],
            "Date": entry["date"],
            "Completed Tasks": tasks_str
        })
        
    df = pd.DataFrame(formatted_rows)
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Study Progress')
        
        worksheet = writer.sheets['Study Progress']
        worksheet.column_dimensions['A'].width = 12
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 60
        
    buffer.seek(0)
    buffer.name = "Study_Progress_Report.xlsx"
    
    await update.message.reply_document(
        document=buffer, 
        filename="Study_Progress_Report.xlsx", 
        caption="📅 Here is your testing history spreadsheet with sequential days!"
    )

class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot is alive and holding the port open!")

def run_dummy_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(('0.0.0.0', port), DummyHandler)
    
    # Active Auto-Ping Keep Alive loop
    def keep_alive():
        time.sleep(10)
        while True:
            try:
                # Calls itself over local network context to reset the sleep timer
                urllib.request.urlopen(f"http://127.0.0.1:{port}", timeout=5)
                print("Self-ping successful: Staying awake!")
            except Exception as e:
                print(f"Self-ping warning: {e}")
            time.sleep(600) # Ping every 10 minutes

    threading.Thread(target=keep_alive, daemon=True).start()
    server.serve_forever()

# --- MAIN EXECUTION ---
if __name__ == '__main__':
    # 1. Start web server with auto-ping keeping it active
    threading.Thread(target=run_dummy_server, daemon=True).start()
    
    # 2. Start the Telegram Bot
    if not TOKEN:
        print("CRITICAL: Cannot start bot without a valid TELEGRAM_TOKEN environment variable.")
        exit(1)
        
    print("Bot is starting...")
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("morning", morning_checklist))
    app.add_handler(CommandHandler("evening", evening_review))
    app.add_handler(CommandHandler("back", back_to_previous_day))
    app.add_handler(CommandHandler("calendar", send_calendar))
    app.add_handler(CallbackQueryHandler(handle_task_click))

    print("Bot is running! Go to Telegram to use it.")
    app.run_polling()