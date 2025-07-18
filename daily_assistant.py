#!/usr/bin/env python3
"""
Daily Assistant - Outlook calendar summary
Gets today's calendar events and provides a clean summary
Run: python daily_assistant.py
"""

import subprocess
import sys
from datetime import datetime, timedelta
import os
import re
import random

def install_package(package_name, import_name=None):
    """Install a package using pip if it's not already installed"""
    if import_name is None:
        import_name = package_name
    
    try:
        __import__(import_name)
        print(f"✅ {package_name} is already installed")
        return True
    except ImportError:
        print(f"📦 {package_name} not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"✅ Successfully installed {package_name}")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install {package_name}: {e}")
            return False

def check_and_install_dependencies():
    """Check and install all required packages"""
    print("🔍 Checking required Python packages...")
    
    required_packages = [
        ("pywin32", "win32com.client"),  # For Outlook integration
        ("pyttsx3", "pyttsx3")           # For text-to-speech
    ]
    
    all_installed = True
    
    for package_name, import_name in required_packages:
        if not install_package(package_name, import_name):
            all_installed = False
    
    if not all_installed:
        print("❌ Some packages failed to install. The script may not work properly.")
        return False
    
    print("✅ All required packages are available!")
    return True

# Check and install dependencies before importing them
if not check_and_install_dependencies():
    print("⚠️ Warning: Some dependencies are missing. Continuing anyway...")

# Now import the packages (they should be available)
try:
    import win32com.client
except ImportError:
    print("❌ Critical: win32com.client not available. Outlook integration will fail.")
    win32com = None

try:
    import pyttsx3
except ImportError:
    print("❌ Critical: pyttsx3 not available. Text-to-speech will be disabled.")
    pyttsx3 = None

# Configuration
INCLUDE_TOMORROW = True  # Set to False to only show today's events

# Initialize TTS engine
def init_tts():
    """Initialize and configure TTS engine"""
    if pyttsx3 is None:
        print("⚠️ pyttsx3 not available - TTS disabled")
        return None
        
    try:
        engine = pyttsx3.init()
        
        # Set properties for better speech
        engine.setProperty('rate', 150)    # Speed of speech
        engine.setProperty('volume', 0.9)  # Volume level (0.0 to 1.0)
        
        # Try to use a specific voice (optional)
        voices = engine.getProperty('voices')
        if voices and len(voices) > 1:
            engine.setProperty('voice', voices[1].id)  # Use second voice if available
            
        return engine
    except Exception as e:
        print(f"⚠️ TTS initialization failed: {e}")
        return None

def speak_text(text, engine=None):
    """Convert text to speech"""
    if engine:
        try:
            engine.say(text)
            engine.runAndWait()
        except Exception as e:
            print(f"⚠️ TTS failed: {e}")
    else:
        print("🔇 TTS not available")

def clean_text_for_tts(text):
    """Remove special characters and clean up text for TTS"""
    if not text:
        return text
    
    # Replace common special characters with spaces or remove them
    text = re.sub(r'[-_/\\|]', ' ', text)  # Replace dashes, underscores, slashes with spaces
    text = re.sub(r'[^\w\s]', '', text)    # Remove all other special characters except letters, numbers, spaces
    text = re.sub(r'\s+', ' ', text)       # Replace multiple spaces with single space
    text = text.strip()                    # Remove leading/trailing spaces
    
    return text

def fix_24h_times_for_tts(text):
    """Fix 24-hour time formats that sound stupid in TTS"""
    if not text:
        return text
    
    # Fix specific patterns like "1200 to 1800" -> "12 to 6"
    # Only target 4-digit military times that are clearly time formats
    def convert_24h_time(match):
        time_str = match.group(0)
        hour = int(time_str[:2])
        minute = int(time_str[2:])
        
        # Convert to simple hour format for TTS
        if hour == 0:
            return "midnight" if minute == 0 else f"12 {minute:02d} AM"
        elif hour == 12:
            return "noon" if minute == 0 else f"12 {minute:02d} PM"
        elif hour > 12:
            return str(hour - 12) if minute == 0 else f"{hour - 12} {minute:02d}"
        else:
            return str(hour) if minute == 0 else f"{hour} {minute:02d}"
    
    # Only replace 4-digit times that look like military time (0000-2359)
    # More specific pattern to avoid false matches
    text = re.sub(r'\b([0-2]\d[0-5]\d)\b', convert_24h_time, text)
    
    return text

def get_random_transition():
    """Get a random transition word for natural speech"""
    transitions = ["followed by", "then", "next"]
    return random.choice(transitions)

def get_time_greeting():
    """Get appropriate greeting based on current time"""
    current_hour = datetime.now().hour
    
    if 5 <= current_hour < 12:
        return "Good morning"
    elif 12 <= current_hour < 17:
        return "Good afternoon"
    elif 17 <= current_hour < 21:
        return "Good evening"
    else:
        return "Good night"

def get_user_name():
    """Get user's name from file or ask for it"""
    name_file = os.path.join(os.path.dirname(__file__), 'user_name.txt')
    
    try:
        if os.path.exists(name_file):
            with open(name_file, 'r', encoding='utf-8') as f:
                name = f.read().strip()
                if name:
                    return name
    except Exception as e:
        print(f"⚠️ Error reading name file: {e}")
    
    # Ask for name if not found
    print("👋 Hi there! I don't know your name yet.")
    name = input("What's your name? ").strip()
    
    if name:
        try:
            with open(name_file, 'w', encoding='utf-8') as f:
                f.write(name)
            print(f"✅ Nice to meet you, {name}! I'll remember that.")
        except Exception as e:
            print(f"⚠️ Couldn't save your name: {e}")
    
    return name

def get_today_calendar_events():
    """Get today's Outlook calendar events and return as structured data - USING WORKING LOGIC FROM outlook_today.py"""
    
    print("📅 Reading LOCAL Outlook calendar...")
    
    if win32com is None:
        print("❌ win32com.client not available - cannot access Outlook")
        return []
    
    try:
        # Connect to local Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get default calendar folder
        calendar_folder = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        
        print(f"📂 Looking in calendar: {calendar_folder.Name}")
        print(f"📊 Total items in calendar: {calendar_folder.Items.Count}")
        
        # Get today's date range
        today = datetime.now().date()
        start_time = datetime.combine(today, datetime.min.time())
        end_time = datetime.combine(today, datetime.max.time())
        
        # Get appointments for today
        appointments = calendar_folder.Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = True
        
        # Use date restriction instead of looping through everything
        today_str = today.strftime('%m/%d/%Y')
        tomorrow_str = (today + timedelta(days=1)).strftime('%m/%d/%Y')
        restriction = f"[Start] >= '{today_str}' AND [Start] < '{tomorrow_str}'"
        
        print(f"🔍 Searching for events on {today_str} with filter: {restriction}")
        
        events = []
        
        try:
            filtered_appointments = appointments.Restrict(restriction)
            print(f"✅ Found {filtered_appointments.Count} events with filter")
            
            events = []
            for appointment in filtered_appointments:
                # Skip birthday and holiday events
                subject = appointment.Subject.lower()
                if ('birthday' in subject or 
                    'holiday' in subject or 
                    'anniversary' in subject or
                    appointment.Start.date() != today):  # Only today's events
                    continue
                    
                events.append(appointment)
                print(f"   📅 {appointment.Subject} at {appointment.Start}")
        except Exception as filter_error:
            print(f"⚠️ Filter failed: {filter_error}")
            print("🔄 Trying manual search (limited to last 100 items)...")
            
            # Fallback: manual search but limit to recent items
            events = []
            count = 0
            for appointment in appointments:
                count += 1
                if count > 100:  # Limit search to prevent hanging
                    print("⏹️ Stopping search at 100 items")
                    break
                try:
                    appt_date = appointment.Start.date()
                    if appt_date == today:
                        # Skip birthday and holiday events
                        subject = appointment.Subject.lower()
                        if ('birthday' in subject or 
                            'holiday' in subject or 
                            'anniversary' in subject):
                            continue
                            
                        events.append(appointment)
                        print(f"   📅 {appointment.Subject} at {appointment.Start}")
                except:
                    continue
    
    except Exception as e:
        print(f"❌ Error accessing Outlook: {e}")
        print("Make sure Outlook is installed and you have calendar events!")
        return []
    
    # Now convert the raw appointment objects to structured data
    events_data = []
    for appointment in events:
        try:
            # Basic required fields first
            subject = str(appointment.Subject) if appointment.Subject else "No Subject"
            start_time = appointment.Start.strftime('%I:%M %p')
            end_time = appointment.End.strftime('%I:%M %p')
            
            event_info = {
                'subject': subject,
                'start_time': start_time,
                'end_time': end_time
            }
            
            # Optional fields with safe access
            try:
                event_info['location'] = str(appointment.Location) if appointment.Location else ''
            except:
                event_info['location'] = ''
            
            # Try multiple approaches to get organizer info (like we learned from debugging)
            organizer_name = ''
            try:
                organizer_name = str(appointment.Organizer) if appointment.Organizer else ''
            except:
                # Try GetOrganizer method if direct access fails
                try:
                    organizer_obj = appointment.GetOrganizer()
                    if organizer_obj:
                        organizer_name = getattr(organizer_obj, 'Name', str(organizer_obj))
                        # Clean up Exchange Online email addresses
                        if '/o=ExchangeLabs/' in organizer_name:
                            organizer_name = organizer_name.split('cn=Recipients/cn=')[0] if 'cn=Recipients/cn=' in organizer_name else organizer_name
                except:
                    organizer_name = ''
            
            event_info['organizer'] = organizer_name
            
            try:
                event_info['is_recurring'] = bool(appointment.IsRecurring)
            except:
                event_info['is_recurring'] = False
            
            try:
                event_info['all_day'] = bool(appointment.AllDayEvent)
            except:
                event_info['all_day'] = False
            
            try:
                event_info['categories'] = str(appointment.Categories) if appointment.Categories else ''
            except:
                event_info['categories'] = ''
            
            try:
                event_info['importance'] = int(appointment.Importance)
            except:
                event_info['importance'] = 1
            
            try:
                event_info['reminder_set'] = bool(appointment.ReminderSet)
            except:
                event_info['reminder_set'] = False
            
            # Calculate duration
            try:
                duration = appointment.End - appointment.Start
                hours = duration.seconds // 3600
                minutes = (duration.seconds % 3600) // 60
                if hours > 0:
                    event_info['duration'] = f"{hours}h {minutes}m"
                else:
                    event_info['duration'] = f"{minutes}m"
            except:
                event_info['duration'] = "Unknown"
            
            # Check for online meeting
            try:
                location = str(event_info['location']).lower() if event_info['location'] else ""
                if "teams" in location:
                    event_info['meeting_type'] = "Microsoft Teams"
                    event_info['is_online'] = True
                elif "zoom" in location:
                    event_info['meeting_type'] = "Zoom"
                    event_info['is_online'] = True
                elif "webex" in location:
                    event_info['meeting_type'] = "WebEx"
                    event_info['is_online'] = True
                elif "http" in location:
                    event_info['meeting_type'] = "Online Meeting"
                    event_info['is_online'] = True
                else:
                    event_info['is_online'] = False
                    event_info['meeting_type'] = "In-person"
            except:
                event_info['is_online'] = False
                event_info['meeting_type'] = "Unknown"
            
            events_data.append(event_info)
            print(f"✅ Processed: {subject}")
            
        except Exception as e:
            print(f"⚠️ Skipping problematic event: {e}")
            continue
    
    return events_data

def get_tomorrow_calendar_events():
    """Get tomorrow's Outlook calendar events and return as structured data"""
    
    print("📅 Reading tomorrow's calendar events...")
    
    if win32com is None:
        print("❌ win32com.client not available - cannot access Outlook")
        return []
    
    try:
        # Connect to local Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get default calendar folder
        calendar_folder = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        
        # Get tomorrow's date range
        tomorrow = datetime.now().date() + timedelta(days=1)
        day_after_tomorrow = tomorrow + timedelta(days=1)
        
        # Get appointments for tomorrow
        appointments = calendar_folder.Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = True
        
        # Use date restriction for tomorrow
        tomorrow_str = tomorrow.strftime('%m/%d/%Y')
        day_after_str = day_after_tomorrow.strftime('%m/%d/%Y')
        restriction = f"[Start] >= '{tomorrow_str}' AND [Start] < '{day_after_str}'"
        
        print(f"🔍 Searching for events on {tomorrow_str} with filter: {restriction}")
        
        events = []
        
        try:
            filtered_appointments = appointments.Restrict(restriction)
            print(f"✅ Found {filtered_appointments.Count} events for tomorrow")
            
            events = []
            for appointment in filtered_appointments:
                # Skip birthday and holiday events
                subject = appointment.Subject.lower()
                if ('birthday' in subject or 
                    'holiday' in subject or 
                    'anniversary' in subject or
                    appointment.Start.date() != tomorrow):  # Only tomorrow's events
                    continue
                    
                events.append(appointment)
                print(f"   📅 {appointment.Subject} at {appointment.Start}")
        except Exception as filter_error:
            print(f"⚠️ Filter failed for tomorrow: {filter_error}")
            return []
    
    except Exception as e:
        print(f"❌ Error accessing Outlook for tomorrow: {e}")
        return []
    
    # Convert the raw appointment objects to structured data (same as today)
    events_data = []
    for appointment in events:
        try:
            # Basic required fields first
            subject = str(appointment.Subject) if appointment.Subject else "No Subject"
            start_time = appointment.Start.strftime('%I:%M %p')
            end_time = appointment.End.strftime('%I:%M %p')
            
            event_info = {
                'subject': subject,
                'start_time': start_time,
                'end_time': end_time
            }
            
            # Optional fields with safe access
            try:
                event_info['location'] = str(appointment.Location) if appointment.Location else ''
            except:
                event_info['location'] = ''
            
            # Try multiple approaches to get organizer info
            organizer_name = ''
            try:
                organizer_name = str(appointment.Organizer) if appointment.Organizer else ''
            except:
                try:
                    organizer_obj = appointment.GetOrganizer()
                    if organizer_obj:
                        organizer_name = getattr(organizer_obj, 'Name', str(organizer_obj))
                        if '/o=ExchangeLabs/' in organizer_name:
                            organizer_name = organizer_name.split('cn=Recipients/cn=')[0] if 'cn=Recipients/cn=' in organizer_name else organizer_name
                except:
                    organizer_name = ''
            
            event_info['organizer'] = organizer_name
            
            events_data.append(event_info)
            print(f"✅ Processed tomorrow: {subject}")
            
        except Exception as e:
            print(f"⚠️ Skipping problematic tomorrow event: {e}")
            continue
    
    return events_data

def get_ai_day_analysis(events_summary, events_data, tomorrow_events_data=None):
    """Generate a simple hard-coded summary of the day and optionally tomorrow"""
    
    print("📝 Creating your day summary...")
    
    try:
        if not events_data and not tomorrow_events_data:
            greeting = get_time_greeting()
            user_name = get_user_name()
            return f"{greeting} {user_name}! You have a free day with no scheduled meetings - perfect time to catch up on personal projects or take a well-deserved break!"
        
        # Get user's name and time greeting
        user_name = get_user_name()
        greeting = get_time_greeting()
        
        def format_time_for_speech(time_str):
            """Convert 09:00 AM to 9 for better TTS (no colons or symbols)"""
            # Extract hour and AM/PM
            match = re.match(r'(\d{1,2}):(\d{2})\s*(AM|PM)', time_str)
            if match:
                hour = int(match.group(1))
                minute = int(match.group(2))
                period = match.group(3)
                
                # Convert to 24h format first for easier handling
                if period == 'PM' and hour != 12:
                    hour += 12
                elif period == 'AM' and hour == 12:
                    hour = 0
                
                # Format for speech - just numbers, no symbols
                if minute == 0:
                    if hour == 0:
                        return "midnight"
                    elif hour == 12:
                        return "noon" 
                    elif hour > 12:
                        return f"{hour-12}"
                    else:
                        return f"{hour}"
                else:
                    # For non-zero minutes, just say the numbers
                    if hour > 12:
                        return f"{hour-12} {minute:02d}"
                    else:
                        return f"{hour} {minute:02d}"
            return time_str
        
        def create_meeting_summary(events, day_name="today"):
            """Create meeting summary for a given day"""
            meeting_parts = []
            for event in events:
                organizer = event['organizer'] if event['organizer'] else "Unknown organizer"
                
                # Clean the subject for TTS
                clean_subject = clean_text_for_tts(event['subject'])
                clean_subject = fix_24h_times_for_tts(clean_subject)
                clean_organizer = clean_text_for_tts(organizer)
                
                # Format times for natural speech
                start_time = format_time_for_speech(event['start_time'])
                end_time = format_time_for_speech(event['end_time'])
                time_range = f"{start_time} to {end_time}"
                
                # Check if the user is the organizer (meeting organized by them)
                if organizer and user_name and user_name.lower() in organizer.lower():
                    meeting_parts.append(f"{clean_subject} from {time_range}")
                else:
                    meeting_parts.append(f"{clean_subject} from {time_range} with {clean_organizer}")
            
            return meeting_parts
        
        # Create today's summary
        result = ""
        if events_data:
            meeting_parts = create_meeting_summary(events_data, "today")
            
            if len(meeting_parts) == 1:
                result = f"{greeting} {user_name}! You start your day with {meeting_parts[0]}."
            elif len(meeting_parts) == 2:
                transition = get_random_transition()
                result = f"{greeting} {user_name}! You start your day with {meeting_parts[0]}, {transition} {meeting_parts[1]}."
            else:
                # For 3+ meetings, use random transitions between each meeting
                result = f"{greeting} {user_name}! You start your day with {meeting_parts[0]}"
                
                # Add middle meetings with random transitions
                for i in range(1, len(meeting_parts) - 1):
                    transition = get_random_transition()
                    result += f", {transition} {meeting_parts[i]}"
                
                # Add the final meeting
                result += f", and finally {meeting_parts[-1]}."
        else:
            result = f"{greeting} {user_name}! You have a free day today."
        
        # Add tomorrow's summary if available
        if tomorrow_events_data:
            tomorrow_parts = create_meeting_summary(tomorrow_events_data, "tomorrow")
            
            if tomorrow_parts:
                if len(tomorrow_parts) == 1:
                    result += f" Tomorrow you have {tomorrow_parts[0]}."
                elif len(tomorrow_parts) == 2:
                    transition = get_random_transition()
                    result += f" Tomorrow you start with {tomorrow_parts[0]}, {transition} {tomorrow_parts[1]}."
                else:
                    result += f" Tomorrow you start with {tomorrow_parts[0]}"
                    for i in range(1, len(tomorrow_parts) - 1):
                        transition = get_random_transition()
                        result += f", {transition} {tomorrow_parts[i]}"
                    result += f", and finally {tomorrow_parts[-1]}."
            else:
                result += " Tomorrow is free with no scheduled meetings."
        
        return result
        
    except Exception as e:
        return f"Sorry, I couldn't create your day summary due to an error: {e}"

def display_daily_summary(events_data, ai_analysis, tomorrow_events_data=None):
    """Display the complete daily summary"""
    
    today = datetime.now().date()
    tomorrow = today + timedelta(days=1)
    
    print(f"\n{'='*70}")
    print(f"🗓️  YOUR DAILY BRIEFING - {today.strftime('%A, %B %d, %Y')}")
    print(f"{'='*70}")
    
    # Show today's events
    if events_data:
        print(f"\n📋 TODAY'S EVENTS ({len(events_data)} total):")
        print("-" * 50)
        
        for i, event in enumerate(events_data, 1):
            print(f"\n{i}. 📅 {event['subject']}")
            print(f"   🕐 {event['start_time']} - {event['end_time']}")
            
            if event['location']:
                if event.get('is_online'):
                    print(f"   💻 {event['location']}")
                else:
                    print(f"   📍 {event['location']}")
            
            if event.get('importance') == 2:
                print(f"   🔴 HIGH PRIORITY")
            
            if event['organizer'] and event['organizer'] != "":
                # Clean up organizer name for display
                organizer_display = event['organizer']
                # Remove Exchange Online path if present
                if '/o=ExchangeLabs/' in organizer_display:
                    # Extract just the name part
                    if 'cn=' in organizer_display:
                        try:
                            parts = organizer_display.split('cn=')
                            if len(parts) > 1:
                                organizer_display = parts[-1].split('-')[0].replace('_', ' ').title()
                        except:
                            pass
                print(f"   👤 Organizer: {organizer_display}")
    else:
        print(f"\n🎉 TODAY - Free day with no scheduled events!")
    
    # Show tomorrow's events if available
    if tomorrow_events_data is not None:
        if tomorrow_events_data:
            print(f"\n📋 TOMORROW'S EVENTS ({len(tomorrow_events_data)} total) - {tomorrow.strftime('%A, %B %d')}:")
            print("-" * 50)
            
            for i, event in enumerate(tomorrow_events_data, 1):
                print(f"\n{i}. 📅 {event['subject']}")
                print(f"   🕐 {event['start_time']} - {event['end_time']}")
                
                if event['location']:
                    print(f"   📍 {event['location']}")
                
                if event['organizer'] and event['organizer'] != "":
                    print(f"   👤 Organizer: {event['organizer']}")
        else:
            print(f"\n🎉 TOMORROW - Free day with no scheduled events!")
    
    # Show day summary
    print(f"\n{'='*70}")
    print(f"📝 DAY SUMMARY:")
    print(f"{'='*70}")
    print(ai_analysis)
    
    print(f"\n{'='*70}")
    print(f"✨ Have a great day! ✨")
    print(f"{'='*70}")

def main():
    """Main function - your daily assistant"""
    
    print("🌅 Good morning! Starting your daily briefing...")
    print("=" * 60)
    
    # Initialize TTS engine
    print("🔊 Initializing text-to-speech...")
    tts_engine = init_tts()
    
    # Get user's name
    user_name = get_user_name()
    
    # Get calendar events
    events_data = get_today_calendar_events()
    
    # Get tomorrow's events if enabled
    tomorrow_events_data = None
    if INCLUDE_TOMORROW:
        tomorrow_events_data = get_tomorrow_calendar_events()
    
    # Get day summary
    ai_analysis = get_ai_day_analysis("", events_data, tomorrow_events_data)
    
    # Display everything
    display_daily_summary(events_data, ai_analysis, tomorrow_events_data)
    
    # Speak the AI analysis
    if ai_analysis and tts_engine:
        print("\n🔊 Reading your daily summary...")
        speak_text(ai_analysis, tts_engine)
    elif ai_analysis:
        print("\n🔇 TTS not available, but here's your summary again:")
        print(ai_analysis)

if __name__ == "__main__":
    main()
