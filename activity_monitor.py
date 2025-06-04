import os
import sys
import time 
import datetime
import threading
import traceback
import tempfile
import logging
import signal
import time
import atexit
import re
import getpass
import socket
import requests
import json
import winreg
import glob
import subprocess
from logging.handlers import RotatingFileHandler
from typing import Dict, List, Optional, Tuple, Set, Any
from enum import Enum
from dataclasses import dataclass

import psutil
import pythoncom
import wmi
import win32gui
import win32process
import win32com.client
import win32api
import win32con
import win32evtlog
import win32evtlogutil
from enum import Enum
from typing import Tuple

# Audio detection imports
try:
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, AudioSession
    AUDIO_DETECTION_AVAILABLE = True
except ImportError:
    AUDIO_DETECTION_AVAILABLE = False

try:
    import win32evtlog
    import win32evtlogutil
    import win32api
    NATIVE_EVENTLOG_AVAILABLE = True
except ImportError:
    NATIVE_EVENTLOG_AVAILABLE = False

try:
    import pythoncom
    import wmi
    WMI_AVAILABLE = True
except ImportError:
    WMI_AVAILABLE = False

    

class EmailTimingMode(Enum):
    INTERVAL = "interval"
    TIME_OF_DAY = "time_of_day"

class ImprovedEmailTiming:
    """
    Improved email timing system that prioritizes sending BEFORE the target time
    to ensure people receive emails before they leave work
    """
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.logger = logging.getLogger(__name__)
        
        # Load timing configuration
        self.mode = self._get_timing_mode()
        self.interval_seconds = self._get_interval_seconds()
        self.daily_time = self._get_daily_time()
        self.last_daily_email_date = None
        
        self.logger.info(f"📧 Email timing mode: {self.mode.value}")
        if self.mode == EmailTimingMode.INTERVAL:
            self.logger.info(f"📧 Email interval: {self.interval_seconds}s ({self._format_duration(self.interval_seconds)})")
        elif self.mode == EmailTimingMode.TIME_OF_DAY:
            self.logger.info(f"📧 Target email time: {self.daily_time}")
            self.logger.info(f"📧 Early send window: 5 minutes before target (hard coded)")
            
            # Calculate actual send window
            try:
                hour, minute = map(int, self.daily_time.split(':'))
                target_time = datetime.time(hour, minute)
                early_time = (datetime.datetime.combine(datetime.date.today(), target_time) - 
                             datetime.timedelta(minutes=5)).time()
                self.logger.info(f"📧 Actual send window: {early_time.strftime('%H:%M')} - {self.daily_time}")
            except ValueError:
                pass
    
    def _get_timing_mode(self) -> EmailTimingMode:
        """Get the email timing mode from config"""
        mode_str = self.config_manager.get_config_value("email_timing_mode", "interval").lower()
        
        if mode_str == "time_of_day":
            return EmailTimingMode.TIME_OF_DAY
        else:
            return EmailTimingMode.INTERVAL
    
    def _get_interval_seconds(self) -> int:
        """Get email interval in seconds"""
        try:
            interval = int(self.config_manager.get_config_value("email_interval", "180"))
            return max(interval, 60)  # Minimum 60 seconds
        except ValueError:
            self.logger.warning("Invalid email_interval in config, using default 180s")
            return 180
    
    def _get_daily_time(self) -> str:
        """Get daily email time (HH:MM format)"""
        daily_time = self.config_manager.get_config_value("daily_email_time", "18:00")
        
        if self._validate_time_format(daily_time):
            return daily_time
        else:
            self.logger.warning(f"Invalid daily_email_time format '{daily_time}', using default 18:00")
            return "18:00"
    
    def _validate_time_format(self, time_str: str) -> bool:
        """Validate HH:MM time format"""
        try:
            time_parts = time_str.split(':')
            if len(time_parts) != 2:
                return False
            
            hour = int(time_parts[0])
            minute = int(time_parts[1])
            
            return 0 <= hour <= 23 and 0 <= minute <= 59
        except (ValueError, IndexError):
            return False
    
    def should_send_email_now(self, last_email_time: float) -> Tuple[bool, str]:
        """
        Check if it's time to send an email
        Returns: (should_send, reason)
        """
        if self.mode == EmailTimingMode.INTERVAL:
            return self._check_interval_timing(last_email_time)
        elif self.mode == EmailTimingMode.TIME_OF_DAY:
            return self._check_daily_timing_improved()
        else:
            return False, "Unknown timing mode"
    
    def _check_interval_timing(self, last_email_time: float) -> Tuple[bool, str]:
        """Check if enough time has passed since last email"""
        current_time = time.time()
        time_since_last = current_time - last_email_time
        
        if time_since_last >= self.interval_seconds:
            return True, f"Interval elapsed ({self._format_duration(int(time_since_last))} since last email)"
        else:
            remaining = self.interval_seconds - time_since_last
            return False, f"Waiting {self._format_duration(int(remaining))} for next interval"
    
    def _check_daily_timing_improved(self) -> Tuple[bool, str]:
        """
        IMPROVED: Check if it's time for daily email with early send capability
        This ensures emails are sent BEFORE people leave work
        """
        now = datetime.datetime.now()
        today_str = now.strftime('%Y-%m-%d')
        
        # Check if we already sent today's email
        if self.last_daily_email_date == today_str:
            return False, f"Daily email already sent today (target: {self.daily_time})"
        
        # Parse target time
        try:
            hour, minute = map(int, self.daily_time.split(':'))
            target_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        except ValueError:
            return False, f"Invalid daily_email_time format: {self.daily_time}"
        
        # Calculate early send time (5 minutes before target)
        early_send_time = target_time - datetime.timedelta(minutes=5)
        
        # If it's past the target time but we haven't sent yet
        if now >= target_time:
            # Send immediately if within a reasonable late window (5 minutes)
            minutes_late = (now - target_time).total_seconds() / 60
            if minutes_late <= 5:
                return True, f"URGENT: {minutes_late:.0f}m late for {self.daily_time} target - sending now!"
            else:
                # Too late for today - people probably left
                return False, f"Too late ({minutes_late:.0f}m past {self.daily_time}) - next email tomorrow"
        
        # If we're in the early send window
        elif now >= early_send_time:
            return True, f"Early send window active (target: {self.daily_time}, sending now to ensure delivery)"
        
        # Not time yet
        else:
            time_until_early = (early_send_time - now).total_seconds()
            return False, f"Email in {self._format_duration(int(time_until_early))} (early window starts {early_send_time.strftime('%H:%M')})"
    
    def mark_daily_email_sent(self):
        """Mark that today's daily email has been sent"""
        if self.mode == EmailTimingMode.TIME_OF_DAY:
            self.last_daily_email_date = datetime.datetime.now().strftime('%Y-%m-%d')
            self.logger.info(f"📧 Marked daily email as sent for {self.last_daily_email_date}")
    
    def get_timing_status(self) -> dict:
        """Get current timing configuration status"""
        status = {
            'mode': self.mode.value,
            'interval_seconds': self.interval_seconds if self.mode == EmailTimingMode.INTERVAL else None,
            'interval_formatted': self._format_duration(self.interval_seconds) if self.mode == EmailTimingMode.INTERVAL else None,
            'daily_time': self.daily_time if self.mode == EmailTimingMode.TIME_OF_DAY else None,
            'early_send_minutes': 5 if self.mode == EmailTimingMode.TIME_OF_DAY else None,  # Hard coded to 5
            'last_daily_email_date': self.last_daily_email_date,
            'next_email_info': self._get_next_email_info()
        }
        
        # Add early send window info for time_of_day mode
        if self.mode == EmailTimingMode.TIME_OF_DAY:
            try:
                hour, minute = map(int, self.daily_time.split(':'))
                target_time = datetime.time(hour, minute)
                early_time = (datetime.datetime.combine(datetime.date.today(), target_time) - 
                             datetime.timedelta(minutes=5)).time()
                status['early_send_time'] = early_time.strftime('%H:%M')
                status['send_window'] = f"{status['early_send_time']} - {self.daily_time}"
            except ValueError:
                status['early_send_time'] = "Invalid"
                status['send_window'] = "Invalid"
        
        return status
    
    def _get_next_email_info(self) -> str:
        """Get human-readable info about when next email will be sent"""
        if self.mode == EmailTimingMode.INTERVAL:
            return f"Next email in up to {self._format_duration(self.interval_seconds)}"
        elif self.mode == EmailTimingMode.TIME_OF_DAY:
            now = datetime.datetime.now()
            today_str = now.strftime('%Y-%m-%d')
            
            if self.last_daily_email_date == today_str:
                return "Daily email already sent today"
            
            try:
                hour, minute = map(int, self.daily_time.split(':'))
                target_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                early_send_time = target_time - datetime.timedelta(minutes=5)
                
                if now < early_send_time:
                    # Before early send window
                    time_until = (early_send_time - now).total_seconds()
                    return f"Next email today at {early_send_time.strftime('%H:%M')} (in {self._format_duration(int(time_until))}) [early window for {self.daily_time} target]"
                elif now < target_time:
                    # In early send window
                    return f"IN SEND WINDOW NOW (target: {self.daily_time})"
                elif (now - target_time).total_seconds() <= 300:  # Within 5 minutes of target
                    return f"LATE for {self.daily_time} target - should send immediately"
                else:
                    # Too late for today
                    tomorrow = now + datetime.timedelta(days=1)
                    tomorrow_early = tomorrow.replace(hour=early_send_time.hour, minute=early_send_time.minute, second=0)
                    time_until = (tomorrow_early - now).total_seconds()
                    return f"Missed today - next email tomorrow at {early_send_time.strftime('%H:%M')} (in {self._format_duration(int(time_until))})"
            except ValueError:
                return f"Next email at {self.daily_time} (invalid time format)"
        
        return "Unknown"
    
    def _format_duration(self, seconds: int) -> str:
        """Format duration as human-readable string"""
        if seconds < 60:
            return f"{seconds}s"
        elif seconds < 3600:
            minutes = seconds // 60
            return f"{minutes}m"
        else:
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            if minutes > 0:
                return f"{hours}h {minutes}m"
            else:
                return f"{hours}h"

@dataclass
class BackgroundVideoSession:
    browser: str
    site: str
    title: str
    start_time: str
    duration: int  # seconds
    impact_level: str  # 'high', 'medium', 'low'
    session_id: str  # Unique identifier for the session

@dataclass
class LoginSession:
    login_time: datetime.datetime
    logout_time: Optional[datetime.datetime] = None
    duration_seconds: Optional[int] = None
    session_type: str = "Normal"  # "Normal", "Startup", "Incomplete"
    
    def calculate_duration(self) -> Optional[int]:
        """Calculate session duration in seconds"""
        if self.login_time and self.logout_time:
            self.duration_seconds = int((self.logout_time - self.login_time).total_seconds())
            return self.duration_seconds
        return None
    
    def is_complete(self) -> bool:
        """Check if session has both login and logout"""
        return self.login_time is not None and self.logout_time is not None
    
    def format_duration(self) -> str:
        """Format duration as human-readable string"""
        if not self.duration_seconds:
            return "Ongoing"
        
        hours = self.duration_seconds // 3600
        minutes = (self.duration_seconds % 3600) // 60
        seconds = self.duration_seconds % 60
        
        if hours > 0:
            return f"{hours}h {minutes}m"
        elif minutes > 0:
            return f"{minutes}m {seconds}s"
        else:
            return f"{seconds}s"

@dataclass
class ProductivityData:
    productive_time: int  # seconds
    unproductive_time: int  # seconds (foreground only)
    background_video_time: int  # seconds (total background video time)
    verified_playing_time: int  # seconds (verified playing time)
    productive_apps: Dict[str, int]  # app_name: seconds
    unproductive_apps: Dict[str, int]  # app_name: seconds
    uncategorized_apps: Dict[str, int]  # ADDED: uncategorized websites
    background_videos: List[BackgroundVideoSession]
    background_video_apps: Dict[str, int]  # background video time by app/site
    verified_playing_apps: Dict[str, int]  # verified playing time by app/site
    date: str
    system_info: Dict[str, Any] = None

class Category(Enum):
    PRODUCTIVE = "Productive"
    UNPRODUCTIVE = "Unproductive"
    UNCATEGORIZED = "Uncategorized"  # ADDED: New category for uncategorized websites

class AppNameCleaner:
    """Helper class to clean and format app names for better display"""
    
    # Mapping for specific app name improvements
    APP_NAME_MAPPING = {
        'olk.exe': 'Outlook (Microsoft Store)',
        'olkbg.exe': 'Outlook (Microsoft Store)',
        'hxoutlook.exe': 'Outlook (Microsoft Store)',
        'outlookforwindows.exe': 'Outlook (Microsoft Store)',
        'winword.exe': 'Microsoft Word',
        'powerpnt.exe': 'Microsoft PowerPoint',
        'excel.exe': 'Microsoft Excel',
        'outlook.exe': 'Microsoft Outlook',
        'msteams.exe': 'Microsoft Teams',
        'onenote.exe': 'Microsoft OneNote',
        'msaccess.exe': 'Microsoft Access',
        'mspub.exe': 'Microsoft Publisher',
        'visio.exe': 'Microsoft Visio',
        'winproj.exe': 'Microsoft Project',
        'chrome.exe': 'Google Chrome',
        'firefox.exe': 'Mozilla Firefox',
        'msedge.exe': 'Microsoft Edge',
        'iexplore.exe': 'Internet Explorer',
        'opera.exe': 'Opera',
        'brave.exe': 'Brave Browser',
        'vivaldi.exe': 'Vivaldi',
        'safari.exe': 'Safari',
        'notepad.exe': 'Notepad',
        'notepad++.exe': 'Notepad++',
        'code.exe': 'Visual Studio Code',
        'devenv.exe': 'Visual Studio',
        'sublime_text.exe': 'Sublime Text',
        'atom.exe': 'Atom',
        'explorer.exe': 'File Explorer',
        'taskmgr.exe': 'Task Manager',
        'calc.exe': 'Calculator',
        'mspaint.exe': 'Paint',
        'cmd.exe': 'Command Prompt',
        'powershell.exe': 'PowerShell',
        'steam.exe': 'Steam',
        'discord.exe': 'Discord',
        'slack.exe': 'Slack',
        'zoom.exe': 'Zoom',
        'skype.exe': 'Skype',
        'spotify.exe': 'Spotify',
        'vlc.exe': 'VLC Media Player'
    }
    
    @classmethod
    def clean_app_name(cls, app_name: str) -> str:
        """
        Clean app name by removing .exe extension and applying custom mappings
        
        Args:
            app_name: Raw app name (e.g., "winword.exe" or "Microsoft Word.exe")
            
        Returns:
            Cleaned app name (e.g., "Microsoft Word")
        """
        if not app_name:
            return app_name
            
        # Convert to lowercase for comparison
        app_lower = app_name.lower().strip()
        
        # Check if we have a specific mapping for this app
        if app_lower in cls.APP_NAME_MAPPING:
            return cls.APP_NAME_MAPPING[app_lower]
        
        # If no specific mapping, just remove .exe extension
        if app_lower.endswith('.exe'):
            # Remove .exe and capitalize properly
            clean_name = app_name[:-4]  # Remove last 4 characters (.exe)
            
            # Handle special cases for better formatting
            if clean_name.lower() == 'chrome':
                return 'Google Chrome'
            elif clean_name.lower() == 'firefox':
                return 'Mozilla Firefox'
            elif clean_name.lower() == 'msedge':
                return 'Microsoft Edge'
            elif clean_name.lower().startswith('ms') and len(clean_name) > 2:
                # Handle MS prefixed apps (msword, mspaint, etc.)
                return f"Microsoft {clean_name[2:].title()}"
            else:
                # Default: just capitalize the name
                return clean_name.title()
        
        # If it doesn't end with .exe, return as-is
        return app_name
    
    @classmethod
    def clean_app_base_name(cls, full_app_title: str) -> str:
        """
        Extract and clean the base app name from a full app title
        
        Args:
            full_app_title: Full title like "winword.exe - Document1.docx"
            
        Returns:
            Cleaned base app name like "Microsoft Word"
        """
        if not full_app_title:
            return full_app_title
            
        # Split on " - " to get the base app name
        if " - " in full_app_title:
            base_name = full_app_title.split(" - ")[0].strip()
        else:
            base_name = full_app_title.strip()
        
        return cls.clean_app_name(base_name)
    
    @classmethod
    def clean_all_exe_from_text(cls, text: str) -> str:
        """Remove all .exe extensions from any text"""
        import re
        return re.sub(r'(\w+)\.exe', r'\1', text, flags=re.IGNORECASE)

# REPLACE the SystemInfoCollector class with this improved version:

class SystemInfoCollector:
    """Collects system and location information for productivity reports"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._cached_info = None
        self._cache_timestamp = 0
        self._cache_duration = 3600  # Cache for 1 hour
    
    def get_system_info(self) -> Dict[str, Any]:
        """Get comprehensive system information including username, IP, and location"""
        current_time = time.time()
        
        # Use cache if available and not expired
        if (self._cached_info and 
            current_time - self._cache_timestamp < self._cache_duration):
            return self._cached_info
        
        system_info = {
            'username': self._get_username(),
            'computer_name': self._get_computer_name(),
            'local_ip': self._get_local_ip(),
            'external_ip': self._get_external_ip(),
            'location': self._get_geolocation(),
            'timestamp': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Cache the results
        self._cached_info = system_info
        self._cache_timestamp = current_time
        
        return system_info
    
    def _get_username(self) -> str:
        """Get the current Windows username"""
        try:
            return getpass.getuser()
        except Exception as e:
            self.logger.debug(f"Error getting username: {e}")
            return "Unknown User"
    
    def _get_computer_name(self) -> str:
        """Get the computer name"""
        try:
            return socket.gethostname()
        except Exception as e:
            self.logger.debug(f"Error getting computer name: {e}")
            return "Unknown Computer"
    
    def _get_local_ip(self) -> str:
        """Get the local IP address"""
        try:
            # Connect to a remote address to get local IP
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            local_ip = s.getsockname()[0]
            s.close()
            return local_ip
        except Exception as e:
            self.logger.debug(f"Error getting local IP: {e}")
            return "Unknown"
    
    def _get_external_ip(self) -> str:
        """Get the external/public IP address"""
        try:
            # Try multiple services in case one is down
            services = [
                "https://api.ipify.org?format=text",
                "https://checkip.amazonaws.com",
                "https://icanhazip.com",
                "https://ipecho.net/plain",
                "https://myexternalip.com/raw"
            ]
            
            for service in services:
                try:
                    response = requests.get(service, timeout=5)
                    if response.status_code == 200:
                        ip = response.text.strip()
                        # Validate IP format
                        if self._is_valid_ip(ip):
                            return ip
                except Exception as e:
                    self.logger.debug(f"IP service {service} failed: {e}")
                    continue
            
            return "Unable to determine"
            
        except Exception as e:
            self.logger.debug(f"Error getting external IP: {e}")
            return "Unable to determine"
    
    def _is_valid_ip(self, ip: str) -> bool:
        """Validate IP address format"""
        try:
            parts = ip.split('.')
            return len(parts) == 4 and all(0 <= int(part) <= 255 for part in parts)
        except:
            return False
    
    def _get_geolocation(self) -> Dict[str, str]:
        """Get geolocation information using multiple services"""
        
        # List of geolocation services to try
        geolocation_services = [
            self._try_ipapi_co,
            self._try_ipinfo_io,
            self._try_ip_api,
            self._try_geolocation_db,
            self._try_ipstack
        ]
        
        for service_func in geolocation_services:
            try:
                result = service_func()
                if result and result.get('country') != 'Unknown':
                    self.logger.info(f"Geolocation successful via {service_func.__name__}")
                    return result
            except Exception as e:
                self.logger.debug(f"Geolocation service {service_func.__name__} failed: {e}")
                continue
        
        self.logger.warning("All geolocation services failed")
        return self._get_fallback_location()
    
    def _try_ipapi_co(self) -> Dict[str, str]:
        """Try ipapi.co geolocation service"""
        response = requests.get("https://ipapi.co/json/", timeout=10, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        if response.status_code == 200:
            data = response.json()
            
            # Check if we got an error response
            if 'error' in data:
                raise Exception(f"API error: {data.get('reason', 'Unknown error')}")
            
            return {
                'city': data.get('city', 'Unknown'),
                'region': data.get('region', 'Unknown'),
                'country': data.get('country_name', 'Unknown'),
                'country_code': data.get('country_code', 'Unknown'),
                'timezone': data.get('timezone', 'Unknown'),
                'latitude': str(data.get('latitude', 'Unknown')),
                'longitude': str(data.get('longitude', 'Unknown')),
                'isp': data.get('org', 'Unknown')
            }
        else:
            raise Exception(f"HTTP {response.status_code}")
    
    def _try_ipinfo_io(self) -> Dict[str, str]:
        """Try ipinfo.io geolocation service"""
        response = requests.get("https://ipinfo.io/json", timeout=10, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        if response.status_code == 200:
            data = response.json()
            
            location_parts = data.get('loc', ',').split(',')
            latitude = location_parts[0] if len(location_parts) > 0 else 'Unknown'
            longitude = location_parts[1] if len(location_parts) > 1 else 'Unknown'
            
            return {
                'city': data.get('city', 'Unknown'),
                'region': data.get('region', 'Unknown'),
                'country': data.get('country', 'Unknown'),
                'country_code': data.get('country', 'Unknown'),
                'timezone': data.get('timezone', 'Unknown'),
                'latitude': latitude,
                'longitude': longitude,
                'isp': data.get('org', 'Unknown')
            }
        else:
            raise Exception(f"HTTP {response.status_code}")
    
    def _try_ip_api(self) -> Dict[str, str]:
        """Try ip-api.com geolocation service"""
        response = requests.get("http://ip-api.com/json/?fields=status,country,countryCode,region,city,timezone,lat,lon,isp", 
                              timeout=10, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        if response.status_code == 200:
            data = response.json()
            
            if data.get('status') != 'success':
                raise Exception(f"API status: {data.get('status')}")
            
            return {
                'city': data.get('city', 'Unknown'),
                'region': data.get('region', 'Unknown'),
                'country': data.get('country', 'Unknown'),
                'country_code': data.get('countryCode', 'Unknown'),
                'timezone': data.get('timezone', 'Unknown'),
                'latitude': str(data.get('lat', 'Unknown')),
                'longitude': str(data.get('lon', 'Unknown')),
                'isp': data.get('isp', 'Unknown')
            }
        else:
            raise Exception(f"HTTP {response.status_code}")
    
    def _try_geolocation_db(self) -> Dict[str, str]:
        """Try geolocation-db.com service"""
        response = requests.get("https://geolocation-db.com/json/", timeout=10, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        if response.status_code == 200:
            data = response.json()
            
            return {
                'city': data.get('city', 'Unknown'),
                'region': data.get('state', 'Unknown'),
                'country': data.get('country_name', 'Unknown'),
                'country_code': data.get('country_code', 'Unknown'),
                'timezone': 'Unknown',  # This service doesn't provide timezone
                'latitude': str(data.get('latitude', 'Unknown')),
                'longitude': str(data.get('longitude', 'Unknown')),
                'isp': 'Unknown'  # This service doesn't provide ISP
            }
        else:
            raise Exception(f"HTTP {response.status_code}")
    
    def _try_ipstack(self) -> Dict[str, str]:
        """Try ipstack.com service (free tier, no API key needed for basic info)"""
        response = requests.get("http://api.ipstack.com/check?access_key=free", timeout=10, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        if response.status_code == 200:
            data = response.json()
            
            # Check if we got a valid response
            if 'country_name' in data:
                return {
                    'city': data.get('city', 'Unknown'),
                    'region': data.get('region_name', 'Unknown'),
                    'country': data.get('country_name', 'Unknown'),
                    'country_code': data.get('country_code', 'Unknown'),
                    'timezone': data.get('time_zone', {}).get('id', 'Unknown') if data.get('time_zone') else 'Unknown',
                    'latitude': str(data.get('latitude', 'Unknown')),
                    'longitude': str(data.get('longitude', 'Unknown')),
                    'isp': 'Unknown'
                }
            else:
                raise Exception("Invalid response format")
        else:
            raise Exception(f"HTTP {response.status_code}")
    
    def _get_fallback_location(self) -> Dict[str, str]:
        """Fallback location information if all geolocation services fail"""
        return {
            'city': 'Unknown',
            'region': 'Unknown', 
            'country': 'Unknown',
            'country_code': 'Unknown',
            'timezone': 'Unknown',
            'latitude': 'Unknown',
            'longitude': 'Unknown',
            'isp': 'Unknown'
        }
    
    def format_location_string(self, location: Dict[str, str]) -> str:
        """Format location information as a readable string"""
        if location['city'] != 'Unknown' and location['country'] != 'Unknown':
            if location['region'] != 'Unknown':
                return f"{location['city']}, {location['region']}, {location['country']}"
            else:
                return f"{location['city']}, {location['country']}"
        elif location['country'] != 'Unknown':
            return location['country']
        else:
            return "Location Unknown"
    
    def test_geolocation_services(self) -> Dict[str, str]:
        """Test all geolocation services and return results"""
        results = {}
        
        services = [
            ('ipapi.co', self._try_ipapi_co),
            ('ipinfo.io', self._try_ipinfo_io),
            ('ip-api.com', self._try_ip_api),
            ('geolocation-db.com', self._try_geolocation_db),
            ('ipstack.com', self._try_ipstack)
        ]
        
        for service_name, service_func in services:
            try:
                result = service_func()
                results[service_name] = f"✅ Success: {result.get('country', 'Unknown')}"
            except Exception as e:
                results[service_name] = f"❌ Failed: {str(e)}"
        
        return results


# ADD this method to test geolocation services - you can call this to debug:

def test_geolocation():
    """Test function to debug geolocation services"""
    print("=== TESTING GEOLOCATION SERVICES ===")
    
    collector = SystemInfoCollector()
    results = collector.test_geolocation_services()
    
    for service, result in results.items():
        print(f"{service:<20} {result}")
    
    print("\n=== FINAL RESULT ===")
    location = collector._get_geolocation()
    print(f"Location: {collector.format_location_string(location)}")
    print(f"Timezone: {location.get('timezone', 'Unknown')}")
    print(f"ISP: {location.get('isp', 'Unknown')}")

@dataclass
class CompleteEnhancedConfig:
    EMAIL_INTERVAL: int = 180
    SLEEP_INTERVAL: int = 1
    LOG_INTERVAL: int = 60
    UNPRODUCTIVE_WARNING_THRESHOLD: int = 600
    MAX_EMAIL_RETRIES: int = 3
    EMAIL_RETRY_DELAY: int = 1
    FRIDAY_ONLY: bool = False
    
    # Log Management Settings
    CLEANUP_DAYS_TO_KEEP: int = 7
    MAX_LOG_SIZE_MB: int = 10
    LOG_BACKUP_COUNT: int = 5
    MAX_ACTIVITY_LOG_SIZE_MB: int = 50
    
    # Dated Backup Settings
    DATED_BACKUP_DAYS_TO_KEEP: int = 30
    ENABLE_DATED_BACKUPS: bool = True
    
    # Missed Report Settings
    MISSED_REPORT_DAYS_BACK: int = 3
    ENABLE_MISSED_REPORT_RECOVERY: bool = True

    LOG_DIR: str = os.path.join(os.getcwd(), "logs")

    def __post_init__(self):
        os.makedirs(self.LOG_DIR, exist_ok=True)
        self.ACTIVITY_LOG = os.path.join(self.LOG_DIR, "monitor_output.log")
        self.DEBUG_LOG = os.path.join(self.LOG_DIR, "startup_debug.log")
        self.EMAIL_TRACK_FILE = os.path.join(self.LOG_DIR, "last_productivity_email_sent.txt")

        if getattr(sys, 'frozen', False):
            self.CONFIG_PATH = os.path.join(os.path.dirname(sys.executable), "config.txt")
        else:
            self.CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.txt")
        
        self.validate_and_setup()
    
    def validate_and_setup(self):
        """Validate configuration and create required directories"""
        if self.EMAIL_INTERVAL < 60:
            self.EMAIL_INTERVAL = 180
        if self.SLEEP_INTERVAL < 1:
            self.SLEEP_INTERVAL = 1
        if self.LOG_INTERVAL < 10:
            self.LOG_INTERVAL = 60
        os.makedirs(self.LOG_DIR, exist_ok=True)

class CompleteEnhancedProductivityDataPersistence:
    """Complete persistence with log management and dated backups for missed reports"""
    
    def __init__(self, config: CompleteEnhancedConfig):
        self.config = config
        self.logger = logging.getLogger(__name__)
        
        # Original persistence files
        self.app_times_file = os.path.join(config.LOG_DIR, "app_times_data.json")
        self.background_video_file = os.path.join(config.LOG_DIR, "background_video_data.json")
        self.session_data_file = os.path.join(config.LOG_DIR, "session_data.json")
        
        # NEW: Dated backup directory for missed report recovery
        self.dated_backup_dir = os.path.join(config.LOG_DIR, "daily_backups")
        if config.ENABLE_DATED_BACKUPS:
            os.makedirs(self.dated_backup_dir, exist_ok=True)
        
    def save_tracking_data(self, tracker: 'ForegroundTracker', bg_tracker: 'BackgroundVideoTracker'):
        """Save current tracking data AND create dated backup for missed reports"""
        
        print(f"🔍 SAVE CALLED: {datetime.datetime.now()}")
        print(f"   Tracker: {tracker is not None}")
        print(f"   BG Tracker: {bg_tracker is not None}")
    
        if tracker:
            with tracker.lock:
                app_count = len(tracker.app_times)
                print(f"   Apps to save: {app_count}")

        try:
            current_date = datetime.datetime.now().strftime('%Y-%m-%d')
            
            # ORIGINAL: Save current session data
            app_times = tracker.get_app_times() if tracker else {}
            
            app_data = {
                'app_times': {str(app): float(time_val) for app, time_val in app_times.items()},
                'last_updated': time.time(),
                'date': current_date,
                'total_apps': len(app_times),
                'total_time': sum(app_times.values()) if app_times else 0
            }
            self._save_json_file(self.app_times_file, app_data)
            
            # Background video data
            if bg_tracker:
                with bg_tracker.lock:
                    bg_video_times = dict(bg_tracker.background_video_times)
                    verified_times = dict(bg_tracker.verified_playing_times)
                
                background_data = {
                    'background_video_times': {str(site): float(time_val) for site, time_val in bg_video_times.items()},
                    'verified_playing_times': {str(site): float(time_val) for site, time_val in verified_times.items()},
                    'last_updated': time.time(),
                    'date': current_date,
                    'total_sites': len(bg_video_times),
                    'total_bg_time': sum(bg_video_times.values()) if bg_video_times else 0
                }
                self._save_json_file(self.background_video_file, background_data)
            
            # NEW: Also save dated backup for missed report recovery
            if self.config.ENABLE_DATED_BACKUPS:
                self._save_dated_backup(tracker, bg_tracker, current_date, app_times, bg_video_times if bg_tracker else {}, verified_times if bg_tracker else {})
            
            # Log save summary
            total_tracked_time = sum(app_times.values()) if app_times else 0
            self.logger.info(f"💾 Saved tracking data: {len(app_times)} apps, {total_tracked_time:.1f}s total time")
            
        except Exception as e:
            self.logger.error(f"Error saving tracking data: {e}")
            self.logger.error(f"Full traceback: {traceback.format_exc()}")

    def _save_dated_backup(self, tracker, bg_tracker, current_date: str, app_times: dict, bg_video_times: dict, verified_times: dict):
        """Save dated backup files for missed report recovery"""
        try:
            # Save dated app times backup
            dated_app_file = os.path.join(self.dated_backup_dir, f"app_times_{current_date}.json")
            app_backup_data = {
                'app_times': {str(app): float(time_val) for app, time_val in app_times.items()},
                'date': current_date,
                'timestamp': time.time(),
                'total_apps': len(app_times),
                'total_time': sum(app_times.values()) if app_times else 0
            }
            self._save_json_file(dated_app_file, app_backup_data)
            
            # Save dated background video backup
            if bg_video_times or verified_times:
                dated_bg_file = os.path.join(self.dated_backup_dir, f"background_video_{current_date}.json")
                bg_backup_data = {
                    'background_video_times': {str(site): float(time_val) for site, time_val in bg_video_times.items()},
                    'verified_playing_times': {str(site): float(time_val) for site, time_val in verified_times.items()},
                    'date': current_date,
                    'timestamp': time.time(),
                    'total_sites': len(bg_video_times),
                    'total_bg_time': sum(bg_video_times.values()) if bg_video_times else 0
                }
                self._save_json_file(dated_bg_file, bg_backup_data)
            
            self.logger.debug(f"💾 Saved dated backup for {current_date}")
            
        except Exception as e:
            self.logger.error(f"Error saving dated backup: {e}")
    
    def load_tracking_data(self, tracker: 'ForegroundTracker', bg_tracker: 'BackgroundVideoTracker'):
        """Load tracking data from disk and restore to trackers"""
        try:
            current_date = datetime.datetime.now().strftime('%Y-%m-%d')
            self.logger.info(f"🔄 Loading tracking data for date: {current_date}")
            
            # Load app times data
            app_data = self._load_json_file(self.app_times_file)
            if app_data:
                self.logger.info(f"📂 App data file found, date: {app_data.get('date', 'Unknown')}")
                
                if app_data.get('date') == current_date:
                    if tracker and 'app_times' in app_data:
                        loaded_app_times = {}
                        for app, time_val in app_data['app_times'].items():
                            try:
                                loaded_app_times[str(app)] = float(time_val)
                            except (ValueError, TypeError):
                                self.logger.warning(f"⚠️  Invalid time value for {app}: {time_val}")
                                continue
                        
                        with tracker.lock:
                            tracker.app_times.clear()
                            tracker.app_times.update(loaded_app_times)
                        
                        restored_apps = len(loaded_app_times)
                        total_time = sum(loaded_app_times.values())
                        self.logger.info(f"✅ Restored {restored_apps} app entries, {total_time:.1f}s total from previous session")
                        
                        if loaded_app_times:
                            top_apps = sorted(loaded_app_times.items(), key=lambda x: x[1], reverse=True)[:3]
                            self.logger.info(f"🔍 Top restored apps: {[(app[:30], f'{time:.1f}s') for app, time in top_apps]}")
                    else:
                        self.logger.warning("⚠️  No app_times data in saved file")
                else:
                    self.logger.info(f"📅 Saved data is from different day ({app_data.get('date')}), starting fresh")
            else:
                self.logger.info("📂 No previous app data file found - starting fresh")
            
            # Load background video data
            bg_data = self._load_json_file(self.background_video_file)
            if bg_data:
                self.logger.info(f"📂 Background video data file found, date: {bg_data.get('date', 'Unknown')}")
                
                if bg_data.get('date') == current_date and bg_tracker:
                    with bg_tracker.lock:
                        if 'background_video_times' in bg_data:
                            bg_tracker.background_video_times.clear()
                            for site, time_val in bg_data['background_video_times'].items():
                                try:
                                    bg_tracker.background_video_times[str(site)] = float(time_val)
                                except (ValueError, TypeError):
                                    self.logger.warning(f"⚠️  Invalid bg video time for {site}: {time_val}")
                        
                        if 'verified_playing_times' in bg_data:
                            bg_tracker.verified_playing_times.clear()
                            for site, time_val in bg_data['verified_playing_times'].items():
                                try:
                                    bg_tracker.verified_playing_times[str(site)] = float(time_val)
                                except (ValueError, TypeError):
                                    self.logger.warning(f"⚠️  Invalid verified time for {site}: {time_val}")
                    
                    restored_videos = len(bg_data.get('background_video_times', {}))
                    total_bg_time = sum(float(t) for t in bg_data.get('background_video_times', {}).values())
                    self.logger.info(f"✅ Restored {restored_videos} background video entries, {total_bg_time:.1f}s total")
                else:
                    self.logger.info(f"📅 Background video data is from different day, starting fresh")
            else:
                self.logger.info("📂 No previous background video data found - starting fresh")
            
        except Exception as e:
            self.logger.error(f"Error loading tracking data: {e}")
            self.logger.error(f"Full traceback: {traceback.format_exc()}")

    def load_historical_data(self, target_date: str) -> Optional[Dict[str, Any]]:
        """Load historical tracking data for a specific date (for missed reports)"""
        if not self.config.ENABLE_DATED_BACKUPS:
            return None
            
        try:
            dated_app_file = os.path.join(self.dated_backup_dir, f"app_times_{target_date}.json")
            dated_bg_file = os.path.join(self.dated_backup_dir, f"background_video_{target_date}.json")
            
            app_data = None
            bg_data = None
            
            if os.path.exists(dated_app_file):
                app_data = self._load_json_file(dated_app_file)
                self.logger.info(f"📂 Found app data for {target_date}: {app_data.get('total_apps', 0)} apps")
            
            if os.path.exists(dated_bg_file):
                bg_data = self._load_json_file(dated_bg_file)
                self.logger.info(f"📂 Found background video data for {target_date}: {bg_data.get('total_sites', 0)} sites")
            
            if app_data or bg_data:
                return {
                    'app_data': app_data,
                    'bg_data': bg_data,
                    'date': target_date
                }
            
            self.logger.info(f"📂 No historical data found for {target_date}")
            return None
            
        except Exception as e:
            self.logger.error(f"Error loading historical data for {target_date}: {e}")
            return None

    def cleanup_old_dated_backups(self):
        """Clean up dated backup files older than configured days"""
        if not self.config.ENABLE_DATED_BACKUPS:
            return
            
        try:
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=self.config.DATED_BACKUP_DAYS_TO_KEEP)
            cutoff_str = cutoff_date.strftime('%Y-%m-%d')
            
            backup_files = glob.glob(os.path.join(self.dated_backup_dir, "*.json"))
            cleaned_count = 0
            
            for backup_file in backup_files:
                try:
                    filename = os.path.basename(backup_file)
                    if '_' in filename:
                        date_part = filename.split('_')[-1].replace('.json', '')
                        if date_part < cutoff_str:
                            os.remove(backup_file)
                            cleaned_count += 1
                            self.logger.debug(f"🧹 Cleaned old backup: {filename}")
                except Exception as e:
                    self.logger.debug(f"Error processing backup file {backup_file}: {e}")
            
            if cleaned_count > 0:
                self.logger.info(f"🧹 Cleaned up {cleaned_count} old dated backup files")
                
        except Exception as e:
            self.logger.error(f"Error cleaning up dated backups: {e}")

    def verify_loaded_data(self, tracker: 'ForegroundTracker', bg_tracker: 'BackgroundVideoTracker') -> Dict[str, Any]:
        """Verify that data was loaded correctly"""
        try:
            verification = {
                'timestamp': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'app_data_loaded': False,
                'bg_data_loaded': False,
                'total_app_time': 0,
                'total_bg_time': 0,
                'app_count': 0,
                'bg_site_count': 0
            }
            
            if tracker:
                with tracker.lock:
                    app_times = dict(tracker.app_times)
                    verification['app_count'] = len(app_times)
                    verification['total_app_time'] = sum(app_times.values())
                    verification['app_data_loaded'] = len(app_times) > 0
                    
                    if app_times:
                        top_apps = sorted(app_times.items(), key=lambda x: x[1], reverse=True)[:3]
                        verification['top_apps'] = [(app[:30], f'{time:.1f}s') for app, time in top_apps]
            
            if bg_tracker:
                with bg_tracker.lock:
                    bg_times = dict(bg_tracker.background_video_times)
                    verification['bg_site_count'] = len(bg_times)
                    verification['total_bg_time'] = sum(bg_times.values())
                    verification['bg_data_loaded'] = len(bg_times) > 0
                    
                    if bg_times:
                        verification['bg_sites'] = list(bg_times.keys())
            
            return verification
            
        except Exception as e:
            self.logger.error(f"Error verifying loaded data: {e}")
            return {'error': str(e)}

    def _save_json_file(self, filepath: str, data: Dict[str, Any]):
        """Safely save data to JSON file with better error handling"""
        try:
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False, default=str)
            
            self.logger.debug(f"💾 Saved data to {filepath}")
            
        except Exception as e:
            self.logger.error(f"Error saving JSON file {filepath}: {e}")
    
    def _load_json_file(self, filepath: str) -> Dict[str, Any]:
        """Safely load data from JSON file with better error handling"""
        if not os.path.exists(filepath):
            self.logger.debug(f"📂 File does not exist: {filepath}")
            return {}
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.logger.debug(f"📂 Loaded data from {filepath}")
            return data
            
        except json.JSONDecodeError as e:
            self.logger.error(f"JSON decode error in {filepath}: {e}")
            backup_path = f"{filepath}.corrupted.{int(time.time())}"
            try:
                os.rename(filepath, backup_path)
                self.logger.info(f"🔄 Moved corrupted file to {backup_path}")
            except:
                pass
            return {}
        except Exception as e:
            self.logger.error(f"Error loading JSON file {filepath}: {e}")
            return {}


# DEBUG FUNCTION to check persistence files
def debug_persistence_files():
    """Debug function to examine persistence files"""
    print("🔍 DEBUGGING PERSISTENCE FILES")
    print("=" * 50)
    
    config = CompleteEnhancedConfig()
    persistence = CompleteEnhancedProductivityDataPersistence(config)
    
    current_date = datetime.datetime.now().strftime('%Y-%m-%d')
    print(f"Current date: {current_date}")
    
    # Check app times file
    print(f"\n📂 App times file: {persistence.app_times_file}")
    if os.path.exists(persistence.app_times_file):
        app_data = persistence._load_json_file(persistence.app_times_file)
        print(f"   File exists: ✅")
        print(f"   Date in file: {app_data.get('date', 'Missing')}")
        print(f"   Apps count: {len(app_data.get('app_times', {}))}")
        print(f"   Total time: {app_data.get('total_time', 0):.1f}s")
        
        if app_data.get('app_times'):
            print("   Top apps:")
            top_apps = sorted(app_data['app_times'].items(), key=lambda x: float(x[1]), reverse=True)[:5]
            for app, time_val in top_apps:
                print(f"     • {app[:40]}: {float(time_val):.1f}s")
    else:
        print("   File exists: ❌")
    
    # Check background video file
    print(f"\n📂 Background video file: {persistence.background_video_file}")
    if os.path.exists(persistence.background_video_file):
        bg_data = persistence._load_json_file(persistence.background_video_file)
        print(f"   File exists: ✅")
        print(f"   Date in file: {bg_data.get('date', 'Missing')}")
        print(f"   Sites count: {len(bg_data.get('background_video_times', {}))}")
        print(f"   Total BG time: {bg_data.get('total_bg_time', 0):.1f}s")
    else:
        print("   File exists: ❌")
    
    print("=" * 50)

# Uncomment to debug persistence files:
# debug_persistence_files()

class WMIConnectionState(Enum):
    DISCONNECTED = "disconnected"
    CONNECTING = "connecting"
    CONNECTED = "connected"
    FAILED = "failed"

class RobustWMIConnection:
    """Handles WMI connections with automatic reconnection on COM errors"""
    
    COM_DISCONNECTION_ERRORS = [
        -2147352567,  # E_UNEXPECTED
        -2147417848,  # RPC_E_DISCONNECTED
        -2147024809,  # ERROR_INVALID_HANDLE
        -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    ]
    
    def __init__(self, namespace="root\\cimv2", max_retries=3, retry_delay=5):
        self.namespace = namespace
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.logger = logging.getLogger(__name__)
        
        self.connection = None
        self.state = WMIConnectionState.DISCONNECTED
        self.last_connection_attempt = 0
        self.connection_failures = 0
        self.lock = threading.Lock()
        
    def _is_com_disconnection_error(self, error) -> bool:
        """Check if the error is a COM disconnection error"""
        if hasattr(error, 'com_error'):
            error_code = getattr(error.com_error, 'hresult', None)
            return error_code in self.COM_DISCONNECTION_ERRORS
            
        # Check error message for common patterns
        error_str = str(error).lower()
        disconnection_patterns = [
            'disconnected from its clients',
            'rpc server is unavailable',
            'invalid handle',
            'com error',
            'exception occurred'
        ]
        
        return any(pattern in error_str for pattern in disconnection_patterns)
    
    def _cleanup_connection(self):
        """Safely cleanup the current WMI connection"""
        try:
            if self.connection:
                self.connection = None
            pythoncom.CoUninitialize()
        except:
            pass  # Ignore cleanup errors
    
    def _establish_connection(self) -> bool:
        """Establish a fresh WMI connection"""
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            self.logger.info(f"Establishing WMI connection to {self.namespace}")
            
            # Create new WMI connection
            self.connection = wmi.WMI(namespace=self.namespace)
            
            # Test the connection with a simple query
            test_query = list(self.connection.Win32_NTLogEvent(Logfile="Security"))[:1]
            
            self.state = WMIConnectionState.CONNECTED
            self.connection_failures = 0
            self.logger.info("WMI connection established successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to establish WMI connection: {e}")
            self._cleanup_connection()
            self.state = WMIConnectionState.FAILED
            return False
    
    def connect(self) -> bool:
        """Connect to WMI with retry logic"""
        with self.lock:
            if self.state == WMIConnectionState.CONNECTED and self.connection:
                return True
            
            # Prevent too frequent reconnection attempts
            if time.time() - self.last_connection_attempt < self.retry_delay:
                return False
            
            self.last_connection_attempt = time.time()
            self.state = WMIConnectionState.CONNECTING
            
            # Clean up any existing connection
            self._cleanup_connection()
            
            # Try to establish new connection
            for attempt in range(1, self.max_retries + 1):
                self.logger.info(f"WMI connection attempt {attempt}/{self.max_retries}")
                
                if self._establish_connection():
                    return True
                
                if attempt < self.max_retries:
                    time.sleep(self.retry_delay)
            
            self.connection_failures += 1
            self.state = WMIConnectionState.FAILED
            self.logger.error(f"All WMI connection attempts failed ({self.connection_failures} total failures)")
            return False
    
    def is_connected(self) -> bool:
        """Check if WMI is currently connected"""
        return self.state == WMIConnectionState.CONNECTED and self.connection is not None
    
    def execute_query(self, query_method, *args, **kwargs):
        """Execute a WMI query with automatic reconnection on failure"""
        if not self.is_connected():
            if not self.connect():
                return None
        
        try:
            # Execute the query
            result = query_method(*args, **kwargs)
            return result
            
        except Exception as e:
            if self._is_com_disconnection_error(e):
                self.logger.warning(f"WMI disconnection detected: {e}")
                self.state = WMIConnectionState.DISCONNECTED
                
                # Try to reconnect and retry once
                if self.connect():
                    try:
                        return query_method(*args, **kwargs)
                    except Exception as retry_error:
                        self.logger.error(f"Query failed even after reconnection: {retry_error}")
                        return None
                else:
                    self.logger.error("Failed to reconnect after WMI disconnection")
                    return None
            else:
                # Non-disconnection error - log and return None
                self.logger.error(f"WMI query error: {e}")
                return None

class PowerShellFreeLoginPoller:
    """Complete PowerShell elimination - uses only native Windows APIs"""
    
    def __init__(self, max_init_time=60):
        self.logger = logging.getLogger(__name__)
        self.max_init_time = max_init_time
        
        # Event tracking
        self.last_seen: Set[tuple] = set()
        self.last_cleanup = time.time()
        self.cleanup_interval = 300
        self.session_start_time = time.time()
        
        # Detection methods (in priority order)
        self.detection_method = None
        self.event_log_handle = None
        self.wmi_connection = None
        
        self.initialization_success = False
        self._initialize()
    
    def _initialize(self):
        """Initialize best available detection method"""
        self.logger.info("Initializing PowerShell-free login detection...")
        
        # Method 1: Try native Windows Event Log API (fastest, most reliable)
        if self._init_native_eventlog():
            self.detection_method = "native_eventlog"
            self.logger.info("✅ Using native Event Log API")
            self.initialization_success = True
            return
        
        # Method 2: Try direct WMI (slower but reliable)
        if self._init_wmi():
            self.detection_method = "wmi_direct"
            self.logger.info("✅ Using direct WMI")
            self.initialization_success = True
            return
        
        # Method 3: Fallback to process monitoring
        self.detection_method = "process_monitor"
        self.logger.info("⚠️ Using process monitoring fallback")
        self.initialization_success = True
    
    def _init_native_eventlog(self) -> bool:
        """Initialize native Windows Event Log access"""
        if not NATIVE_EVENTLOG_AVAILABLE:
            return False
        
        try:
            self.event_log_handle = win32evtlog.OpenEventLog(None, "Security")
            
            # Test read access
            test_events = win32evtlog.ReadEventLog(
                self.event_log_handle,
                win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                0
            )
            
            if test_events:
                self.logger.info("✅ Native Event Log access confirmed")
                return True
            
        except Exception as e:
            self.logger.debug(f"Native Event Log init failed: {e}")
            if self.event_log_handle:
                try:
                    win32evtlog.CloseEventLog(self.event_log_handle)
                except:
                    pass
                self.event_log_handle = None
        
        return False
    
    def _init_wmi(self) -> bool:
        """Initialize direct WMI connection"""
        if not WMI_AVAILABLE:
            return False
        
        try:
            pythoncom.CoInitialize()
            self.wmi_connection = wmi.WMI(namespace="root\\cimv2")
            
            # Test query
            test_events = list(self.wmi_connection.Win32_NTLogEvent(Logfile="Security"))[:1]
            
            if test_events:
                self.logger.info("✅ Direct WMI access confirmed")
                return True
            
        except Exception as e:
            self.logger.debug(f"Direct WMI init failed: {e}")
            self.wmi_connection = None
        
        return False
    
    def poll_events(self):
        """Poll for login/logout events using active method"""
        if not self.initialization_success:
            return []
        
        self._cleanup_old_events()
        
        if self.detection_method == "native_eventlog":
            return self._poll_native_eventlog()
        elif self.detection_method == "wmi_direct":
            return self._poll_wmi_direct()
        elif self.detection_method == "process_monitor":
            return self._poll_process_monitor()
        
        return []
    
    def _poll_native_eventlog(self):
        """Poll using native Windows Event Log API"""
        if not self.event_log_handle:
            return []
        
        try:
            events = []
            cutoff_time = datetime.datetime.now() - datetime.timedelta(hours=1)
            
            # Read recent events
            event_records = win32evtlog.ReadEventLog(
                self.event_log_handle,
                win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                0
            )
            
            for event in event_records[:50]:  # Limit to recent 50 events
                if event.EventID not in [4624, 4634]:
                    continue
                
                try:
                    event_time = datetime.datetime.fromtimestamp(event.TimeGenerated.timestamp())
                except:
                    continue
                
                # Skip old events
                if event_time < cutoff_time:
                    break
                
                # Check if already seen
                key = (event.RecordNumber, event.EventID)
                if key in self.last_seen:
                    continue
                
                # Filter for user logons only
                if self._is_user_logon(event):
                    mock_event = type('MockEvent', (), {
                        'EventCode': event.EventID,
                        'TimeGenerated': event_time.strftime('%Y-%m-%d %H:%M:%S'),
                        'RecordNumber': event.RecordNumber
                    })()
                    
                    events.append(mock_event)
                    self.last_seen.add(key)
            
            if events:
                self.logger.info(f"Native API found {len(events)} login events")
            
            return events
            
        except Exception as e:
            self.logger.debug(f"Native event log polling failed: {e}")
            return []
    
    def _poll_wmi_direct(self):
        """Poll using direct WMI"""
        if not self.wmi_connection:
            return []
        
        try:
            events = []
            today = datetime.datetime.now().date()
            
            # Query login events
            try:
                login_events = list(self.wmi_connection.Win32_NTLogEvent(
                    Logfile="Security",
                    EventCode="4624"
                ))[:20]
                
                logout_events = list(self.wmi_connection.Win32_NTLogEvent(
                    Logfile="Security",
                    EventCode="4634"
                ))[:20]
                
                all_events = login_events + logout_events
                
            except Exception as e:
                # Fallback: query without filters
                self.logger.debug(f"WMI filtered query failed, trying unfiltered: {e}")
                all_events = list(self.wmi_connection.Win32_NTLogEvent(Logfile="Security"))[:100]
                all_events = [e for e in all_events if e.EventCode in ['4624', '4634']]
            
            for event in all_events:
                try:
                    # Parse WMI timestamp
                    time_str = event.TimeGenerated.split('.')[0]
                    event_time = datetime.datetime.strptime(time_str, '%Y%m%d%H%M%S')
                    
                    # Only today's events
                    if event_time.date() != today:
                        continue
                    
                    # Check if already seen
                    key = (getattr(event, 'RecordNumber', 0), event.EventCode)
                    if key in self.last_seen:
                        continue
                    
                    mock_event = type('MockEvent', (), {
                        'EventCode': int(event.EventCode),
                        'TimeGenerated': event_time.strftime('%Y-%m-%d %H:%M:%S'),
                        'RecordNumber': getattr(event, 'RecordNumber', 0)
                    })()
                    
                    events.append(mock_event)
                    self.last_seen.add(key)
                    
                except Exception as parse_error:
                    self.logger.debug(f"Error parsing WMI event: {parse_error}")
                    continue
            
            if events:
                self.logger.info(f"Direct WMI found {len(events)} login events")
            
            return events
            
        except Exception as e:
            self.logger.debug(f"Direct WMI polling failed: {e}")
            return []
    
    def _poll_process_monitor(self):
        """Fallback: Monitor for system state changes"""
        try:
            # This is a simple fallback that detects system state
            # In practice, you could monitor specific processes like winlogon.exe
            # or detect session state changes
            
            # For now, return empty - this prevents errors but doesn't detect events
            # You could enhance this to monitor process starts/stops of user session processes
            
            return []
            
        except Exception as e:
            self.logger.debug(f"Process monitor polling failed: {e}")
            return []
    
    def _is_user_logon(self, event) -> bool:
        """Check if event is a real user logon (not system/service)"""
        try:
            # Get event description
            event_desc = win32evtlogutil.SafeFormatMessage(event, None)
            if not event_desc:
                return False
            
            desc_lower = event_desc.lower()
            
            # Filter out system accounts
            system_keywords = [
                'system', 'anonymous logon', 'dcom-in', 'dwm-', 'umfd-',
                'font cache', 'local service', 'network service', '$'
            ]
            
            if any(keyword in desc_lower for keyword in system_keywords):
                return False
            
            # Look for interactive logon types
            if event.EventID == 4624:  # Logon
                # Type 2 = Interactive, Type 11 = CachedInteractive
                if any(pattern in desc_lower for pattern in ['logon type:\t\t2', 'logon type:\t\t11']):
                    return True
            elif event.EventID == 4634:  # Logoff
                return True
            
            return False
            
        except Exception as e:
            self.logger.debug(f"Error checking user logon: {e}")
            return True  # Default to including event if we can't parse it
    
    def _cleanup_old_events(self):
        """Clean up old event tracking"""
        now = time.time()
        if now - self.last_cleanup > self.cleanup_interval:
            if len(self.last_seen) > 1000:
                # Keep only recent half
                self.last_seen = set(list(self.last_seen)[-500:])
                self.logger.debug("Cleaned up old event tracking data")
            self.last_cleanup = now
    
    def get_status(self) -> dict:
        """Get current status"""
        return {
            'detection_method': self.detection_method,
            'initialization_success': self.initialization_success,
            'tracked_events': len(self.last_seen),
            'native_eventlog_available': NATIVE_EVENTLOG_AVAILABLE and self.event_log_handle is not None,
            'wmi_available': WMI_AVAILABLE and self.wmi_connection is not None,
            'powershell_free': True
        }
    
    def cleanup(self):
        """Clean up resources"""
        if self.event_log_handle:
            try:
                win32evtlog.CloseEventLog(self.event_log_handle)
            except:
                pass
            self.event_log_handle = None

class ImprovedLoginLogoutPoller:
    """PowerShell-free login/logout poller - wrapper for compatibility"""
    
    def __init__(self, max_init_time=60):
        self.logger = logging.getLogger(__name__)
        self.max_init_time = max_init_time
        
        # Use PowerShell-free implementation
        self.poller = PowerShellFreeLoginPoller(max_init_time)
        
        # Compatibility properties
        self.initialization_success = self.poller.initialization_success
        self.fallback_mode = self.poller.detection_method != "native_eventlog"
        self.query_method = self.poller.detection_method
    
    def poll_events(self):
        """Poll for events - PowerShell-free"""
        return self.poller.poll_events()
    
    def get_status(self) -> dict:
        """Get status - PowerShell-free"""
        status = self.poller.get_status()
        
        # Add compatibility fields
        status.update({
            'wmi_state': 'connected' if status.get('wmi_available') else 'disconnected',
            'fallback_mode': self.fallback_mode,
            'connection_failures': 0,  # Not applicable anymore
            'query_method': self.query_method
        })
        
        return status

class SessionTracker:
    """Tracks and manages login/logout sessions"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def parse_login_logout_events(self, events: List[str]) -> List[LoginSession]:
        """Parse login/logout events and create session objects"""
        sessions = []
        pending_login = None
        
        # Sort events by timestamp
        sorted_events = self._sort_events_by_time(events)
        
        for event in sorted_events:
            event_time, event_type = self._parse_event(event)
            
            if not event_time:
                continue
            
            if event_type == "login" or event_type == "startup":
                # Handle new login
                if pending_login:
                    # Previous login without logout - mark as incomplete
                    pending_login.session_type = "Incomplete"
                    sessions.append(pending_login)
                
                # Start new session
                session_type = "Startup" if event_type == "startup" else "Normal"
                pending_login = LoginSession(
                    login_time=event_time,
                    session_type=session_type
                )
            
            elif event_type == "logout" or event_type == "shutdown":
                if pending_login:
                    # Complete the session
                    pending_login.logout_time = event_time
                    pending_login.calculate_duration()
                    sessions.append(pending_login)
                    pending_login = None
                else:
                    # Logout without login - create orphaned logout session
                    orphaned_session = LoginSession(
                        login_time=event_time,  # Use logout time as placeholder
                        logout_time=event_time,
                        session_type="Orphaned Logout"
                    )
                    sessions.append(orphaned_session)
        
        # Handle ongoing session
        if pending_login:
            # Session still active
            pending_login.session_type = "Ongoing"
            sessions.append(pending_login)
        
        return sessions
    
    def _sort_events_by_time(self, events: List[str]) -> List[str]:
        """Sort events by timestamp"""
        def extract_time(event_str):
            # Extract timestamp from event string like "[2025-06-01 15:55:31] ..."
            try:
                timestamp_match = re.search(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]', event_str)
                if timestamp_match:
                    return datetime.datetime.strptime(timestamp_match.group(1), '%Y-%m-%d %H:%M:%S')
                return datetime.datetime.min
            except:
                return datetime.datetime.min
        
        return sorted(events, key=extract_time)
    
    def _parse_event(self, event_str: str) -> tuple[Optional[datetime.datetime], Optional[str]]:
        """Parse individual event string to extract time and type"""
        try:
            # Extract timestamp
            timestamp_match = re.search(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]', event_str)
            if not timestamp_match:
                return None, None
            
            event_time = datetime.datetime.strptime(timestamp_match.group(1), '%Y-%m-%d %H:%M:%S')
            event_str_lower = event_str.lower()
            
            # Determine event type
            if "user logged in" in event_str_lower:
                return event_time, "login"
            elif "user logged out" in event_str_lower:
                return event_time, "logout"
            elif "system startup detected" in event_str_lower:
                return event_time, "startup"
            elif "system shutdown detected" in event_str_lower:
                return event_time, "shutdown"
            
            return event_time, "unknown"
            
        except Exception as e:
            self.logger.debug(f"Error parsing event: {event_str} - {e}")
            return None, None
    
    def calculate_total_session_time(self, sessions: List[LoginSession]) -> int:
        """Calculate total time user was logged in (in seconds)"""
        total_seconds = 0
        for session in sessions:
            if session.is_complete():
                total_seconds += session.duration_seconds
        return total_seconds
    
    def get_session_summary(self, sessions: List[LoginSession]) -> dict:
        """Get summary statistics for sessions"""
        completed_sessions = [s for s in sessions if s.is_complete()]
        ongoing_sessions = [s for s in sessions if not s.is_complete() and s.session_type == "Ongoing"]
        
        total_time = self.calculate_total_session_time(sessions)
        
        return {
            'total_sessions': len(sessions),
            'completed_sessions': len(completed_sessions),
            'ongoing_sessions': len(ongoing_sessions),
            'total_logged_time': total_time,
            'average_session_duration': total_time // len(completed_sessions) if completed_sessions else 0,
            'longest_session': max([s.duration_seconds for s in completed_sessions], default=0),
            'shortest_session': min([s.duration_seconds for s in completed_sessions], default=0)
        }

class CompleteEnhancedActivityLogger:
    def __init__(self, config: CompleteEnhancedConfig):
        self.config = config
        self.log_buffer = []
        self.buffer_lock = threading.Lock()
        self.login_logout_events = []
        
        # Sent reports tracking
        self.sent_reports_file = os.path.join(config.LOG_DIR, "sent_reports.json")
        
        self._setup_enhanced_logging()
        self._cleanup_old_files()
        self._add_startup_separator()

    def _setup_enhanced_logging(self):
        """Enhanced logging with rotation and better formatting"""
        os.makedirs(self.config.LOG_DIR, exist_ok=True)
        
        # Setup rotating file handler for debug log
        debug_handler = RotatingFileHandler(
            self.config.DEBUG_LOG,
            maxBytes=self.config.MAX_LOG_SIZE_MB * 1024 * 1024,
            backupCount=self.config.LOG_BACKUP_COUNT,
            encoding='utf-8'
        )
        
        console_handler = logging.StreamHandler()
        
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        debug_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        logging.basicConfig(
            level=logging.INFO,
            handlers=[debug_handler, console_handler],
            force=True
        )
        
        self.logger = logging.getLogger(__name__)
        self._setup_activity_log_rotation()

    def _setup_activity_log_rotation(self):
        """Setup rotation for the activity log file"""
        try:
            if os.path.exists(self.config.ACTIVITY_LOG):
                size_mb = os.path.getsize(self.config.ACTIVITY_LOG) / (1024 * 1024)
                if size_mb > self.config.MAX_ACTIVITY_LOG_SIZE_MB:
                    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                    archive_name = f"{self.config.ACTIVITY_LOG}.{timestamp}.old"
                    os.rename(self.config.ACTIVITY_LOG, archive_name)
                    self.logger.info(f"📦 Archived large activity log: {archive_name}")
        except Exception as e:
            self.logger.debug(f"Activity log rotation check failed: {e}")

    def _add_startup_separator(self):
        """Add a clear separator when script starts"""
        startup_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        separator = f"""
{'='*80}
🚀 ACTIVITY MONITOR STARTED: {startup_time}
📋 LOG MANAGEMENT: Enabled (cleanup after {self.config.CLEANUP_DAYS_TO_KEEP} days)
📊 MISSED REPORTS: Enabled (recovery up to {self.config.MISSED_REPORT_DAYS_BACK} days back)
💾 DATED BACKUPS: {'Enabled' if self.config.ENABLE_DATED_BACKUPS else 'Disabled'}
{'='*80}"""
        
        self.logger.info(separator)

    def _cleanup_old_files(self):
        """Clean up files older than configured days"""
        try:
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=self.config.CLEANUP_DAYS_TO_KEEP)
            cutoff_timestamp = cutoff_date.timestamp()
            
            cleanup_patterns = [
                # Productivity reports
                os.path.join(self.config.LOG_DIR, "productivity_report_*.txt"),
                os.path.join(self.config.LOG_DIR, "MISSED_productivity_report_*.txt"),
                # Old JSON data files
                os.path.join(self.config.LOG_DIR, "app_times_data_*.json"),
                os.path.join(self.config.LOG_DIR, "background_video_data_*.json"),
                # Archived log files
                os.path.join(self.config.LOG_DIR, "*.old"),
                os.path.join(self.config.LOG_DIR, "startup_debug.log.*"),
                os.path.join(self.config.LOG_DIR, "monitor_output.log.*")
            ]
            
            cleaned_count = 0
            for pattern in cleanup_patterns:
                for filepath in glob.glob(pattern):
                    try:
                        file_time = os.path.getmtime(filepath)
                        if file_time < cutoff_timestamp:
                            os.remove(filepath)
                            cleaned_count += 1
                            self.logger.info(f"🧹 Cleaned up old file: {os.path.basename(filepath)}")
                    except Exception as e:
                        self.logger.debug(f"Could not clean {filepath}: {e}")
            
            if cleaned_count > 0:
                self.logger.info(f"🧹 Cleanup complete: {cleaned_count} old files removed")
            else:
                self.logger.debug("🧹 No old files to clean up")
                
        except Exception as e:
            self.logger.error(f"Error during file cleanup: {e}")

    # Sent Reports Tracking Methods
    def mark_report_sent(self, date: str, report_type: str = "daily"):
        """Mark a report as successfully sent"""
        try:
            sent_reports = self._load_sent_reports()
            
            if date not in sent_reports:
                sent_reports[date] = {}
            
            sent_reports[date][report_type] = {
                'sent_time': datetime.datetime.now().isoformat(),
                'timestamp': time.time()
            }
            
            # Keep only last 30 days of records
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=30)
            cutoff_str = cutoff_date.strftime('%Y-%m-%d')
            
            sent_reports = {
                date: reports for date, reports in sent_reports.items() 
                if date >= cutoff_str
            }
            
            with open(self.sent_reports_file, 'w', encoding='utf-8') as f:
                json.dump(sent_reports, f, indent=2)
            
            self.logger.info(f"📝 Marked {report_type} report for {date} as sent")
            
        except Exception as e:
            self.logger.error(f"Error marking report as sent: {e}")

    def was_report_sent(self, date: str, report_type: str = "daily") -> bool:
        """Check if a report was already sent for a specific date"""
        try:
            sent_reports = self._load_sent_reports()
            return date in sent_reports and report_type in sent_reports[date]
        except Exception as e:
            self.logger.debug(f"Error checking sent reports: {e}")
            return False

    def _load_sent_reports(self) -> dict:
        """Load the sent reports tracking file"""
        try:
            if os.path.exists(self.sent_reports_file):
                with open(self.sent_reports_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            self.logger.debug(f"Error loading sent reports: {e}")
        return {}

    def get_unsent_recent_dates(self, days_back: int = None) -> list:
        """Get list of recent dates that don't have sent reports"""
        if days_back is None:
            days_back = self.config.MISSED_REPORT_DAYS_BACK
            
        try:
            sent_reports = self._load_sent_reports()
            unsent_dates = []
            
            for i in range(1, days_back + 1):  # Start from yesterday
                check_date = datetime.datetime.now() - datetime.timedelta(days=i)
                date_str = check_date.strftime('%Y-%m-%d')
                
                if not self.was_report_sent(date_str, "daily"):
                    unsent_dates.append(date_str)
            
            return sorted(unsent_dates)
            
        except Exception as e:
            self.logger.error(f"Error getting unsent dates: {e}")
            return []

    # Original ActivityLogger methods
    def debug_log(self, message: str):
        self.logger.info(message)

    def buffer_log_entry(self, entry: str):
        with self.buffer_lock:
            clean_entry = AppNameCleaner.clean_all_exe_from_text(entry)
            self.log_buffer.append(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {clean_entry}")

    def buffer_login_logout_event(self, event: str):
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with self.buffer_lock:
            clean_event = AppNameCleaner.clean_all_exe_from_text(event)
            formatted_event = f"[{timestamp}] {clean_event}"
            self.log_buffer.append(formatted_event)
            self.login_logout_events.append(formatted_event)

    def get_recent_login_logout_events(self) -> List[str]:
        with self.buffer_lock:
            return self.login_logout_events.copy()

    def flush_buffer(self):
        with self.buffer_lock:
            if self.log_buffer:
                try:
                    with open(self.config.ACTIVITY_LOG, "a", encoding="utf-8") as f:
                        f.write("\n".join(self.log_buffer) + "\n")
                    self.log_buffer.clear()
                except IOError as e:
                    self.logger.error(f"Failed to write to activity log: {e}")

class ConfigManager:
    def __init__(self, config_path: str):
        self.config_path = config_path
        self.logger = logging.getLogger(__name__)

    def get_config_value(self, key: str, default: str = "") -> str:
        if not os.path.isfile(self.config_path):
            self.logger.warning(f"Config file not found: {self.config_path}")
            return default

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith(f"{key}="):
                        return line.split("=", 1)[1].strip()
        except IOError as e:
            self.logger.error(f"Error reading config file: {e}")

        self.logger.info(f"Config key '{key}' not found. Using default.")
        return default

    def get_email_config(self) -> Optional[str]:
        email = self.get_config_value("to_email", "")
        if not email or "@" not in email:
            self.logger.warning("Invalid or missing 'to_email' in config.txt.")
            return None
        return email

    def is_friday_only_enabled(self) -> bool:
        """Check if Friday-only mode is enabled from config file"""
        friday_only = self.get_config_value("friday_only", "false").lower()
        return friday_only in ['true', '1', 'yes', 'on']

# Enhanced ConfigManager with new email timing methods
class EnhancedConfigManager(ConfigManager):
    """Enhanced ConfigManager with email timing support"""
    
    def __init__(self, config_path: str):
        super().__init__(config_path)
        self.email_timing = ImprovedEmailTiming(self)
    
    def get_email_timing_status(self) -> dict:
        """Get comprehensive email timing status"""
        base_status = self.email_timing.get_timing_status()
        base_status.update({
            'config_file': self.config_path,
            'config_exists': os.path.exists(self.config_path),
            'to_email': self.get_email_config()
        })
        return base_status
    
    def log_timing_configuration(self):
        """Log the current timing configuration for debugging"""
        status = self.get_email_timing_status()
        
        logger = logging.getLogger(__name__)
        logger.info("📧 IMPROVED EMAIL TIMING CONFIGURATION:")
        logger.info(f"   Mode: {status['mode']}")
        
        if status['mode'] == 'interval':
            logger.info(f"   Interval: {status['interval_formatted']} ({status['interval_seconds']}s)")
        elif status['mode'] == 'time_of_day':
            logger.info(f"   Daily time: {status['daily_time']}")
            if status['last_daily_email_date']:
                logger.info(f"   Last sent: {status['last_daily_email_date']}")
        
        logger.info(f"   Next email: {status['next_email_info']}")
        logger.info(f"   To email: {status['to_email'] or 'Not configured'}")


class AppCategorizer:
    PRODUCTIVE_KEYWORDS = [
        'word', 'excel', 'powerpoint', 'onenote', 'outlook',
        'teams', 'onedrive', 'sharepoint', 'access', 'publisher',
        'visio', 'project', 'visual studio', 'github', 'stackoverflow'
    ]

     # UPDATED: Better shopping site patterns
    UNPRODUCTIVE_KEYWORDS = [
        # Social Media (specific)
        'facebook.com', 'instagram.com', 'twitter.com', 'tiktok.com', 
        'snapchat.com', 'pinterest.com', 'reddit.com',
        
        # Video/Entertainment (specific)
        'youtube.com', 'netflix.com', 'hulu.com', 'disney.com', 
        'twitch.tv', 'vimeo.com', 'dailymotion.com',
        
        # Gaming (specific)
        'steam.com', 'epic games', 'xbox.com', 'playstation.com',
        'runescape', 'oldschool runescape', 'osrs', 'jagex',
        
        # UPDATED: Shopping sites (broader patterns)
        'amazon.com/dp/', 'amazon.com/gp/product', 'amazon.com/Mineral', 'amazon.com/[A-Z]',  # Amazon product pages
        'ebay.com/itm', 'ebay.com/p/', 'ebay.com/b/',  # eBay listings and categories
        'etsy.com/listing', 'etsy.com/shop/',  # Etsy products and shops
        'walmart.com/ip/', 'target.com/p/', 'bestbuy.com/site/',  # Other major retailers
        
        # Music/Audio (specific)
        'spotify.com', 'soundcloud.com', 'pandora.com'
    ]
    
    # ADDED: System processes that should be ignored
    SYSTEM_PROCESSES = [
        'explorer.exe', 'notepad.exe', 'taskmgr.exe', 'dwm.exe',
        'shellexperiencehost.exe', 'searchui.exe', 'startmenuexperiencehost.exe',
        'cortana.exe', 'winlogon.exe', 'csrss.exe', 'smss.exe',
        'lsass.exe', 'services.exe', 'svchost.exe', 'conhost.exe',
        'audiodg.exe', 'spoolsv.exe', 'searchindexer.exe',
        'applicationframehost.exe', 'runtimebroker.exe',
        'backgroundtaskhost.exe', 'taskhostw.exe', 'sihost.exe',
        'ctfmon.exe', 'fontdrvhost.exe', 'unsecapp.exe',
        'searchprotocolhost.exe', 'searchfilterhost.exe'
    ]

    @classmethod
    def categorize_app(cls, name: str) -> Optional[Category]:
        """ENHANCED: Better categorization with shopping site detection"""
        if not name:
            return None
            
        name_lower = name.lower()

        # Check if it's a system process first
        exe_name = name_lower.split(' - ')[0].strip() if ' - ' in name_lower else name_lower
        if any(sys_proc in exe_name for sys_proc in cls.SYSTEM_PROCESSES):
            return None

        # Check if it's a browser with website
        if cls.is_browser_with_website(name):
            page_content = cls._extract_page_content_from_browser(name)
            
            # Amazon detection
            if 'amazon.com' in page_content.lower():
                return Category.UNPRODUCTIVE
            
            # Gaming detection
            gaming_indicators = [
                'runescape', 'osrs', 'oldschool runescape', 'jagex',
                'world of warcraft', 'wow', 'league of legends', 'valorant',
                'fortnite', 'minecraft', 'steam', 'epic games'
            ]
            
            if any(indicator in page_content.lower() for indicator in gaming_indicators):
                return Category.UNPRODUCTIVE
            
            # General categorization
            if cls._is_clearly_productive(page_content):
                return Category.PRODUCTIVE
            elif cls._is_clearly_unproductive(page_content):
                return Category.UNPRODUCTIVE
            else:
                return Category.UNCATEGORIZED

        # For non-browser apps
        for keyword in cls.PRODUCTIVE_KEYWORDS:
            if keyword in name_lower:
                return Category.PRODUCTIVE
        
        for keyword in cls.UNPRODUCTIVE_KEYWORDS:
            if keyword in name_lower:
                return Category.UNPRODUCTIVE
            
        return None

    @classmethod
    def is_system_process(cls, app_name: str) -> bool:
        """Check if an app is a Windows system process"""
        if not app_name:
            return False
            
        app_lower = app_name.lower()
        exe_name = app_lower.split(' - ')[0].strip() if ' - ' in app_lower else app_lower
        
        return any(sys_proc in exe_name for sys_proc in cls.SYSTEM_PROCESSES)
    
    @classmethod
    def _extract_page_content_from_browser(cls, full_title: str) -> str:
        """Extract meaningful content from browser window title"""
        if not full_title:
            return ""
            
        content = full_title
        
        # Remove common browser suffixes
        browser_suffixes = [
            '— Mozilla Firefox', '- Mozilla Firefox', 
            '- Google Chrome', '- Microsoft Edge',
            '- Safari'
        ]
        
        for suffix in browser_suffixes:
            content = content.replace(suffix, '')
        
        # Remove leading browser name
        if content.startswith('Mozilla Firefox - '):
            content = content[18:]  # Remove "Mozilla Firefox - "
        elif content.startswith('Google Chrome - '):
            content = content[16:]
        elif content.startswith('Microsoft Edge - '):
            content = content[17:]
        
        return content.strip()

    @classmethod
    def _is_clearly_productive(cls, page_content: str) -> bool:
        """Check if page content is clearly work-related"""
        if not page_content:
            return False
            
        content_lower = page_content.lower()
        
        productive_indicators = [
            # Work platforms
            'microsoft teams', 'slack', 'zoom', 'google workspace',
            'office 365', 'sharepoint', 'onedrive',
            
            # Development
            'github', 'gitlab', 'stack overflow', 'stackoverflow',
            'visual studio', 'documentation', 'developer',
            
            # Professional services
            'linkedin', 'salesforce', 'jira', 'confluence',
            
            # Email/Calendar
            'gmail', 'outlook', 'calendar', 'email'
        ]
        
        return any(indicator in content_lower for indicator in productive_indicators)

    @classmethod  
    def _is_clearly_unproductive(cls, page_content: str) -> bool:
        """ENHANCED: Better shopping and entertainment detection"""
        if not page_content:
            return False
            
        content_lower = page_content.lower()
        
        unproductive_indicators = [
            # Social Media (exact matches)
            'facebook', 'instagram', 'twitter', 'tiktok', 
            'snapchat', 'reddit', 'pinterest',
            
            # Video Entertainment  
            'youtube', 'netflix', 'hulu', 'disney+', 'twitch',
            
            # Gaming
            'steam', 'epic games', 'gaming', 'game',
            'runescape', 'oldschool runescape', 'osrs',
            
            # ADDED: Shopping indicators
            'amazon.com', 'amazon shopping', 'buy now', 'add to cart',
            'ebay', 'etsy', 'walmart', 'target', 'best buy',
            'price', 'shipping', 'reviews', 'product details',
            
            # Entertainment News
            'entertainment', 'celebrity', 'gossip', 'memes'
        ]
        
        # ENHANCED: Check for Amazon-specific patterns
        amazon_patterns = [
            '/dp/', '/gp/product', 'amazon.com',
            'add to cart', 'buy now', 'prime delivery',
            'customer reviews', 'product details'
        ]
        
        if any(pattern in content_lower for pattern in amazon_patterns):
            return True
        
        return any(indicator in content_lower for indicator in unproductive_indicators)


    @classmethod
    def is_browser_with_website(cls, app_title: str) -> bool:
        """Check if this is a browser window with actual web content"""
        if not app_title:
            return False
            
        browser_indicators = [
            'mozilla firefox', 'google chrome', 'microsoft edge',
            'safari', 'opera', 'brave'
        ]
        
        title_lower = app_title.lower()
        has_browser = any(indicator in title_lower for indicator in browser_indicators)
        
        if not has_browser:
            return False
        
        # Exclude empty tabs and browser-only windows
        exclude_patterns = [
            'new tab', 'about:blank', 'chrome://newtab',
            'edge://newtab', 'about:newtab'
        ]
        
        if any(pattern in title_lower for pattern in exclude_patterns):
            return False
        
        # Must have actual content (indicated by " - " separator)
        has_content = ' - ' in app_title and len(app_title.split(' - ')) >= 2
        
        return has_content

class AudioDetector:
    """Detects which processes are currently playing audio"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.available = AUDIO_DETECTION_AVAILABLE

    def get_audio_playing_processes(self) -> Dict[int, Dict]:
        """Get PIDs of processes currently playing audio with their info"""
        if not self.available:
            return {}

        audio_processes = {}

        try:
            # Get all audio sessions
            sessions = AudioUtilities.GetAllSessions()

            for session in sessions:
                if session.Process:
                    pid = session.Process.pid
                    process_name = session.Process.name()

                    # Check if session is currently playing audio
                    if session.SimpleAudioVolume and session.State == 1:  # AudioSessionStateActive
                        audio_processes[pid] = {
                            'name': process_name,
                            'volume': session.SimpleAudioVolume.GetMasterVolume(),
                            'muted': session.SimpleAudioVolume.GetMute()
                        }

        except Exception as e:
            self.logger.debug(f"Error detecting audio processes: {e}")

        return audio_processes

class SystemMonitor:
    VIDEO_STREAMING_SITES = [
        'youtube', 'youtu.be', 'netflix', 'hulu', 'disney', 'twitch', 'vimeo',
        'tiktok', 'instagram', 'facebook', 'twitter', 'dailymotion', 'vevo',
        'crunchyroll', 'funimation', 'amazon prime', 'paramount', 'peacock',
        'hbo max', 'apple tv', 'spotify', 'soundcloud', 'pandora'
    ]

    BROWSER_PROCESSES = [
        'chrome.exe', 'firefox.exe', 'msedge.exe', 'opera.exe',
        'brave.exe', 'vivaldi.exe', 'safari.exe', 'iexplore.exe'
    ]

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.debug_counter = 0
        self.last_successful_app = None
        self.permission_issues_logged = False

    def get_foreground_process_name(self) -> Optional[str]:
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name().lower()
        except (psutil.AccessDenied, psutil.NoSuchProcess, win32gui.error) as e:
            self.logger.debug(f"Error getting foreground process: {e}")
            return None

    def get_foreground_app_with_title(self) -> str:
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return "Unknown"

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid <= 0:
                self.logger.debug(f"Invalid PID received: {pid}")
                return "Unknown"

            proc = psutil.Process(pid)
            name = proc.name()
            title = win32gui.GetWindowText(hwnd)
            return f"{name} - {title}"
        except (psutil.AccessDenied, psutil.NoSuchProcess, win32gui.error, ValueError) as e:
            self.logger.debug(f"Error getting foreground app with title: {e}")
            return "Unknown"

    def get_clean_foreground_app_with_title(self) -> str:
        """Enhanced version with debugging and fallback methods"""
        self.debug_counter += 1
        
        # Try multiple methods to get foreground app info
        result = self._try_get_foreground_app_method1()
        if result != "Unknown":
            if self.debug_counter % 100 == 0:  # Log success every 100 iterations
                self.logger.info(f"✅ Successfully tracking app: {result[:50]}...")
            self.last_successful_app = result
            return result
        
        # Method 2: Try with different approach
        result = self._try_get_foreground_app_method2()
        if result != "Unknown":
            if self.debug_counter % 100 == 0:
                self.logger.info(f"✅ Method 2 success: {result[:50]}...")
            self.last_successful_app = result
            return result
        
        # Method 3: Try basic window enumeration
        result = self._try_get_foreground_app_method3()
        if result != "Unknown":
            if self.debug_counter % 100 == 0:
                self.logger.info(f"✅ Method 3 success: {result[:50]}...")
            self.last_successful_app = result
            return result
        
        # Log detailed debug info every 50 iterations if still failing
        if self.debug_counter % 50 == 0:
            self._log_detailed_debug_info()
        
        return "Unknown"

    def _try_get_foreground_app_method1(self) -> str:
        """Original method with enhanced error handling"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                self.logger.debug("Method 1: No foreground window handle")
                return "Unknown"

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid <= 0:
                self.logger.debug(f"Method 1: Invalid PID: {pid}")
                return "Unknown"

            try:
                proc = psutil.Process(pid)
                name = proc.name()
                title = win32gui.GetWindowText(hwnd)
                
                # Clean the app name before combining
                clean_name = AppNameCleaner.clean_app_name(name)
                result = f"{clean_name} - {title}" if title else clean_name
                
                if result and result != " - ":
                    return result
                    
            except psutil.AccessDenied:
                # Try to get at least the window title
                title = win32gui.GetWindowText(hwnd)
                if title:
                    return f"Protected Process - {title}"
                return "Protected Process"
            except psutil.NoSuchProcess:
                title = win32gui.GetWindowText(hwnd)
                if title:
                    return f"Terminated Process - {title}"
                return "Terminated Process"
                
        except Exception as e:
            self.logger.debug(f"Method 1 failed: {e}")
        
        return "Unknown"

    def _try_get_foreground_app_method2(self) -> str:
        """Alternative method using different Windows APIs"""
        try:
            # Get the foreground window
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return "Unknown"
            
            # Get window title first (this usually works even when process access fails)
            title = win32gui.GetWindowText(hwnd)
            
            # Get window class name as backup identifier
            class_name = win32gui.GetClassName(hwnd)
            
            # Try to get process name
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid > 0:
                    proc = psutil.Process(pid)
                    process_name = AppNameCleaner.clean_app_name(proc.name())
                    
                    if title:
                        return f"{process_name} - {title}"
                    else:
                        return process_name
            except:
                # If process access fails, use title and class name
                if title:
                    return f"Unknown Process - {title}"
                elif class_name:
                    return f"Window Class: {class_name}"
                    
        except Exception as e:
            self.logger.debug(f"Method 2 failed: {e}")
        
        return "Unknown"

    def _try_get_foreground_app_method3(self) -> str:
        """Fallback method using window enumeration"""
        try:
            current_foreground = win32gui.GetForegroundWindow()
            if not current_foreground:
                return "Unknown"
            
            def window_callback(hwnd, results):
                if hwnd == current_foreground and win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    
                    if title:
                        results.append(f"Active Window - {title}")
                    elif class_name:
                        results.append(f"Window: {class_name}")
                return True
            
            results = []
            win32gui.EnumWindows(window_callback, results)
            
            if results:
                return results[0]
                
        except Exception as e:
            self.logger.debug(f"Method 3 failed: {e}")
        
        return "Unknown"

    def _log_detailed_debug_info(self):
        """Log detailed debugging information"""
        if not self.permission_issues_logged:
            self.logger.warning("🔍 DEBUGGING FOREGROUND APP DETECTION ISSUES")
            self.logger.warning("=" * 60)
            
            # Test basic Windows API access
            try:
                hwnd = win32gui.GetForegroundWindow()
                self.logger.warning(f"✅ GetForegroundWindow(): {hwnd}")
                
                if hwnd:
                    title = win32gui.GetWindowText(hwnd)
                    self.logger.warning(f"✅ Window title: '{title}'")
                    
                    class_name = win32gui.GetClassName(hwnd)
                    self.logger.warning(f"✅ Window class: '{class_name}'")
                    
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        self.logger.warning(f"✅ Process ID: {pid}")
                        
                        if pid > 0:
                            try:
                                proc = psutil.Process(pid)
                                self.logger.warning(f"✅ Process name: {proc.name()}")
                                self.logger.warning(f"✅ Process exe: {proc.exe()}")
                            except psutil.AccessDenied:
                                self.logger.warning("❌ Process access denied - may need admin privileges")
                            except psutil.NoSuchProcess:
                                self.logger.warning("❌ Process no longer exists")
                            except Exception as pe:
                                self.logger.warning(f"❌ Process error: {pe}")
                    except Exception as we:
                        self.logger.warning(f"❌ Window process error: {we}")
                else:
                    self.logger.warning("❌ No foreground window detected")
                    
            except Exception as e:
                self.logger.warning(f"❌ Basic Windows API test failed: {e}")
            
            # Test current user context
            try:
                import getpass
                current_user = getpass.getuser()
                self.logger.warning(f"🔍 Running as user: {current_user}")
                
                # Check if running as administrator
                try:
                    import ctypes
                    is_admin = ctypes.windll.shell32.IsUserAnAdmin()
                    self.logger.warning(f"🔍 Running as admin: {is_admin}")
                except:
                    self.logger.warning("🔍 Cannot determine admin status")
                    
            except Exception as ue:
                self.logger.warning(f"❌ User context error: {ue}")
            
            # Show recommendations
            self.logger.warning("\n💡 RECOMMENDATIONS:")
            self.logger.warning("1. Try running the script as Administrator")
            self.logger.warning("2. Check Windows Defender/Antivirus blocking")
            self.logger.warning("3. Ensure the script has proper permissions")
            self.logger.warning("4. Test with different user accounts")
            
            if self.last_successful_app:
                self.logger.warning(f"🔍 Last successful detection: {self.last_successful_app}")
            
            self.logger.warning("=" * 60)
            self.permission_issues_logged = True

    def test_basic_functionality(self) -> Dict[str, Any]:
        """Test basic functionality and return results"""
        test_results = {
            'timestamp': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'tests_passed': 0,
            'tests_failed': 0,
            'details': []
        }
        
        # Test 1: Basic Windows API
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                test_results['tests_passed'] += 1
                test_results['details'].append(f"✅ GetForegroundWindow: {hwnd}")
            else:
                test_results['tests_failed'] += 1
                test_results['details'].append("❌ GetForegroundWindow returned None")
        except Exception as e:
            test_results['tests_failed'] += 1
            test_results['details'].append(f"❌ GetForegroundWindow failed: {e}")
        
        # Test 2: Process enumeration
        try:
            process_count = len(list(psutil.process_iter()))
            test_results['tests_passed'] += 1
            test_results['details'].append(f"✅ Process enumeration: {process_count} processes")
        except Exception as e:
            test_results['tests_failed'] += 1
            test_results['details'].append(f"❌ Process enumeration failed: {e}")
        
        # Test 3: Current foreground app
        try:
            current_app = self.get_clean_foreground_app_with_title()
            if current_app != "Unknown":
                test_results['tests_passed'] += 1
                test_results['details'].append(f"✅ Current app detected: {current_app[:50]}...")
            else:
                test_results['tests_failed'] += 1
                test_results['details'].append("❌ Cannot detect current app")
        except Exception as e:
            test_results['tests_failed'] += 1
            test_results['details'].append(f"❌ App detection failed: {e}")
        
        # Test 4: Admin privileges
        try:
            import ctypes
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if is_admin:
                test_results['tests_passed'] += 1
                test_results['details'].append("✅ Running with admin privileges")
            else:
                test_results['tests_failed'] += 1
                test_results['details'].append("❌ Not running with admin privileges")
        except Exception as e:
            test_results['tests_failed'] += 1
            test_results['details'].append(f"❌ Admin check failed: {e}")
        
        return test_results

    def get_all_browser_windows(self) -> List[Dict[str, str]]:
        browser_windows = []
        foreground_hwnd = win32gui.GetForegroundWindow()

        def enum_window_callback(hwnd, windows):
            if not win32gui.IsWindowVisible(hwnd):
                return True

            title = win32gui.GetWindowText(hwnd)
            if not title.strip():
                return True

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid <= 0:
                return True

            try:
                proc = psutil.Process(pid)
                process_name = proc.name().lower()
                if process_name in self.BROWSER_PROCESSES:
                    windows.append({
                        'hwnd': hwnd,
                        'process_name': process_name,
                        'title': title,
                        'is_foreground': hwnd == foreground_hwnd,
                        'pid': pid
                    })
            except (psutil.AccessDenied, psutil.NoSuchProcess, win32gui.error) as e:
                self.logger.debug(f"Error processing window {hwnd}: {e}")

            return True

        try:
            win32gui.EnumWindows(enum_window_callback, browser_windows)
        except Exception as e:
            self.logger.error(f"Error enumerating windows: {e}")

        return browser_windows

    def detect_background_video_activity(self) -> List[Dict[str, str]]:
        browser_windows = self.get_all_browser_windows()
        background_video_windows = []

        for window in browser_windows:
            if window['is_foreground']:
                continue

            title_lower = window['title'].lower()
            detected_sites = [site for site in self.VIDEO_STREAMING_SITES if site in title_lower]

            if detected_sites:
                background_video_windows.append({
                    'process_name': window['process_name'],
                    'title': window['title'],
                    'detected_sites': detected_sites,
                    'hwnd': window['hwnd']
                })

        return background_video_windows


# ADD this test function to debug the system monitor
def test_system_monitor():
    """Test function to debug SystemMonitor functionality"""
    print("🔍 TESTING SYSTEM MONITOR FUNCTIONALITY")
    print("=" * 60)
    
    monitor = SystemMonitor()
    
    # Run basic functionality test
    test_results = monitor.test_basic_functionality()
    
    print(f"⏰ Test time: {test_results['timestamp']}")
    print(f"✅ Tests passed: {test_results['tests_passed']}")
    print(f"❌ Tests failed: {test_results['tests_failed']}")
    print("\n📋 Test Details:")
    
    for detail in test_results['details']:
        print(f"   {detail}")
    
    print(f"\n🎯 Current foreground app test:")
    for i in range(5):
        current_app = monitor.get_clean_foreground_app_with_title()
        print(f"   Attempt {i+1}: {current_app}")
        time.sleep(1)
    
    if test_results['tests_failed'] > test_results['tests_passed']:
        print("\n🚨 ISSUES DETECTED:")
        print("1. Try running this script as Administrator")
        print("2. Check if antivirus is blocking the script")
        print("3. Ensure you have necessary Windows permissions")
        print("4. Try restarting the script")
    else:
        print("\n✅ System Monitor appears to be working correctly!")

# Uncomment and run this to test:
# test_system_monitor()

class BackgroundVideoTracker(threading.Thread):
    """Simplified background video tracker with audio detection only"""

    VIDEO_STREAMING_SITES = [
        'youtube', 'youtu.be', 'netflix', 'hulu', 'disney', 'twitch', 'vimeo',
        'tiktok', 'instagram', 'facebook', 'twitter', 'dailymotion', 'vevo',
        'crunchyroll', 'funimation', 'amazon prime', 'paramount', 'peacock',
        'hbo max', 'apple tv', 'spotify', 'soundcloud', 'pandora'
    ]

    BROWSER_PROCESSES = [
        'chrome.exe', 'firefox.exe', 'msedge.exe', 'opera.exe',
        'brave.exe', 'vivaldi.exe', 'safari.exe', 'iexplore.exe'
    ]

    def __init__(self, config, activity_logger):
        super().__init__(daemon=True)
        self.config = config
        self.activity_logger = activity_logger
        self.logger = logging.getLogger(__name__)
        self.audio_detector = AudioDetector()

        # Track background video sessions
        self.active_background_videos: Dict[str, Dict] = {}
        self.background_video_times: Dict[str, float] = {}
        self.verified_playing_times: Dict[str, float] = {}
        self.lock = threading.Lock()
        self.shutdown_event = threading.Event()

        # Log detection capabilities
        capabilities = []
        if self.audio_detector.available:
            capabilities.append("audio detection")
        else:
            capabilities.append("basic window detection")

        self.logger.info(f"Background video tracking enabled with: {', '.join(capabilities)}")

    def run(self):
        self.logger.info("BackgroundVideoTracker thread starting...")

        while not self.shutdown_event.is_set():
            try:
                self._update_background_video_tracking()
                self.shutdown_event.wait(self.config.SLEEP_INTERVAL)
            except Exception as e:
                self.logger.error(f"Error in background video tracking: {e}")
                self.shutdown_event.wait(5)

    def _update_background_video_tracking(self):
        """Update background video tracking with audio verification"""
        current_background_videos = self._get_current_background_videos()
        audio_playing_processes = self.audio_detector.get_audio_playing_processes()
        now = time.time()

        with self.lock:
            current_window_ids = set()

            for video_info in current_background_videos:
                window_id = f"{video_info['hwnd']}_{video_info['process_name']}"
                current_window_ids.add(window_id)
                site_name = self._extract_site_name(video_info['title'], video_info['detected_sites'])

                # Audio detection
                is_playing_audio = video_info['pid'] in audio_playing_processes

                if window_id not in self.active_background_videos:
                    # New background video started
                    session_id = self._generate_session_id(video_info)
                    self.active_background_videos[window_id] = {
                        'site': site_name,
                        'start_time': now,
                        'browser': video_info['process_name'],
                        'title': video_info['title'],
                        'last_update': now,
                        'pid': video_info['pid'],
                        'is_playing': is_playing_audio,
                        'session_id': session_id
                    }

                    # Log detection
                    status = "playing (audio)" if is_playing_audio else "paused/silent"
                    clean_browser_name = AppNameCleaner.clean_app_name(video_info['process_name'])
                    self.activity_logger.buffer_log_entry(
                        f"Background video detected: {site_name} ({clean_browser_name}) - {status}"
                    )
                else:
                    # Update existing background video
                    session = self.active_background_videos[window_id]
                    elapsed = now - session['last_update']
                    session['last_update'] = now

                    # Log state changes
                    old_playing = session.get('is_playing', False)
                    if old_playing != is_playing_audio:
                        if is_playing_audio:
                            self.activity_logger.buffer_log_entry(f"{site_name} started playing (audio detected)")
                        else:
                            self.activity_logger.buffer_log_entry(f"{site_name} stopped playing")

                    session['is_playing'] = is_playing_audio

                    # Always track total background time
                    if site_name not in self.background_video_times:
                        self.background_video_times[site_name] = 0.0
                    self.background_video_times[site_name] += elapsed

                    # Track verified playing time (with audio)
                    if is_playing_audio:
                        if site_name not in self.verified_playing_times:
                            self.verified_playing_times[site_name] = 0.0
                        self.verified_playing_times[site_name] += elapsed

            # Handle stopped background videos
            stopped_videos = set(self.active_background_videos.keys()) - current_window_ids
            for window_id in stopped_videos:
                session = self.active_background_videos[window_id]
                total_duration = now - session['start_time']
                verified_time = self.verified_playing_times.get(session['site'], 0)

                log_parts = [f"Total: {int(total_duration)}s"]
                if verified_time > 0:
                    log_parts.append(f"Playing: {int(verified_time)}s")

                self.activity_logger.buffer_log_entry(
                    f"Background video stopped: {session['site']} ({', '.join(log_parts)})"
                )
                del self.active_background_videos[window_id]

    def _get_current_background_videos(self) -> List[Dict]:
        """Get currently active background videos with PID info"""
        browser_windows = []
        foreground_hwnd = win32gui.GetForegroundWindow()

        def enum_window_callback(hwnd, windows):
            if not win32gui.IsWindowVisible(hwnd):
                return True

            title = win32gui.GetWindowText(hwnd)
            if not title.strip():
                return True

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid <= 0:
                return True

            try:
                proc = psutil.Process(pid)
                process_name = proc.name().lower()
                if process_name in self.BROWSER_PROCESSES:
                    windows.append({
                        'hwnd': hwnd,
                        'process_name': process_name,
                        'title': title,
                        'is_foreground': hwnd == foreground_hwnd,
                        'pid': pid
                    })
            except (psutil.AccessDenied, psutil.NoSuchProcess) as e:
                self.logger.debug(f"Error processing window {hwnd}: {e}")

            return True

        try:
            win32gui.EnumWindows(enum_window_callback, browser_windows)
        except Exception as e:
            self.logger.error(f"Error enumerating windows: {e}")

        # Filter for background video windows
        background_video_windows = []
        for window in browser_windows:
            if window['is_foreground']:
                continue

            title_lower = window['title'].lower()
            detected_sites = [site for site in self.VIDEO_STREAMING_SITES if site in title_lower]

            if detected_sites:
                background_video_windows.append({
                    'process_name': window['process_name'],
                    'title': window['title'],
                    'detected_sites': detected_sites,
                    'hwnd': window['hwnd'],
                    'pid': window['pid']
                })

        return background_video_windows

    def _extract_site_name(self, title: str, detected_sites: List[str]) -> str:
        """Extract clean site name"""
        if detected_sites:
            primary_site = detected_sites[0]

            site_mapping = {
                'youtube': 'YouTube',
                'youtu.be': 'YouTube',
                'netflix': 'Netflix',
                'hulu': 'Hulu',
                'disney': 'Disney+',
                'twitch': 'Twitch',
                'tiktok': 'TikTok',
                'instagram': 'Instagram',
                'facebook': 'Facebook',
                'twitter': 'Twitter/X',
                'spotify': 'Spotify',
                'soundcloud': 'SoundCloud',
                'amazon prime': 'Amazon Prime'
            }

            return site_mapping.get(primary_site, primary_site.title())

        return "Unknown Video Site"

    def _generate_session_id(self, video_info: Dict) -> str:
        """Generate a unique session ID for the video"""
        return f"{video_info['process_name']}_{video_info['pid']}_{time.time()}"

    def get_background_video_times(self) -> Dict[str, int]:
        """Get total background video times (including paused)"""
        with self.lock:
            now = time.time()
            for session in self.active_background_videos.values():
                site_name = session['site']
                elapsed = now - session['last_update']
                session['last_update'] = now

                if site_name not in self.background_video_times:
                    self.background_video_times[site_name] = 0.0
                self.background_video_times[site_name] += elapsed

                # Update verified playing time if currently playing
                if session.get('is_playing', False):
                    if site_name not in self.verified_playing_times:
                        self.verified_playing_times[site_name] = 0.0
                    self.verified_playing_times[site_name] += elapsed

            return {site: int(seconds) for site, seconds in self.background_video_times.items()}

    def get_verified_playing_times(self) -> Dict[str, int]:
        """Get times when videos were actually playing (with audio)"""
        with self.lock:
            return {site: int(seconds) for site, seconds in self.verified_playing_times.items()}

    def get_total_background_video_time(self) -> int:
        """Get total background video time"""
        background_times = self.get_background_video_times()
        return sum(background_times.values())

    def get_total_verified_playing_time(self) -> int:
        """Get total verified playing time"""
        verified_times = self.get_verified_playing_times()
        return sum(verified_times.values())

    def stop(self):
        """Stop the background video tracker"""
        self.shutdown_event.set()
        self.join(timeout=5)

        now = time.time()
        with self.lock:
            for session in self.active_background_videos.values():
                total_duration = now - session['start_time']
                self.activity_logger.buffer_log_entry(
                    f"Background video stopped (shutdown): {session['site']} (Duration: {int(total_duration)}s)"
                )

class ForegroundTracker(threading.Thread):
    def __init__(self, config: CompleteEnhancedConfig, activity_logger: CompleteEnhancedActivityLogger):
        super().__init__(daemon=True)
        self.config = config
        self.activity_logger = activity_logger
        self.system_monitor = SystemMonitor()
        self.logger = logging.getLogger(__name__)

        self.app_times: Dict[str, float] = {}
        self.current_key: Optional[str] = None
        self.current_start = time.time()
        self.lock = threading.Lock()
        self.shutdown_event = threading.Event()

        self.last_unproductive_title: Optional[str] = None
        self.unproductive_start_time: Optional[float] = None

    def run(self):
        self.logger.info("ForegroundTracker thread starting...")
        iteration = 0

        while not self.shutdown_event.is_set():
            try:
                iteration += 1
                active_key = self.system_monitor.get_clean_foreground_app_with_title()

                if iteration % 10 == 0:
                    self.logger.info(f"Tracking iteration #{iteration}, current app: {active_key}")
                    with self.lock:
                        self.logger.info(f"Total tracked apps so far: {len(self.app_times)}")

                self._track_current_app()
                self.shutdown_event.wait(self.config.SLEEP_INTERVAL)
            except Exception as e:
                self.logger.error(f"Error in tracking loop: {e}")
                self.shutdown_event.wait(5)

    def _track_current_app(self):
        active_key = self.system_monitor.get_clean_foreground_app_with_title()
        now = time.time()

        self._handle_unproductive_tracking(active_key, now)
        self._update_app_times(active_key, now)

    def _handle_unproductive_tracking(self, active_key: str, now: float):
        category = AppCategorizer.categorize_app(active_key)
        is_unproductive = (category == Category.UNPRODUCTIVE)

        if is_unproductive and self.last_unproductive_title != active_key:
            self.last_unproductive_title = active_key
            self.unproductive_start_time = now
            # Clean the app name before logging
            clean_name = AppNameCleaner.clean_app_base_name(active_key)
            self.activity_logger.buffer_log_entry(f"Unproductive tab opened: {clean_name}")

        elif not is_unproductive and self.last_unproductive_title:
            duration = now - self.unproductive_start_time if self.unproductive_start_time else 0
            # Clean the app name before logging
            clean_name = AppNameCleaner.clean_app_base_name(self.last_unproductive_title)
            self.activity_logger.buffer_log_entry(
                f"Unproductive tab closed: {clean_name} (Open for {int(duration)}s)"
            )
            self.last_unproductive_title = None
            self.unproductive_start_time = None

    def _update_app_times(self, active_key: str, now: float):
        """Enhanced app time tracking with system process filtering"""
        
        # Filter out system processes before tracking
        if AppCategorizer.is_system_process(active_key):
            # Don't track system processes, but don't reset current_key either
            # This prevents system processes from interrupting real app tracking
            return
        
        # Filter out very short-lived or invalid entries
        if not active_key or active_key == "Unknown" or len(active_key.strip()) < 3:
            return
        
        with self.lock:
            if active_key != self.current_key:
                # Save time for previous app if it was valid
                if self.current_key and not AppCategorizer.is_system_process(self.current_key):
                    elapsed = now - self.current_start
                    # Only track if the elapsed time is meaningful (at least 1 second)
                    if elapsed >= 1.0:
                        self.app_times[self.current_key] = self.app_times.get(self.current_key, 0) + elapsed

                self.current_key = active_key
                self.current_start = now

    def stop(self):
        self.shutdown_event.set()
        self.join(timeout=5)

        now = time.time()
        with self.lock:
            if self.current_key:
                elapsed = now - self.current_start
                self.app_times[self.current_key] = self.app_times.get(self.current_key, 0) + elapsed

        if self.last_unproductive_title and self.unproductive_start_time:
            duration = now - self.unproductive_start_time
            # Clean the app name before logging
            clean_name = AppNameCleaner.clean_app_base_name(self.last_unproductive_title)
            self.activity_logger.buffer_log_entry(
                f"Unproductive tab closed: {clean_name} (Open for {int(duration)}s)"
            )

    def get_app_times(self) -> Dict[str, float]:
        with self.lock:
            return {app: round(secs) for app, secs in self.app_times.items()}

class ActivityReporter:
    def __init__(self, config: CompleteEnhancedConfig, activity_logger: CompleteEnhancedActivityLogger):
        self.config = config
        self.activity_logger = activity_logger
        self.last_logged_times: Dict[str, int] = {}
        self.logger = logging.getLogger(__name__)

    def log_activity(self, tracker: ForegroundTracker):
        self.logger.info("=== STARTING log_activity ===")
        try:
            app_times = tracker.get_app_times()
            self.logger.info(f"Got app_times: {len(app_times)} items")

            categorized = self._categorize_app_times(app_times)
            self.logger.info("=== FINISHED _categorize_app_times ===")

            has_new_activity = self._has_new_activity(categorized)
            self.logger.info(f"Has new activity: {has_new_activity}")

            if has_new_activity:
                self.logger.info("About to write activity logs...")
                self._write_activity_logs(categorized)
            else:
                self.logger.info("No new activity detected, skipping log write")

            self._check_background_video_activity()
            self.logger.info("=== FINISHED log_activity ===")

        except Exception as e:
            self.logger.error(f"Error in log_activity: {e}")
            self.logger.error(f"Full traceback: {traceback.format_exc()}")

    def _categorize_app_times(self, app_times: Dict[str, float]) -> Dict[Category, Dict[str, List[Tuple[str, int]]]]:
        categorized = {
            Category.PRODUCTIVE: {},
            Category.UNPRODUCTIVE: {},
            Category.UNCATEGORIZED: {}  # ADDED: Include uncategorized
        }

        self.logger.info(f"Processing {len(app_times)} apps for categorization")

        for app, secs in app_times.items():
            secs = int(secs)
            if secs == 0:
                continue

            prev_secs = self.last_logged_times.get(app, -1)

            if secs > prev_secs:
                category = AppCategorizer.categorize_app(app)

                if category is None:
                    continue

                # Clean the base app name
                base = AppNameCleaner.clean_app_base_name(app)
                title = app.replace(app.split(" - ")[0] + " - ", "").strip() if " - " in app else ""

                if base not in categorized[category]:
                    categorized[category][base] = []

                categorized[category][base].append((title, secs))
                self.last_logged_times[app] = secs

        return categorized

    def _has_new_activity(self, categorized: Dict[Category, Dict[str, List[Tuple[str, int]]]]) -> bool:
        return any(categorized[category] for category in categorized)

    def _write_activity_logs(self, categorized: Dict[Category, Dict[str, List[Tuple[str, int]]]]):
        now = datetime.datetime.now()
        log_time = now.strftime("%Y-%m-%d %H:%M")

        try:
            with open(self.config.ACTIVITY_LOG, "a", encoding="utf-8") as log:
                header = f"=== Activity Log: {log_time} ===\n"
                log.write(header)

                # UPDATE THIS LINE to include UNCATEGORIZED:
                totals = {Category.PRODUCTIVE: 0, Category.UNPRODUCTIVE: 0, Category.UNCATEGORIZED: 0}

                # UPDATE THIS LINE to include all three categories:
                for category in [Category.PRODUCTIVE, Category.UNPRODUCTIVE, Category.UNCATEGORIZED]:
                    self._write_category_section(log, category, categorized[category], totals)

                footer = self._generate_summary_footer(totals)
                
                # Clean .exe extensions from the footer before writing
                clean_footer = AppNameCleaner.clean_all_exe_from_text(footer)
                log.write(clean_footer)

        except IOError as e:
            self.logger.error(f"Error writing activity logs: {e}")

    def _write_category_section(self, log, category: Category, apps: Dict[str, List[Tuple[str, int]]], totals: Dict[Category, int]):
        if not apps:
            return
        
        section_header = f"[{category.value} Applications]\n"
        log.write(section_header)

        for base, windows in apps.items():
            total_app_time = sum(t[1] for t in windows)
            totals[category] += total_app_time

            # base is already cleaned by AppNameCleaner in _categorize_app_times
            app_line = f"{base} - Total Time: {total_app_time} seconds\n"
            log.write(app_line)

            for title, sec in sorted(windows, key=lambda x: -x[1]):
                warn = " WARNING" if sec >= self.config.UNPRODUCTIVE_WARNING_THRESHOLD and category == Category.UNPRODUCTIVE else ""
                detail_line = f"    * {title} ({sec}s){warn}\n"
                log.write(detail_line)

    def _generate_summary_footer(self, totals: Dict[Category, int]) -> str:
        return f"""Summary:
Productive: {totals[Category.PRODUCTIVE]}s
Unproductive: {totals[Category.UNPRODUCTIVE]}s
Uncategorized: {totals[Category.UNCATEGORIZED]}s
=== End of Log ===

"""

    def _check_background_video_activity(self):
        system_monitor = SystemMonitor()
        background_videos = system_monitor.detect_background_video_activity()

        if background_videos:
            for video_window in background_videos:
                sites_detected = ", ".join(video_window['detected_sites'])
                message = (f"Background video detected: {video_window['process_name']} - "
                          f"{video_window['title']} (Sites: {sites_detected})")

                self.activity_logger.buffer_log_entry(message)
                self.logger.info(message)


# ADD THIS NEW CLASS BEFORE HybridOutlookManager

class ImprovedStoreOutlookDetector:
    """Enhanced detection for Microsoft Store Outlook with multiple search methods"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def _check_your_specific_outlook(self) -> Optional[Dict]:
        """Check for your specific Outlook installation path"""
        try:
            # Your specific installation pattern
            base_pattern = r"C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_*"
            
            matches = glob.glob(base_pattern)
            self.logger.debug(f"Checking pattern: {base_pattern}")
            self.logger.debug(f"Found matches: {matches}")
            
            for match in matches:
                if os.path.isdir(match):
                    self.logger.debug(f"Scanning directory: {match}")
                    
                    # Look specifically for olk.exe and olkBg.exe
                    for exe_name in ['olk.exe', 'olkBg.exe']:
                        exe_path = os.path.join(match, exe_name)
                        
                        if os.path.exists(exe_path) and os.access(exe_path, os.X_OK):
                            self.logger.info(f"🎯 Found your specific Outlook: {exe_path}")
                            
                            return {
                                'path': match,
                                'executable': exe_path,
                                'package_name': os.path.basename(match),
                                'type': 'store',
                                'detection_method': 'specific_user_path',
                                'executable_name': exe_name
                            }
            
            # Also check if the exact path exists
            exact_path = r"C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2025.522.100_x64__8wekyb3d8bbwe"
            if os.path.exists(exact_path):
                self.logger.debug(f"Found exact path: {exact_path}")
                
                for exe_name in ['olk.exe', 'olkBg.exe']:
                    exe_path = os.path.join(exact_path, exe_name)
                    
                    if os.path.exists(exe_path):
                        self.logger.info(f"🎯 Found exact Outlook installation: {exe_path}")
                        
                        return {
                            'path': exact_path,
                            'executable': exe_path,
                            'package_name': os.path.basename(exact_path),
                            'type': 'store',
                            'detection_method': 'exact_user_path',
                            'executable_name': exe_name
                        }
            
        except Exception as e:
            self.logger.debug(f"Specific Outlook check failed: {e}")
        
        return None

    
    def find_store_outlook_comprehensive(self) -> Optional[Dict]:
        """Comprehensive Store Outlook detection - YOUR INSTALLATION FIRST"""
        
        # Method 0: Check your specific installation FIRST
        store_outlook = self._check_your_specific_outlook()
        if store_outlook:
            self.logger.info(f"✅ Found via specific user path: {store_outlook['executable']}")
            return store_outlook
        
        # Method 1: Check Windows Apps directory patterns
        store_outlook = self._check_windows_apps_directory()
        if store_outlook:
            return store_outlook
        
        # Method 2: Check Windows Registry for installed apps
        store_outlook = self._check_registry_installed_apps()
        if store_outlook:
            return store_outlook
        
        # Method 3: Check WindowsApps via PowerShell (SILENT)
        store_outlook = self._check_via_powershell()
        if store_outlook:
            return store_outlook
        
        # Method 4: Check common Store app locations
        store_outlook = self._check_common_store_locations()
        if store_outlook:
            return store_outlook
        
        # Method 5: Check for new Outlook (Microsoft 365)
        store_outlook = self._check_new_outlook_365()
        if store_outlook:
            return store_outlook
            
        return None
    
    def _check_windows_apps_directory(self) -> Optional[Dict]:
        """Enhanced WindowsApps directory search"""
        try:
            # Multiple search patterns for different Outlook versions
            search_patterns = [
                # New Outlook patterns
                "Microsoft.OutlookForWindows_*",
                "*OutlookForWindows*", 
                "Microsoft.Outlook_*",
                "*Microsoft.Outlook*",
                
                # Legacy patterns
                "microsoft.windowscommunicationsapps_*",
                "*WindowsCommunicationsApps*",
                
                # Office 365 patterns
                "Microsoft.Office.Desktop_*",
                "*Office.Desktop*"
            ]
            
            base_paths = [
                os.path.expandvars(r"%ProgramFiles%\WindowsApps"),
                os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps"),
                os.path.expandvars(r"%ProgramFiles(x86)%\WindowsApps")
            ]
            
            for base_path in base_paths:
                if not os.path.exists(base_path):
                    continue
                
                for pattern in search_patterns:
                    search_path = os.path.join(base_path, pattern)
                    matches = glob.glob(search_path)
                    
                    for match in matches:
                        if os.path.isdir(match):
                            outlook_info = self._scan_directory_for_outlook(match)
                            if outlook_info:
                                self.logger.info(f"🏪 Found Store Outlook via directory scan: {outlook_info['executable']}")
                                return outlook_info
            
        except Exception as e:
            self.logger.debug(f"Directory search failed: {e}")
        
        return None
    
    def _scan_directory_for_outlook(self, directory: str) -> Optional[Dict]:
        """Scan a directory for Outlook executables"""
        outlook_executables = [
        # Your specific Outlook executables
        'olk.exe',
        'olkbg.exe',
        
        # Standard Store Outlook executables  
        'hxoutlook.exe',
        'outlook.exe', 
        'outlookforwindows.exe',
        'hxmail.exe',
        'mail.exe',
        'microsoft.outlook.exe'
    ]
        
        try:
            for root, dirs, files in os.walk(directory):
                for file in files:
                    if file.lower() in outlook_executables:
                        executable_path = os.path.join(root, file)
                        
                        # Verify it's actually executable
                        if os.access(executable_path, os.X_OK):
                            return {
                                'path': directory,
                                'executable': executable_path,
                                'package_name': os.path.basename(directory),
                                'type': 'store',
                                'detection_method': 'directory_scan'
                            }
        except Exception as e:
            self.logger.debug(f"Error scanning directory {directory}: {e}")
        
        return None
    
    def _check_registry_installed_apps(self) -> Optional[Dict]:
        """Check Windows Registry for installed Store apps"""
        try:
            registry_paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications",
                r"SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages"
            ]
            
            outlook_patterns = [
                "microsoft.outlookforwindows",
                "microsoft.outlook", 
                "microsoft.office.desktop",
                "microsoft.windowscommunicationsapps"
            ]
            
            for registry_path in registry_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path) as key:
                        i = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(key, i)
                                
                                # Check if this looks like an Outlook package
                                if any(pattern in subkey_name.lower() for pattern in outlook_patterns):
                                    # Try to get more details about this package
                                    package_info = self._get_package_details_from_registry(registry_path, subkey_name)
                                    if package_info:
                                        return package_info
                                
                                i += 1
                            except OSError:
                                break
                                
                except Exception as e:
                    self.logger.debug(f"Registry path {registry_path} failed: {e}")
                    continue
                    
        except Exception as e:
            self.logger.debug(f"Registry search failed: {e}")
        
        return None
    
    def _get_package_details_from_registry(self, base_path: str, package_name: str) -> Optional[Dict]:
        """Get detailed package information from registry"""
        try:
            package_path = f"{base_path}\\{package_name}"
            
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, package_path) as package_key:
                try:
                    install_location, _ = winreg.QueryValueEx(package_key, "PackageRootFolder")
                    
                    if os.path.exists(install_location):
                        outlook_info = self._scan_directory_for_outlook(install_location)
                        if outlook_info:
                            outlook_info['detection_method'] = 'registry'
                            return outlook_info
                            
                except FileNotFoundError:
                    # Try alternative registry value names
                    alt_names = ["InstallLocation", "Path", "PackageFullName"]
                    for alt_name in alt_names:
                        try:
                            location, _ = winreg.QueryValueEx(package_key, alt_name)
                            if os.path.exists(location):
                                outlook_info = self._scan_directory_for_outlook(location)
                                if outlook_info:
                                    outlook_info['detection_method'] = 'registry_alt'
                                    return outlook_info
                        except FileNotFoundError:
                            continue
                            
        except Exception as e:
            self.logger.debug(f"Package details extraction failed: {e}")
        
        return None
    
    def _check_via_powershell(self) -> Optional[Dict]:
        """Use PowerShell to find Store Outlook - COMPLETELY SILENT"""
        try:
            ps_commands = [
                "Get-AppxPackage | Where-Object {$_.Name -like '*Outlook*'} | Select-Object Name,InstallLocation",
                "Get-AppxPackage | Where-Object {$_.Name -like '*Mail*'} | Select-Object Name,InstallLocation",
                "Get-AppxPackage | Where-Object {$_.Name -like '*Office*'} | Select-Object Name,InstallLocation"
            ]
            
            for ps_command in ps_commands:
                try:
                    # COMPLETELY SILENT POWERSHELL EXECUTION
                    result = subprocess.run(
                        [
                            "powershell.exe",
                            "-WindowStyle", "Hidden",
                            "-NoProfile",
                            "-NonInteractive",
                            "-NoLogo",
                            "-ExecutionPolicy", "Bypass",
                            "-Command", ps_command
                        ],
                        capture_output=True,
                        text=True,
                        timeout=30,
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        shell=False,
                        stdin=subprocess.DEVNULL,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.DEVNULL
                    )
                    
                    if result.returncode == 0 and result.stdout:
                        outlook_info = self._parse_powershell_output(result.stdout)
                        if outlook_info:
                            outlook_info['detection_method'] = 'powershell_silent'
                            return outlook_info
                            
                except subprocess.TimeoutExpired:
                    self.logger.debug("PowerShell command timed out (silent)")
                except Exception as e:
                    self.logger.debug(f"PowerShell command failed (silent): {e}")
                    
        except Exception as e:
            self.logger.debug(f"PowerShell detection failed (silent): {e}")
        
        return None
    
    def _parse_powershell_output(self, output: str) -> Optional[Dict]:
        """Parse PowerShell Get-AppxPackage output"""
        try:
            lines = output.strip().split('\n')
            
            name = None
            install_location = None
            
            for line in lines:
                if line.strip().startswith('Name'):
                    name = line.split(':', 1)[1].strip() if ':' in line else None
                elif line.strip().startswith('InstallLocation'):
                    install_location = line.split(':', 1)[1].strip() if ':' in line else None
            
            if name and install_location and os.path.exists(install_location):
                outlook_info = self._scan_directory_for_outlook(install_location)
                if outlook_info:
                    outlook_info['package_name'] = name
                    return outlook_info
                    
        except Exception as e:
            self.logger.debug(f"PowerShell output parsing failed: {e}")
        
        return None
    
    def _check_common_store_locations(self) -> Optional[Dict]:
        """Check common Store app installation locations"""
        common_locations = [
            os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps\microsoft.outlook_*"),
            os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps\*outlook*"),
            os.path.expandvars(r"%AppData%\Microsoft\Outlook"),
            os.path.expandvars(r"%LocalAppData%\Packages\Microsoft.OutlookForWindows_*"),
            os.path.expandvars(r"%LocalAppData%\Packages\*OutlookForWindows*"),
        ]
        
        for location_pattern in common_locations:
            try:
                matches = glob.glob(location_pattern)
                for match in matches:
                    if os.path.isdir(match):
                        outlook_info = self._scan_directory_for_outlook(match)
                        if outlook_info:
                            outlook_info['detection_method'] = 'common_locations'
                            return outlook_info
            except Exception as e:
                self.logger.debug(f"Common location search failed for {location_pattern}: {e}")
        
        return None
    
    def _check_new_outlook_365(self) -> Optional[Dict]:
        """Check for new Outlook (Microsoft 365) installation"""
        try:
            # Check for Office 365 Outlook installation
            office_locations = [
                os.path.expandvars(r"%ProgramFiles%\Microsoft Office\root\Office16"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\root\Office16"),
                os.path.expandvars(r"%ProgramFiles%\Microsoft Office\Office16"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\Office16"),
            ]
            
            for location in office_locations:
                outlook_exe = os.path.join(location, "OUTLOOK.EXE")
                if os.path.exists(outlook_exe):
                    return {
                        'path': location,
                        'executable': outlook_exe,
                        'package_name': 'Microsoft Office 365',
                        'type': 'desktop',
                        'detection_method': 'office_365'
                    }
        
        except Exception as e:
            self.logger.debug(f"Office 365 detection failed: {e}")
        
        return None

class HybridOutlookManager:
    """Hybrid Outlook manager: Store Outlook for UI, Desktop Outlook for email sending"""
    
    def __init__(self, config: CompleteEnhancedConfig, to_email: Optional[str]):
        self.config = config
        self.to_email = to_email
        self.logger = logging.getLogger(__name__)
        
        # Track which Outlook versions are available
        self.store_outlook_info = None
        self.desktop_outlook_available = False
        self.store_outlook_launched = False
        
        # Initialize detection
        self._detect_outlook_installations()

    def _detect_outlook_installations(self):
        """Detect both Store and Desktop Outlook installations"""
        self.logger.info("🔍 Detecting Outlook installations for hybrid setup...")
        
        # Detect Store Outlook
        self.store_outlook_info = self._find_store_outlook_installation()
        
        # Detect Desktop Outlook
        self.desktop_outlook_available = self._check_desktop_outlook()
        
        if self.store_outlook_info and self.desktop_outlook_available:
            self.logger.info("✅ HYBRID SETUP POSSIBLE: Both Store and Desktop Outlook detected")
            self.logger.info(f"   📱 Store UI: {self.store_outlook_info.get('path', 'Unknown')}")
            self.logger.info(f"   🖥️ Desktop Email: Available for COM interface")
        elif self.desktop_outlook_available:
            self.logger.info("⚡ Desktop-only setup: Using Office 365 Outlook")
        elif self.store_outlook_info:
            self.logger.info("📱 Store-only setup: Using Microsoft Store Outlook")
        else:
            self.logger.warning("❌ No Outlook installations detected")

    def _find_store_outlook_installation(self) -> Optional[Dict]:
        """Enhanced Store Outlook detection using comprehensive search"""
        detector = ImprovedStoreOutlookDetector()
        return detector.find_store_outlook_comprehensive()

    def _check_desktop_outlook(self) -> bool:
        """Check if desktop Outlook is available for COM interface"""
        try:
            # Try to create Outlook COM object
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # Test basic functionality
            namespace = outlook.GetNamespace("MAPI")
            version = getattr(outlook, 'Version', 'Unknown')
            
            self.logger.info(f"🖥️ Desktop Outlook available (Version: {version})")
            return True
            
        except Exception as e:
            self.logger.debug(f"Desktop Outlook check failed: {e}")
            return False

    def launch_store_outlook_ui(self) -> bool:
        """Launch Store Outlook for user interface - UPDATED FOR YOUR INSTALLATION"""
        if not self.store_outlook_info:
            self.logger.warning("📱 Store Outlook not available for UI")
            return False
            
        try:
            executable = self.store_outlook_info['executable']
            executable_name = self.store_outlook_info.get('executable_name', 'olk.exe')
            
            self.logger.info(f"🚀 Launching Store Outlook UI: {executable}")
            
            # Launch Store Outlook with proper flags for UI display
            if executable_name.lower() == 'olk.exe':
                # Use olk.exe for main UI
                process = subprocess.Popen(
                    [executable],
                    shell=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NEW_CONSOLE
                )
            elif executable_name.lower() == 'olkbg.exe':
                # Try olk.exe first if olkBg.exe was detected
                olk_path = executable.replace('olkBg.exe', 'olk.exe')
                if os.path.exists(olk_path):
                    self.logger.info(f"🔄 Using olk.exe instead of olkBg.exe for UI: {olk_path}")
                    process = subprocess.Popen(
                        [olk_path],
                        shell=False,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NEW_CONSOLE
                    )
                else:
                    # Fallback to olkBg.exe
                    process = subprocess.Popen(
                        [executable],
                        shell=False,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NEW_CONSOLE
                    )
            else:
                # Other executables
                process = subprocess.Popen(
                    [executable],
                    shell=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NEW_CONSOLE
                )
            
            # Give it time to start and check if it's running
            import time
            time.sleep(3)
            
            # Verify the process is still running
            if process.poll() is None:
                self.store_outlook_launched = True
                self.logger.info("✅ Store Outlook UI launched successfully")
                return True
            else:
                self.logger.warning("⚠️ Store Outlook process exited immediately")
                return False
            
        except Exception as e:
            self.logger.error(f"❌ Failed to launch Store Outlook UI: {e}")
            return False

    def send_email_via_desktop_background(self, subject: str, body: str, attachment_path: str) -> bool:
        """Send email via desktop Outlook COM interface in background"""
        if not self.desktop_outlook_available:
            self.logger.error("❌ Desktop Outlook not available for email sending")
            return False
            
        if not self.to_email:
            self.logger.warning("❌ No recipient email configured")
            return False

        if not os.path.exists(attachment_path):
            self.logger.error(f"❌ Attachment not found: {attachment_path}")
            return False

        try:
            self.logger.info("📧 Sending email via desktop Outlook (background)...")
            
            # Connect to desktop Outlook COM interface
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # Create email
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = self.to_email
            mail.Subject = subject
            mail.Body = body
            
            # Add attachment
            mail.Attachments.Add(Source=attachment_path)
            
            # Send email
            mail.Send()
            
            self.logger.info("✅ Email sent successfully via desktop Outlook")
            self._update_email_timestamp()
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Desktop Outlook email sending failed: {e}")
            return False

    def setup_hybrid_environment(self) -> Dict[str, bool]:
        """Set up the hybrid environment: Always launch Store UI when available"""
        results = {
            'store_ui_launched': False,
            'desktop_email_ready': False,
            'hybrid_ready': False
        }
        
        self.logger.info("🔧 Setting up ACTIVE hybrid Outlook environment...")
        
        # Step 1: Always launch Store Outlook for UI (if available)
        if self.store_outlook_info:
            if self.is_store_outlook_running():
                self.logger.info("📱 Store Outlook is already running - using existing instance")
                results['store_ui_launched'] = True
                self.store_outlook_launched = True
            else:
                self.logger.info("🚀 Launching Store Outlook UI for immediate use...")
                results['store_ui_launched'] = self.launch_store_outlook_ui()
            
        # Step 2: Verify desktop Outlook is ready for email sending
        if self.desktop_outlook_available:
            try:
                # Test desktop Outlook connection
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                
                results['desktop_email_ready'] = True
                self.logger.info("✅ Desktop Outlook ready for background email sending")
                
            except Exception as e:
                self.logger.warning(f"⚠️ Desktop Outlook preparation failed: {e}")
        
        # Step 3: Determine if hybrid setup is successful
        results['hybrid_ready'] = (
            results['desktop_email_ready'] and  # Must have email capability
            (results['store_ui_launched'] or not self.store_outlook_info)  # UI if available
        )
        
        if results['hybrid_ready']:
            if results['store_ui_launched']:
                self.logger.info("🎉 ACTIVE HYBRID SETUP COMPLETE: Store UI active + Desktop Email ready")
            else:
                self.logger.info("🎉 DESKTOP SETUP COMPLETE: Desktop Email ready (Store UI unavailable)")
        else:
            self.logger.warning("⚠️ Hybrid setup incomplete - some features may not work")
            
        return results
    
    def _get_available_outlook_methods(self) -> List[Tuple[str, Dict]]:
        """Compatibility method - returns available email methods for logging"""
        methods = []
        
        if self.store_outlook_info and self.desktop_outlook_available:
            methods.append(("Hybrid (Store UI + Desktop Email)", {
                'type': 'hybrid',
                'store_path': self.store_outlook_info.get('path'),
                'desktop_available': True
            }))
        elif self.desktop_outlook_available:
            methods.append(("Desktop Outlook Only", {
                'type': 'desktop',
                'version': 'Available'
            }))
        elif self.store_outlook_info:
            methods.append(("Store Outlook Only", {
                'type': 'store',
                'path': self.store_outlook_info.get('path')
            }))
        
        return methods

    def send_email_via_outlook(self, subject: str, body: str, attachment_path: str) -> bool:
        """Backward compatibility method - uses hybrid approach"""
        return self.send_email_hybrid(subject, body, attachment_path)

    def send_email_hybrid(self, subject: str, body: str, attachment_path: str) -> bool:
        """Send email using TRUE hybrid approach: Launch Store UI + Send via Desktop"""
        
        self.logger.info("📧 Starting HYBRID email process...")
        
        # Step 1: Launch Store Outlook UI ONLY if not already running
        ui_launched = False
        if self.store_outlook_info:
            # Check if Store Outlook is already running
            if self.is_store_outlook_running():
                self.logger.info("📱 Store Outlook is already running - no need to launch")
                ui_launched = True
                self.store_outlook_launched = True
            else:
                self.logger.info("🚀 Step 1: Launching Store Outlook UI...")
                ui_launched = self.launch_store_outlook_ui()
            
            if ui_launched:
                self.logger.info("✅ Store Outlook UI is now open")
                # Give UI time to fully load
                import time
                time.sleep(2)
            else:
                self.logger.warning("⚠️ Store Outlook UI failed to launch, continuing with desktop-only")
        else:
            self.logger.info("📱 No Store Outlook available, using desktop-only approach")
        
        # Step 2: Send email via desktop Outlook (more reliable)
        self.logger.info("📨 Step 2: Sending email via desktop Outlook...")
        email_sent = self.send_email_via_desktop_background(subject, body, attachment_path)
        
        if email_sent:
            self.logger.info("✅ Email sent successfully via desktop Outlook")
            
            # Step 3: Show hybrid success notification
            if ui_launched:
                self.logger.info("🎉 HYBRID SUCCESS: Store UI opened + Email sent via desktop")
            else:
                self.logger.info("📧 DESKTOP SUCCESS: Email sent (Store UI unavailable)")
                self._show_email_sent_notification(subject)
        else:
            self.logger.error("❌ Email sending failed")
            
            # If email failed but UI launched, show the UI anyway
            if ui_launched:
                self.logger.info("📱 Store Outlook UI is available for manual email sending")
        
        return email_sent

    def is_store_outlook_running(self) -> bool:
        """Check if Store Outlook is currently running - ENHANCED"""
        if not self.store_outlook_info:
            return False
            
        try:
            executable_name = self.store_outlook_info.get('executable_name', 'olk.exe')
            executable_path = self.store_outlook_info.get('executable', '')
            
            # Check for running processes
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    # Check by process name
                    if proc.info['name'] and proc.info['name'].lower() == executable_name.lower():
                        self.logger.debug(f"Found Store Outlook process by name: {proc.info['name']} (PID: {proc.info['pid']})")
                        return True
                        
                    # Check by executable path
                    if proc.info['exe'] and executable_path.lower() in proc.info['exe'].lower():
                        self.logger.debug(f"Found Store Outlook process by path: {proc.info['exe']} (PID: {proc.info['pid']})")
                        return True
                        
                    # Check for any Outlook Store processes
                    if proc.info['exe'] and 'outlookforwindows' in proc.info['exe'].lower():
                        self.logger.debug(f"Found Store Outlook process: {proc.info['exe']} (PID: {proc.info['pid']})")
                        return True
                        
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
                    
        except Exception as e:
            self.logger.debug(f"Error checking if Store Outlook is running: {e}")
        
        return False

    def get_outlook_status(self) -> Dict[str, Any]:
        """Get current status of Outlook installations and setup"""
        return {
            'store_outlook_available': self.store_outlook_info is not None,
            'store_outlook_path': self.store_outlook_info.get('path') if self.store_outlook_info else None,
            'store_outlook_launched': self.store_outlook_launched,
            'desktop_outlook_available': self.desktop_outlook_available,
            'hybrid_capable': self.store_outlook_info is not None and self.desktop_outlook_available,
            'email_method': 'hybrid' if (self.store_outlook_info and self.desktop_outlook_available) else 'desktop_only' if self.desktop_outlook_available else 'none'
        }

    def should_send_productivity_report(self, last_email_time: float) -> bool:
        """Check if it's time to send productivity report"""
        if not self.to_email:
            return False
        return time.time() - last_email_time >= self.config.EMAIL_INTERVAL

    def _update_email_timestamp(self):
        """Update the last email timestamp"""
        try:
            with open(self.config.EMAIL_TRACK_FILE, "w") as f:
                f.write(str(time.time()))
        except IOError as e:
            self.logger.error(f"Failed to update email timestamp: {e}")

    def get_last_email_time(self) -> float:
        """Get the timestamp of the last sent email"""
        if os.path.exists(self.config.EMAIL_TRACK_FILE):
            try:
                with open(self.config.EMAIL_TRACK_FILE, "r") as f:
                    return float(f.read().strip())
            except (IOError, ValueError) as e:
                self.logger.warning(f"Error reading email timestamp: {e}")
        return 0.0

    def test_hybrid_setup(self) -> Dict[str, Any]:
        """Test the hybrid setup without sending actual email"""
        self.logger.info("🧪 TESTING HYBRID OUTLOOK SETUP")
        
        test_results = {
            'store_detection': False,
            'desktop_detection': False,
            'store_launch': False,
            'desktop_email_test': False,
            'overall_success': False
        }
        
        # Test Store detection
        if self.store_outlook_info:
            test_results['store_detection'] = True
            self.logger.info(f"✅ Store Outlook detected: {self.store_outlook_info['executable']}")
            
            # Test Store launch
            if self.launch_store_outlook_ui():
                test_results['store_launch'] = True
                self.logger.info("✅ Store Outlook launched successfully")
        else:
            self.logger.info("📱 Store Outlook not detected")
        
        # Test Desktop detection and email capability
        if self.desktop_outlook_available:
            test_results['desktop_detection'] = True
            self.logger.info("✅ Desktop Outlook detected")
            
            # Test email creation (don't send)
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = self.to_email or "test@example.com"
                mail.Subject = "[TEST] Hybrid Setup Test"
                mail.Body = "This is a test email (not sent)"
                
                # Test attachment capability with a temporary file
                import tempfile
                with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
                    f.write("Test attachment")
                    temp_file = f.name
                
                mail.Attachments.Add(Source=temp_file)
                
                # Don't send, just verify creation worked
                test_results['desktop_email_test'] = True
                self.logger.info("✅ Desktop email creation test passed")
                
                # Clean up
                os.unlink(temp_file)
                
            except Exception as e:
                self.logger.error(f"❌ Desktop email test failed: {e}")
        else:
            self.logger.info("🖥️ Desktop Outlook not detected")
        
        # Overall success evaluation
        test_results['overall_success'] = (
            test_results['desktop_detection'] and 
            test_results['desktop_email_test']
        )
        
        if test_results['overall_success']:
            if test_results['store_detection']:
                self.logger.info("🎉 HYBRID SETUP TEST: SUCCESS (Store UI + Desktop Email)")
            else:
                self.logger.info("🎉 DESKTOP SETUP TEST: SUCCESS (Desktop Only)")
        else:
            self.logger.warning("⚠️ SETUP TEST: FAILED - Email functionality not available")
        
        return test_results
    

    def debug_outlook_detection(self) -> Dict[str, any]:
        """Debug Outlook detection issues with comprehensive search"""
        detector = ImprovedStoreOutlookDetector()
        
        debug_info = {
            'timestamp': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'search_methods_tested': [],
            'installations_found': [],
            'search_paths_checked': [],
            'search_paths_exist': {},
            'errors_encountered': []
        }
        
        # Test each detection method individually
        methods = [
            ('Windows Apps Directory', detector._check_windows_apps_directory),
            ('Registry Search', detector._check_registry_installed_apps),
            ('PowerShell Query', detector._check_via_powershell),
            ('Common Locations', detector._check_common_store_locations),
            ('Office 365 Detection', detector._check_new_outlook_365)
        ]
        
        for method_name, method_func in methods:
            try:
                self.logger.info(f"🔍 Testing {method_name}...")
                result = method_func()
                debug_info['search_methods_tested'].append(method_name)
                
                if result:
                    debug_info['installations_found'].append({
                        'method': method_name,
                        'result': result
                    })
                    self.logger.info(f"✅ {method_name} found: {result.get('executable', 'Unknown path')}")
                else:
                    self.logger.info(f"❌ {method_name} found nothing")
                    
            except Exception as e:
                error_msg = f"{method_name}: {str(e)}"
                debug_info['errors_encountered'].append(error_msg)
                self.logger.error(f"💥 {method_name} error: {e}")
        
        # Check search paths
        search_paths = [
            os.path.expandvars(r"%ProgramFiles%\WindowsApps"),
            os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps"),
            os.path.expandvars(r"%LocalAppData%\Packages"),
            os.path.expandvars(r"%ProgramFiles%\Microsoft Office\root\Office16"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\root\Office16")
        ]
        
        debug_info['search_paths_checked'] = search_paths
        for path in search_paths:
            debug_info['search_paths_exist'][path] = os.path.exists(path)
        
        return debug_info

# Enhanced HybridOutlookManager with new timing system
class EnhancedHybridOutlookManager(HybridOutlookManager):
    """Enhanced Outlook manager with flexible email timing"""
    
    def __init__(self, config: CompleteEnhancedConfig, config_manager: EnhancedConfigManager):
        # Get email from enhanced config manager
        to_email = config_manager.get_email_config()
        super().__init__(config, to_email)
        
        self.config_manager = config_manager
        self.email_timing = config_manager.email_timing

        # Track when script started to prevent immediate startup emails
        self.script_start_time = time.time()    
        
        # Log timing configuration
        self.config_manager.log_timing_configuration()
    
    def should_send_productivity_report(self, last_email_time: float) -> Tuple[bool, str]:
        """Enhanced report timing check with detailed reason"""
        if not self.to_email:
            return False, "No email address configured"
        
        should_send, reason = self.email_timing.should_send_email_now(last_email_time)
        
        if should_send:
            self.logger.info(f"📧 Time to send report: {reason}")
        else:
            self.logger.debug(f"📧 Not sending report: {reason}")
        
        return should_send, reason
    
    def send_email_with_timing_update(self, subject: str, body: str, attachment_path: str) -> bool:
        """Send email and update timing state"""
        success = self.send_email_hybrid(subject, body, attachment_path)
        
        if success:
            # Update timing state
            if self.email_timing.mode == EmailTimingMode.TIME_OF_DAY:
                self.email_timing.mark_daily_email_sent()
            
            # Update timestamp file (for compatibility)
            self._update_email_timestamp()
            
            self.logger.info("📧 Email sent and timing state updated")
        
        return success

def comprehensive_outlook_test(self) -> bool:
    """Run comprehensive Outlook detection test with detailed logging"""
    self.logger.info("🧪 RUNNING COMPREHENSIVE OUTLOOK DETECTION TEST")
    self.logger.info("=" * 60)
    
    # Get debug information
    debug_info = self.debug_outlook_detection()
    
    # Log results
    self.logger.info(f"📊 DETECTION RESULTS:")
    self.logger.info(f"   Methods tested: {len(debug_info['search_methods_tested'])}")
    self.logger.info(f"   Installations found: {len(debug_info['installations_found'])}")
    self.logger.info(f"   Errors encountered: {len(debug_info['errors_encountered'])}")
    
    # Show found installations
    if debug_info['installations_found']:
        self.logger.info(f"\n✅ FOUND INSTALLATIONS:")
        for i, installation in enumerate(debug_info['installations_found'], 1):
            self.logger.info(f"   {i}. Method: {installation['method']}")
            self.logger.info(f"      Path: {installation['result'].get('executable', 'Unknown')}")
            self.logger.info(f"      Type: {installation['result'].get('type', 'Unknown')}")
    
    # Show path availability
    self.logger.info(f"\n📂 SEARCH PATHS:")
    for path, exists in debug_info['search_paths_exist'].items():
        status = "✅" if exists else "❌"
        self.logger.info(f"   {status} {path}")
    
    # Show errors
    if debug_info['errors_encountered']:
        self.logger.info(f"\n💥 ERRORS:")
        for error in debug_info['errors_encountered']:
            self.logger.info(f"   • {error}")
    
    # Test current detection
    current_store_info = self._find_store_outlook_installation()
    current_desktop_available = self._check_desktop_outlook()
    
    self.logger.info(f"\n🎯 CURRENT STATUS:")
    self.logger.info(f"   Store Outlook: {'✅ FOUND' if current_store_info else '❌ NOT FOUND'}")
    self.logger.info(f"   Desktop Outlook: {'✅ AVAILABLE' if current_desktop_available else '❌ NOT AVAILABLE'}")
    
    if current_store_info:
        self.logger.info(f"   Store Path: {current_store_info.get('executable', 'Unknown')}")
        self.logger.info(f"   Detection Method: {current_store_info.get('detection_method', 'Unknown')}")
    
    success = current_store_info is not None or current_desktop_available
    
    self.logger.info(f"\n🏁 OVERALL: {'✅ SUCCESS' if success else '❌ FAILED'}")
    self.logger.info("=" * 60)
    
    return success

class ProfessionalReportGenerator:
    def __init__(self, config, target_productivity=70):
        self.config = config
        self.target_productivity = target_productivity
        self.logger = logging.getLogger(__name__)
        self.session_tracker = SessionTracker()
        self.system_info_collector = SystemInfoCollector()

        self.video_impact_levels = {
            'high': ['youtube', 'netflix', 'hulu', 'disney', 'twitch', 'tiktok', 'facebook', 'instagram'],
            'medium': ['amazon prime', 'paramount', 'peacock', 'hbo max', 'apple tv', 'vimeo'],
            'low': ['spotify', 'soundcloud', 'pandora', 'apple music', 'amazon music']
        }

    def generate_daily_report(self, productivity_data: ProductivityData) -> str:
        """Generate the new organized daily report with uncategorized websites (FIXED: excludes uncategorized from productivity score)"""
        
        sections = []  # Initialize sections immediately at the start
        
        try:
            # Calculate times
            total_productive_time = productivity_data.productive_time
            total_unproductive_time = productivity_data.unproductive_time + productivity_data.background_video_time
            total_uncategorized_time = sum(productivity_data.uncategorized_apps.values()) if productivity_data.uncategorized_apps else 0
            
            # FIXED: Exclude uncategorized time from productivity score calculation
            total_time_for_score = total_productive_time + total_unproductive_time
            total_time = total_productive_time + total_unproductive_time + total_uncategorized_time  # For display purposes

            if total_time == 0:
                return self._generate_organized_no_data_report(productivity_data.date)

            # FIXED: Calculate productivity score only based on productive vs unproductive time
            if total_time_for_score == 0:
                # Special case: if no categorized activity, show N/A instead of 100%
                raw_productivity_score = None
                score_display = "N/A (no categorized activity)"
                score_level = "INSUFFICIENT DATA"
                score_emoji = "📊"
            else:
                raw_productivity_score = round((total_productive_time / total_time_for_score) * 100)
                if raw_productivity_score >= 80:
                    score_level = "EXCELLENT"
                    score_emoji = "🟢"
                elif raw_productivity_score >= 70:
                    score_level = "GOOD"  
                    score_emoji = "🟡"
                elif raw_productivity_score >= 60:
                    score_level = "FAIR"
                    score_emoji = "🟠"
                else:
                    score_level = "NEEDS IMPROVEMENT"
                    score_emoji = "🔴"
                score_display = f"{raw_productivity_score}% ({score_level})"

            # 1. HEADER WITH SYSTEM INFO
            sections.append(self._generate_professional_header_with_info_fixed(
                productivity_data.date, score_display, score_emoji, productivity_data.system_info
            ))
            
            # 2. EXECUTIVE DASHBOARD (updated to show correct calculation)
            sections.append(self._generate_executive_dashboard_fixed(
                productivity_data, raw_productivity_score, total_productive_time, 
                total_unproductive_time, total_uncategorized_time, total_time, total_time_for_score
            ))
            
            # 3. SESSION ANALYSIS
            sections.append(self._generate_session_analysis(productivity_data.date))
            
            # 4. PRODUCTIVITY BREAKDOWN
            sections.append(self._generate_productivity_breakdown(
                productivity_data.productive_apps, productivity_data.unproductive_apps
            ))

            # 5. UNCATEGORIZED WEBSITES (only if there's meaningful data)
            if productivity_data.uncategorized_apps and total_uncategorized_time > 5:  # Only show if > 5 seconds
                sections.append(self._generate_uncategorized_websites_section_cleaned(
                    productivity_data.uncategorized_apps
                ))
            
            # 6. BACKGROUND ACTIVITY
            if productivity_data.background_video_time > 0 or productivity_data.background_videos:
                sections.append(self._generate_background_activity_analysis(
                    productivity_data.background_videos,
                    productivity_data.background_video_apps,
                    productivity_data.verified_playing_apps
                ))
            
            # 7. DETAILED APPENDIX (FIXED: completely removes uncategorized section)
            sections.append(self._generate_detailed_appendix_cleaned(
                productivity_data.productive_apps, productivity_data.unproductive_apps,
                productivity_data.background_video_apps
            ))
            
            report_content = "\n".join(sections)
            
            # Final cleanup: Remove ALL .exe extensions from the entire report
            report_content = AppNameCleaner.clean_all_exe_from_text(report_content)
            
            return report_content
            
        except Exception as e:
            self.logger.error(f"Error generating daily report: {e}")
            self.logger.error(f"Full traceback: {traceback.format_exc()}")
            
            # Return a basic error report instead of crashing
            return f"""
================================================================================
                    DAILY PRODUCTIVITY REPORT - ERROR
================================================================================

Date: {productivity_data.date}
Status: Report Generation Failed

ERROR: {str(e)}

A basic fallback report could not be generated due to the error above.
Please check the logs for more details.

================================================================================
"""

    def _generate_uncategorized_websites_section_cleaned(self, uncategorized_apps: Dict[str, int]) -> str:
        """CLEANED: Generate section for uncategorized websites with better title extraction"""
        
        if not uncategorized_apps:
            return ""
        
        section = f"""

    🌐 UNCATEGORIZED WEBSITES
    {'─' * 50}

    📊 UNCATEGORIZED WEBSITES VISITED:
    Total Time:              {self._format_duration(sum(uncategorized_apps.values()))}
    Unique Sites:            {len(uncategorized_apps)}

    🔍 DETAILED BREAKDOWN:"""

        # Clean and aggregate websites with better title extraction
        cleaned_sites = {}
        for app_title, duration in uncategorized_apps.items():
            # Extract clean website name from messy browser titles
            clean_name = self._extract_clean_website_name(app_title)
            if clean_name in cleaned_sites:
                cleaned_sites[clean_name] += duration
            else:
                cleaned_sites[clean_name] = duration
        
        sorted_sites = sorted(cleaned_sites.items(), key=lambda x: x[1], reverse=True)
        
        for i, (site_name, duration) in enumerate(sorted_sites, 1):
            duration_str = self._format_duration(duration)
            percentage = round((duration / sum(uncategorized_apps.values())) * 100) if uncategorized_apps else 0
            section += f"\n   {i}. {site_name:<35} {duration_str:>8} ({percentage}%)"
        
        return section

    def _generate_executive_dashboard_fixed(self, productivity_data: ProductivityData,
                                       productivity_score: int, total_productive_time: int,
                                       total_unproductive_time: int, total_uncategorized_time: int, 
                                       total_time: int, total_time_for_score: int) -> str:
        """FIXED: Generate executive dashboard with proper score handling"""
        
        # Get session information
        all_events = self._get_all_login_logout_events(productivity_data.date)
        tracker = ChainedSessionTracker()
        sessions = tracker.parse_login_logout_events(all_events)
        session_summary = tracker.get_session_summary(sessions)
        
        # Calculate percentages for display (relative to total time including uncategorized)
        productive_pct = round((total_productive_time / total_time) * 100) if total_time > 0 else 0
        unproductive_pct = round((productivity_data.unproductive_time / total_time) * 100) if total_time > 0 else 0
        background_pct = round((productivity_data.background_video_time / total_time) * 100) if total_time > 0 else 0
        uncategorized_pct = round((total_uncategorized_time / total_time) * 100) if total_time > 0 else 0
        
        dashboard = f"""

    🎯 EXECUTIVE DASHBOARD
    {'─' * 50}

    📊 PRODUCTIVITY METRICS:"""

        if productivity_score is None:
            dashboard += f"\n   Overall Score:           N/A (insufficient categorized activity)"
        else:
            # Calculate percentages for score explanation (relative to categorized time only)
            score_productive_pct = round((total_productive_time / total_time_for_score) * 100) if total_time_for_score > 0 else 0
            score_unproductive_pct = round((total_unproductive_time / total_time_for_score) * 100) if total_time_for_score > 0 else 0
            dashboard += f"\n   Overall Score:           {productivity_score}% (based on {self._format_duration(total_time_for_score)} categorized time)"
            dashboard += f"\n   Score Breakdown:         {score_productive_pct}% productive, {score_unproductive_pct}% unproductive"

        dashboard += f"""

    ⏱️  TIME ALLOCATION:
    Total Active Time:       {self._format_duration(total_time)}
    
    📈 Productive Work:      {self._format_duration(total_productive_time)} ({productive_pct}%)
    📉 Unproductive:         {self._format_duration(productivity_data.unproductive_time)} ({unproductive_pct}%)"""

        # Only show uncategorized if it's significant (> 10 seconds)
        if total_uncategorized_time > 10:
            dashboard += f"\n   🌐 Uncategorized:       {self._format_duration(total_uncategorized_time)} ({uncategorized_pct}%) *excluded from score"
        
        dashboard += f"\n   📺 Background Video:     {self._format_duration(productivity_data.background_video_time)} ({background_pct}%)"

        if AUDIO_DETECTION_AVAILABLE and productivity_data.verified_playing_time > 0:
            verified_pct = round((productivity_data.verified_playing_time / total_time) * 100)
            dashboard += f"\n   🔊 Verified Playing:     {self._format_duration(productivity_data.verified_playing_time)} ({verified_pct}%)"

        dashboard += f"""

    🖥️  SESSION SUMMARY:
    Login Sessions:          {session_summary['total_sessions']} ({session_summary['completed_sessions']} completed)
    Total Login Time:        {self._format_duration(session_summary['total_logged_time'])}
    Currently Active:        {"Yes" if session_summary['ongoing_sessions'] > 0 else "No"}"""

        # Only show the note if there are uncategorized items
        if total_uncategorized_time > 10:
            dashboard += f"\n\n💡 NOTE: Productivity score excludes uncategorized websites to focus on clearly productive vs unproductive activities."

        return dashboard

    def _generate_detailed_appendix_cleaned(self, productive_apps: Dict[str, int],
                                       unproductive_apps: Dict[str, int],
                                       background_video_apps: Dict[str, int]) -> str:
        """COMPLETELY CLEANED: Generate detailed appendix without any uncategorized section"""
        
        section = f"""

    📋 DETAILED ACTIVITY LOG
    {'─' * 50}"""

        # All Productive Apps
        if productive_apps:
            section += "\n\n✅ ALL PRODUCTIVE ACTIVITIES:"
            aggregated_productive = self._aggregate_website_data(productive_apps)
            sorted_productive = sorted(aggregated_productive.items(), 
                                    key=lambda x: x[1]['total'], reverse=True)
            
            for app_name, data in sorted_productive:
                duration = self._format_duration(data['total'])
                section += f"\n   • {app_name:<40} {duration:>10}"
        
        # All Unproductive Apps
        if unproductive_apps:
            section += "\n\n⚠️  ALL TIME DRAINS:"
            aggregated_unproductive = self._aggregate_website_data(unproductive_apps)
            sorted_unproductive = sorted(aggregated_unproductive.items(), 
                                    key=lambda x: x[1]['total'], reverse=True)
            
            for app_name, data in sorted_unproductive:
                duration = self._format_duration(data['total'])
                warning = " ⚠️" if data['total'] >= 600 else ""
                section += f"\n   • {app_name:<40} {duration:>10}{warning}"
        
        # All Background Video
        if background_video_apps:
            section += "\n\n📺 ALL BACKGROUND VIDEO:"
            sorted_background = sorted(background_video_apps.items(), key=lambda x: x[1], reverse=True)
            
            for site, duration_seconds in sorted_background:
                duration = self._format_duration(duration_seconds)
                section += f"\n   • {site:<40} {duration:>10}"
        
        section += f"\n\n{'─' * 50}"
        section += f"\nReport generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        section += f"\n{'=' * 80}"
        
        return section

    def _generate_professional_header_with_info_fixed(self, date: str, score_display: str, 
                                                 score_emoji: str, system_info: Dict[str, Any]) -> str:
        """Generate professional header with fixed score display"""
        formatted_date = datetime.datetime.strptime(date, '%Y-%m-%d').strftime('%A, %B %d, %Y')
        
        # Format system information
        if system_info:
            username = system_info.get('username', 'Unknown User')
            computer_name = system_info.get('computer_name', 'Unknown Computer')
            external_ip = system_info.get('external_ip', 'Unknown')
            local_ip = system_info.get('local_ip', 'Unknown')
            location_info = system_info.get('location', {})
            
            # Format location
            location_str = self.system_info_collector.format_location_string(location_info)
            
            # Format timezone
            timezone = location_info.get('timezone', 'Unknown') if location_info else 'Unknown'
            
            system_section = f"""
    👤 USER INFORMATION:
    User:                    {username}
    Computer:                {computer_name}
    
    🌐 NETWORK & LOCATION:
    External IP:             {external_ip}
    Local IP:                {local_ip}
    Location:                {location_str}
    Timezone:                {timezone}"""
            
            if location_info and location_info.get('isp') != 'Unknown':
                system_section += f"\n   ISP:                     {location_info['isp']}"
        else:
            system_section = """
    👤 USER INFORMATION:
    User:                    Information unavailable
    
    🌐 NETWORK & LOCATION:
    IP Address:              Information unavailable
    Location:                Information unavailable"""
        
        return f"""
    {'=' * 80}
                        DAILY PRODUCTIVITY REPORT
    {'=' * 80}

    Date: {formatted_date}
    Productivity Score: {score_display} {score_emoji}
    {system_section}

    {'=' * 80}"""

    def _generate_session_analysis(self, report_date: str) -> str:
        """IMPROVED: Session analysis with chained sessions - replaces the original method"""
        
        all_events = self._get_all_login_logout_events(report_date)
        
        if not all_events:
            return f"""

    📋 SESSION ANALYSIS
    {'─' * 50}
    No login/logout events recorded today."""

        # Use chained session tracker for better session handling
        tracker = ChainedSessionTracker()
        sessions = tracker.parse_login_logout_events(all_events)
        
        # Group sessions into logical chains
        chains = self._group_sessions_into_chains(sessions)
        
        section = f"""

    📋 SESSION ANALYSIS
    {'─' * 50}

    📅 Today's Work Sessions:"""

        for i, chain in enumerate(chains, 1):
            start_time = chain['start_time'].strftime('%H:%M:%S')
            
            # Calculate total duration and determine if ongoing
            total_duration = 0
            is_ongoing = False
            
            for session in chain['sessions']:
                if session.duration_seconds:
                    total_duration += session.duration_seconds
                elif session.session_type and "Ongoing" in session.session_type:
                    # Calculate ongoing time
                    ongoing_time = (datetime.datetime.now() - session.login_time).total_seconds()
                    total_duration += ongoing_time
                    is_ongoing = True
            
            if is_ongoing:
                end_time = "ACTIVE"
            else:
                end_time = chain['end_time'].strftime('%H:%M:%S') if chain['end_time'] else "Unknown"
            
            duration_str = self._format_duration(int(total_duration))
            
            if len(chain['sessions']) == 1:
                # Single session
                if is_ongoing:
                    section += f"\n   {i}. {start_time} → {end_time} ({duration_str}) 🔄 Ongoing"
                else:
                    section += f"\n   {i}. {start_time} → {end_time} ({duration_str}) ✅ Complete"
            else:
                # Session chain - Extended work period
                parts = len(chain['sessions'])
                
                if is_ongoing:
                    section += f"\n   {i}. {start_time} → {end_time} ({duration_str}) 🔗 Extended Work ({parts} parts, ongoing)"
                else:
                    section += f"\n   {i}. {start_time} → {end_time} ({duration_str}) 🔗 Extended Work ({parts} parts)"
                
                # Show individual parts
                for j, part in enumerate(chain['sessions'], 1):
                    part_start = part.login_time.strftime('%H:%M:%S')
                    
                    if part.logout_time:
                        part_end = part.logout_time.strftime('%H:%M:%S')
                        part_duration = self._format_duration(part.duration_seconds or 0)
                    else:
                        part_end = "ACTIVE"
                        ongoing_duration = int((datetime.datetime.now() - part.login_time).total_seconds())
                        part_duration = self._format_duration(ongoing_duration)
                    
                    section += f"\n      └─ Part {j}: {part_start} → {part_end} ({part_duration})"
        
        # Add summary and explanations
        total_time = sum(chain['total_duration'] for chain in chains)
        extended_sessions = [c for c in chains if len(c['sessions']) > 1]
        
        section += f"""

    📊 SESSION SUMMARY:
    Total Work Time:         {self._format_duration(int(total_time))}
    Work Sessions:           {len(chains)}"""
        
        if extended_sessions:
            section += f"\n   Extended Sessions:       {len(extended_sessions)}"
            section += f"""

    💡 EXTENDED WORK NOTES:
    Sessions longer than 4 hours are split into parts for analysis.
    ✅ All work time is preserved across session parts."""
        
        return section

    def _generate_productivity_breakdown(self, productive_apps: Dict[str, int], 
                                       unproductive_apps: Dict[str, int]) -> str:
        """Generate detailed productivity breakdown"""
        
        section = f"""

💼 PRODUCTIVITY BREAKDOWN
{'─' * 50}"""

        # Top Productive Activities
        if productive_apps:
            section += "\n\n🏆 TOP PRODUCTIVE ACTIVITIES:"
            aggregated_productive = self._aggregate_website_data(productive_apps)
            sorted_productive = sorted(aggregated_productive.items(), 
                                     key=lambda x: x[1]['total'], reverse=True)[:5]
            
            for i, (app_name, data) in enumerate(sorted_productive, 1):
                duration = self._format_duration(data['total'])
                percentage = round((data['total'] / sum(productive_apps.values())) * 100) if productive_apps else 0
                section += f"\n   {i}. {app_name:<30} {duration:>10} ({percentage}%)"
        else:
            section += "\n\n🏆 TOP PRODUCTIVE ACTIVITIES:\n   No productive activities recorded"

        # Top Time Drains
        if unproductive_apps:
            section += "\n\n⚠️  TOP TIME DRAINS:"
            aggregated_unproductive = self._aggregate_website_data(unproductive_apps)
            sorted_unproductive = sorted(aggregated_unproductive.items(), 
                                       key=lambda x: x[1]['total'], reverse=True)[:5]
            
            for i, (app_name, data) in enumerate(sorted_unproductive, 1):
                duration = self._format_duration(data['total'])
                percentage = round((data['total'] / sum(unproductive_apps.values())) * 100) if unproductive_apps else 0
                warning = " 🚨" if data['total'] >= 600 else ""  # 10+ minutes warning
                section += f"\n   {i}. {app_name:<30} {duration:>10} ({percentage}%){warning}"
        else:
            section += "\n\n⚠️  TOP TIME DRAINS:\n   ✅ Excellent focus - no significant time drains detected!"

        return section

    def _generate_background_activity_analysis(self, background_videos: List[BackgroundVideoSession],
                                             background_video_apps: Dict[str, int],
                                             verified_playing_apps: Dict[str, int]) -> str:
        """Generate background activity analysis"""
        
        section = f"""

📺 BACKGROUND ACTIVITY ANALYSIS
{'─' * 50}"""

        if background_video_apps:
            total_background = sum(background_video_apps.values())
            total_verified = sum(verified_playing_apps.values()) if verified_playing_apps else 0
            
            section += f"\n\n📊 BACKGROUND VIDEO SUMMARY:"
            section += f"\n   Total Background Time:   {self._format_duration(total_background)}"
            
            if AUDIO_DETECTION_AVAILABLE and total_verified > 0:
                section += f"\n   Verified Playing Time:   {self._format_duration(total_verified)}"
                paused_time = total_background - total_verified
                if paused_time > 0:
                    section += f"\n   Paused/Silent Time:      {self._format_duration(paused_time)}"
            
            section += f"\n\n🎯 BACKGROUND ACTIVITY BREAKDOWN:"
            
            # Sort by total time
            sorted_background = sorted(background_video_apps.items(), key=lambda x: x[1], reverse=True)
            
            for i, (site, total_time) in enumerate(sorted_background, 1):
                playing_time = verified_playing_apps.get(site, 0) if verified_playing_apps else 0
                
                impact_level = self._classify_video_impact(site)
                impact_emoji = {"high": "🔴", "medium": "🟡", "low": "🟢"}.get(impact_level, "⚪")
                
                if AUDIO_DETECTION_AVAILABLE and playing_time > 0:
                    section += f"\n   {i}. {site:<25} {self._format_duration(total_time):>8} total, {self._format_duration(playing_time):>8} playing {impact_emoji}"
                else:
                    section += f"\n   {i}. {site:<25} {self._format_duration(total_time):>8} {impact_emoji}"
        
        return section

    def _generate_organized_no_data_report(self, date: str) -> str:
        """Generate organized no-data report with system info"""
        formatted_date = datetime.datetime.strptime(date, '%Y-%m-%d').strftime('%A, %B %d, %Y')
        
        # Try to get system info even for no-data report
        try:
            if hasattr(self, 'monitor') and hasattr(self.monitor, '_collect_system_info'):
                system_info = self.monitor._collect_system_info()
            else:
                system_info = self.system_info_collector.get_system_info()
        except:
            system_info = None
        
        # Format system information
        if system_info:
            username = system_info.get('username', 'Unknown User')
            computer_name = system_info.get('computer_name', 'Unknown Computer')
            external_ip = system_info.get('external_ip', 'Unknown')
            location_info = system_info.get('location', {})
            
            # Format location
            if location_info:
                location_str = self.system_info_collector.format_location_string(location_info)
            else:
                location_str = "Unknown"
            
            system_section = f"""
👤 USER INFORMATION:
   User:                    {username}
   Computer:                {computer_name}
   
🌐 NETWORK & LOCATION:
   External IP:             {external_ip}
   Location:                {location_str}"""
        else:
            system_section = """
👤 USER INFORMATION:
   User:                    Information unavailable
   
🌐 NETWORK & LOCATION:
   Location:                Information unavailable"""
        
        report_content = f"""
{'=' * 80}
                     DAILY PRODUCTIVITY REPORT
{'=' * 80}

Date: {formatted_date}
Productivity Score: No Data Available 📊
{system_section}

{'=' * 80}

🎯 EXECUTIVE DASHBOARD
{'─' * 50}

📊 PRODUCTIVITY METRICS:
   Overall Score:           No data recorded

⏱️  TIME ALLOCATION:
   Total Active Time:       0s
   Productive Work:         0s (0%)
   Unproductive:           0s (0%)
   Uncategorized:         0s (0%)
   Background Video:       0s (0%)

🖥️  SESSION SUMMARY:
   Login Sessions:          0 (0 completed)
   Total Login Time:        0s
   Currently Active:        Unknown

📋 SESSION ANALYSIS
{'─' * 50}
No login/logout events recorded today.

💼 PRODUCTIVITY BREAKDOWN
{'─' * 50}

🏆 TOP PRODUCTIVE ACTIVITIES:
   No productive activities recorded

⚠️  TOP TIME DRAINS:
   No activities recorded

🌐 UNCATEGORIZED WEBSITES
{'─' * 50}

📊 UNCATEGORIZED WEBSITES VISITED:
   No websites recorded

📋 DETAILED ACTIVITY LOG
{'─' * 50}

✅ ALL PRODUCTIVE ACTIVITIES:
   No activities recorded

⚠️  ALL TIME DRAINS:
   No activities recorded

🌐 ALL UNCATEGORIZED WEBSITES:
   No activities recorded

📺 ALL BACKGROUND VIDEO:
   No activities recorded

{'─' * 50}
Report generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'=' * 80}"""

        return AppNameCleaner.clean_all_exe_from_text(report_content)

        # STEP 3: Add this helper method to your ProfessionalReportGenerator class
    def _group_sessions_into_chains(self, sessions: List[LoginSession]) -> List[dict]:
        """Group related sessions into chains for display"""
        chains = []
        current_chain = []
        
        for session in sessions:
            if "Part" in session.session_type:
                # This is part of a chain
                current_chain.append(session)
            else:
                # Close current chain if exists
                if current_chain:
                    total_duration = sum(s.duration_seconds or 0 for s in current_chain)
                    # Add ongoing time for last session if applicable
                    if current_chain[-1].session_type and "Ongoing" in current_chain[-1].session_type:
                        ongoing_time = int((datetime.datetime.now() - current_chain[-1].login_time).total_seconds())
                        total_duration += ongoing_time
                    
                    chains.append({
                        'sessions': current_chain,
                        'total_duration': total_duration,
                        'start_time': current_chain[0].login_time,
                        'end_time': current_chain[-1].logout_time,
                    })
                    current_chain = []
                
                # Add single session
                duration = session.duration_seconds or 0
                if session.session_type == "Ongoing":
                    duration = int((datetime.datetime.now() - session.login_time).total_seconds())
                
                chains.append({
                    'sessions': [session],
                    'total_duration': duration,
                    'start_time': session.login_time,
                    'end_time': session.logout_time,
                })
        
        # Handle final chain
        if current_chain:
            total_duration = sum(s.duration_seconds or 0 for s in current_chain)
            if current_chain[-1].session_type and "Ongoing" in current_chain[-1].session_type:
                ongoing_time = int((datetime.datetime.now() - current_chain[-1].login_time).total_seconds())
                total_duration += ongoing_time
            
            chains.append({
                'sessions': current_chain,
                'total_duration': total_duration,
                'start_time': current_chain[0].login_time,
                'end_time': current_chain[-1].logout_time,
            })
        
        return chains
    
    def _get_all_login_logout_events(self, report_date: str) -> List[str]:
        """Get all login/logout events from multiple sources"""
        all_events = []
        
        # Method 1: Get events from activity logger memory (current session)
        if hasattr(self, 'activity_logger'):
            recent_events = self.activity_logger.get_recent_login_logout_events()
            all_events.extend(recent_events)
        
        # Method 2: Read from activity log file
        try:
            if os.path.exists(self.config.ACTIVITY_LOG):
                with open(self.config.ACTIVITY_LOG, "r", encoding="utf-8") as f:
                    content = f.read()
                    for line in content.split('\n'):
                        if any(keyword in line.lower() for keyword in [
                            "user logged in", "user logged out", 
                            "system startup detected", "system shutdown detected"
                        ]):
                            all_events.append(line.strip())
        except Exception as e:
            self.logger.error(f"Error reading login/logout events from file: {e}")

        # Remove duplicates while preserving order
        seen = set()
        unique_events = []
        for event in all_events:
            if event not in seen:
                seen.add(event)
                unique_events.append(event)

        return unique_events

    def _classify_video_impact(self, site: str) -> str:
        site_lower = site.lower()

        for impact_level, sites in self.video_impact_levels.items():
            if any(video_site in site_lower for video_site in sites):
                return impact_level

        return 'medium'

    def _aggregate_website_data(self, apps_data: Dict[str, int]) -> Dict[str, Dict]:
        website_data = {}

        for app_title, seconds in apps_data.items():
            if self._is_browser_entry(app_title):
                website = self.extract_website_from_title(app_title)
                browser = self._extract_browser_from_app_title(app_title)

                if website not in website_data:
                    website_data[website] = {'total': 0, 'browsers': {}}

                website_data[website]['total'] += seconds
                website_data[website]['browsers'][browser] = website_data[website]['browsers'].get(browser, 0) + seconds
            else:
                # Clean the app name before using it
                clean_app_name = AppNameCleaner.clean_app_base_name(app_title)
                
                if clean_app_name not in website_data:
                    website_data[clean_app_name] = {'total': 0, 'browsers': {}}

                website_data[clean_app_name]['total'] += seconds
                website_data[clean_app_name]['browsers']['app'] = website_data[clean_app_name]['browsers'].get('app', 0) + seconds

        return website_data

    def _is_browser_entry(self, app_title: str) -> bool:
        browser_indicators = [
            'chrome.exe', 'firefox.exe', 'msedge.exe', 'safari.exe',
            'mozilla firefox', 'google chrome', 'microsoft edge'
        ]

        title_lower = app_title.lower()
        return any(indicator in title_lower for indicator in browser_indicators)

    def _extract_browser_from_app_title(self, app_title: str) -> str:
        title_lower = app_title.lower()

        if 'chrome' in title_lower:
            return 'Chrome'
        elif 'firefox' in title_lower:
            return 'Firefox'
        elif 'edge' in title_lower:
            return 'Edge'
        elif 'safari' in title_lower:
            return 'Safari'
        else:
            return 'Browser'

    def _format_duration(self, seconds: int) -> str:
        if seconds < 60:
            return f"{seconds}s"
        elif seconds < 3600:
            minutes = seconds // 60
            return f"{minutes}m"
        else:
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            if minutes > 0:
                return f"{hours}h {minutes}m"
            else:
                return f"{hours}h"

    def save_report_to_file(self, report_content: str, date: str) -> str:
        report_filename = f"productivity_report_{date}.txt"
        report_path = os.path.join(self.config.LOG_DIR, report_filename)

        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
            return report_path
        except IOError as e:
            print(f"Error saving report: {e}")
            return None

    def extract_website_from_title(self, title: str) -> str:
        """ENHANCED: Better Amazon and shopping site recognition"""
        title_lower = title.lower()
        
        # SPECIFIC: Amazon detection (all Amazon URLs)
        if 'amazon.com' in title_lower:
            return 'Amazon Shopping'
        
        # Enhanced website patterns
        website_patterns = {
            # Shopping
            'ebay': ['ebay.com', 'ebay '],
            'etsy': ['etsy.com', 'etsy '],
            'walmart': ['walmart.com', 'walmart '],
            'target': ['target.com', 'target '],
            'best_buy': ['best buy', 'bestbuy.com'],
            
            # Social Media
            'youtube': ['youtube', 'youtu.be'],
            'facebook': ['facebook', 'fb.com'],
            'instagram': ['instagram'],
            'twitter': ['twitter', 'x.com'],
            'linkedin': ['linkedin'],
            'reddit': ['reddit'],
            'tiktok': ['tiktok'],
            
            # Entertainment
            'netflix': ['netflix'],
            'hulu': ['hulu'],
            'disney+': ['disney', 'disneyplus'],
            'twitch': ['twitch'],
            'spotify': ['spotify'],
            
            # Gaming
            'runescape': ['runescape', 'osrs', 'oldschool runescape'],
            'steam': ['steam'],
            
            # Productive
            'google': ['google search', 'google.com'],
            'gmail': ['gmail', 'mail.google'],
            'github': ['github'],
            'stackoverflow': ['stackoverflow', 'stack overflow'],
        }

        for website, patterns in website_patterns.items():
            for pattern in patterns:
                if pattern in title_lower:
                    return self.format_website_name(website)

        title_lower = title.lower()

        for website, patterns in website_patterns.items():
            for pattern in patterns:
                if pattern in title_lower:
                    return self.format_website_name(website)

        if ' - ' in title:
            parts = title.split(' - ')
            if len(parts) >= 2:
                potential_site = parts[-1].strip()
                if self.is_likely_website_name(potential_site):
                    return self.clean_website_name(potential_site)

        if ' | ' in title:
            parts = title.split(' | ')
            if len(parts) >= 2:
                potential_site = parts[-1].strip()
                if self.is_likely_website_name(potential_site):
                    return self.clean_website_name(potential_site)

        browser_suffixes = [
            'Mozilla Firefox', 'Google Chrome', 'Chrome',
            'Microsoft Edge', 'Edge', 'Safari'
        ]

        for suffix in browser_suffixes:
            if title.endswith(suffix):
                clean_title = title.replace(suffix, '').strip()
                if clean_title:
                    website = self.extract_domain_from_clean_title(clean_title)
                    if website:
                        return website

        return self.clean_website_name(title)

    def format_website_name(self, website_key: str) -> str:
        formatting = {
            'youtube': 'YouTube',
            'facebook': 'Facebook',
            'instagram': 'Instagram',
            'twitter': 'Twitter/X',
            'linkedin': 'LinkedIn',
            'reddit': 'Reddit',
            'tiktok': 'TikTok',
            'netflix': 'Netflix',
            'hulu': 'Hulu',
            'disney+': 'Disney+',
            'amazon_prime': 'Amazon Prime',
            'twitch': 'Twitch',
            'spotify': 'Spotify',
            'amazon': 'Amazon',
            'ebay': 'eBay',
            'etsy': 'Etsy',
            'google': 'Google',
            'gmail': 'Gmail',
            'gaming': 'Gaming Platforms'
        }

        return formatting.get(website_key, website_key.title())

    def _extract_clean_website_name(self, app_title: str) -> str:
        """FIXED: Extract clean website name from browser title with better logic"""
        if not app_title:
            return "Unknown Website"
        
        # Remove browser name patterns first
        title = app_title
        
        # Remove browser prefixes and suffixes
        browser_patterns = [
            ("Mozilla Firefox - ", ""),
            (" — Mozilla Firefox", ""),
            (" - Mozilla Firefox", ""),
            ("Google Chrome - ", ""),
            (" - Google Chrome", ""),
            ("Microsoft Edge - ", ""),
            (" - Microsoft Edge", ""),
        ]
        
        for pattern, replacement in browser_patterns:
            title = title.replace(pattern, replacement)
        
        # Now extract the website name intelligently
        title = title.strip()
        
        # Method 1: Look for known website patterns first
        website_patterns = {
            'runescape': 'RuneScape',
            'oldschool runescape': 'Old School RuneScape', 
            'osrs': 'Old School RuneScape',
            'youtube': 'YouTube',
            'facebook': 'Facebook',
            'instagram': 'Instagram',
            'twitter': 'Twitter/X',
            'linkedin': 'LinkedIn',
            'reddit': 'Reddit',
            'netflix': 'Netflix',
            'amazon': 'Amazon',
            'google': 'Google',
            'gmail': 'Gmail',
            'github': 'GitHub',
            'stackoverflow': 'Stack Overflow',
            'microsoft': 'Microsoft',
            'outlook': 'Outlook',
            'teams': 'Microsoft Teams'
        }
        
        title_lower = title.lower()
        for pattern, clean_name in website_patterns.items():
            if pattern in title_lower:
                return clean_name
        
        # Method 2: Handle "Page Title - Website Name" format
        if " - " in title:
            parts = title.split(" - ")
            
            # For RuneScape example: "Play old school runescape - world server list"
            # We want the first part (the action/page) to determine the site
            first_part = parts[0].strip().lower()
            last_part = parts[-1].strip()
            
            # Check if first part contains website indicators
            if any(indicator in first_part for indicator in ['runescape', 'osrs', 'play old school']):
                return 'Old School RuneScape'
            
            # If last part looks like a website name (short and no generic words)
            generic_suffixes = ['world server list', 'home page', 'main page', 'login', 'dashboard', 
                            'settings', 'profile', 'about', 'contact', 'help', 'support']
            
            if (len(last_part) < 30 and 
                not any(suffix in last_part.lower() for suffix in generic_suffixes) and
                ('.' in last_part or last_part.istitle())):
                return last_part
            
            # Otherwise, try to extract from first part
            return self._extract_site_from_page_title(parts[0])
        
        # Method 3: Handle "Page Title | Website Name" format  
        elif " | " in title:
            parts = title.split(" | ")
            last_part = parts[-1].strip()
            
            # Usually the website name is after the |
            if len(last_part) < 30 and last_part.istitle():
                return last_part
                
            return self._extract_site_from_page_title(parts[0])
        
        # Method 4: Try to extract domain-like patterns
        elif "." in title and len(title) < 50:
            # Might be a direct domain
            return self._clean_domain_name(title)
        
        # Method 5: Last resort - use the title but try to make it meaningful
        return self._extract_site_from_page_title(title)

    def _extract_site_from_page_title(self, page_title: str) -> str:
        """Extract likely website name from page title"""
        if not page_title:
            return "Unknown Website"
        
        page_lower = page_title.lower().strip()
        
        # Common patterns that indicate the website
        site_indicators = {
            'play old school runescape': 'Old School RuneScape',
            'runescape': 'RuneScape', 
            'osrs': 'Old School RuneScape',
            'youtube': 'YouTube',
            'watch': 'YouTube',  # "Watch something" usually means YouTube
            'facebook': 'Facebook',
            'instagram': 'Instagram',
            'twitter': 'Twitter/X',
            'reddit': 'Reddit',
            'netflix': 'Netflix',
            'amazon shopping': 'Amazon',
            'gmail': 'Gmail',
            'google search': 'Google',
            'github': 'GitHub',
            'stack overflow': 'Stack Overflow',
            'microsoft teams': 'Microsoft Teams',
            'outlook': 'Outlook'
        }
    
        for indicator, site_name in site_indicators.items():
            if indicator in page_lower:
                return site_name
        
        # If nothing matches, clean up the title
        if len(page_title) > 30:
            return page_title[:30].strip() + "..."
        else:
            return page_title.strip().title()

    def _clean_domain_name(self, domain: str) -> str:
        """Clean a domain name for display"""
        domain = domain.strip().lower()
        
        # Remove common prefixes
        if domain.startswith('www.'):
            domain = domain[4:]
        if domain.startswith('http://'):
            domain = domain[7:]
        if domain.startswith('https://'):
            domain = domain[8:]
        
        # Remove common suffixes
        if domain.endswith('.com'):
            domain = domain[:-4]
        elif domain.endswith('.org'):
            domain = domain[:-4]
        elif domain.endswith('.net'):
            domain = domain[:-4]
        
        # Capitalize properly
        return domain.title()

    def is_likely_website_name(self, text: str) -> bool:
        text_lower = text.lower()

        website_indicators = [
            '.com', '.org', '.net', '.edu', '.gov',
            'youtube', 'google', 'facebook', 'twitter',
            'netflix', 'amazon', 'microsoft', 'apple'
        ]

        return any(indicator in text_lower for indicator in website_indicators) or len(text) < 20

    def clean_website_name(self, name: str) -> str:
        suffixes_to_remove = [
            '.com', '.org', '.net', '.edu', '.gov',
            'www.', 'https://', 'http://'
        ]

        cleaned = name.strip()
        for suffix in suffixes_to_remove:
            cleaned = cleaned.replace(suffix, '')

        if len(cleaned) <= 15:
            return cleaned.title()
        else:
            return cleaned

    def extract_domain_from_clean_title(self, title: str) -> Optional[str]:
        domain_patterns = [
            r'(\w+)\.com',
            r'(\w+)\.org',
            r'(\w+)\.net'
        ]

        title_lower = title.lower()

        for pattern in domain_patterns:
            match = re.search(pattern, title_lower)
            if match:
                return self.format_website_name(match.group(1))

        return None

def cleanup_old_persistence_data():
    """Utility function to clean up old persistence data"""
    try:
        config = CompleteEnhancedConfig()
        persistence = CompleteEnhancedProductivityDataPersistence(config)
        persistence.cleanup_old_data(days_to_keep=7)
        print("✅ Old persistence data cleaned up")
    except Exception as e:
        print(f"❌ Error cleaning up persistence data: {e}")

def test_fixed_wmi_queries():
    """Test the fixed WMI query methods"""
    print("=== TESTING FIXED WMI QUERIES ===")
    
    try:
        poller = ImprovedLoginLogoutPoller(max_init_time=30)
        
        if poller.initialization_success:
            print(f"✓ Initialization successful")
            print(f"✓ Query method: {poller.query_method}")
            print(f"✓ Fallback mode: {poller.fallback_mode}")
            
            # Test polling
            events = poller.poll_events()
            print(f"✓ Poll successful: {len(events)} events found")
            
            status = poller.get_status()
            print(f"✓ Status: {status}")
            
            return True
        else:
            print("✗ Initialization failed")
            return False
            
    except Exception as e:
        print(f"✗ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

class ActivityMonitor:
    def __init__(self):
        self.config = CompleteEnhancedConfig()
        self.activity_logger = CompleteEnhancedActivityLogger(self.config)
        self.session_tracker = ChainedSessionTracker()
        
        # Initialize persistence manager
        self.persistence = CompleteEnhancedProductivityDataPersistence(self.config)
        
        # Use improved login/logout poller
        self.login_logout_poller = None
        self.wmi_initialization_complete = threading.Event()
        self.wmi_initialization_failed = False
        
        # Other components
        self.config_manager = EnhancedConfigManager(self.config.CONFIG_PATH)
        self.email_manager = EnhancedHybridOutlookManager(self.config, self.config_manager)

         # Always log Store Outlook detection results
        try:
            self.activity_logger.debug_log("🔍 Checking for Microsoft Store Outlook...")
            available_methods = self.email_manager._get_available_outlook_methods()
            
            store_found = any('store' in method[0].lower() for method in available_methods)
            self.activity_logger.debug_log(f"Store Outlook Status: {'✅ DETECTED' if store_found else '❌ NOT FOUND'}")
            
            if available_methods:
                self.activity_logger.debug_log(f"Available email methods:")
                for method_name, method_info in available_methods:
                    priority = "HIGH" if 'store' in method_name.lower() else "NORMAL"
                    self.activity_logger.debug_log(f"   • {method_name} ({priority} priority)")
            
        except Exception as e:
            self.activity_logger.debug_log(f"Email detection error: {e}")
        
        self.tracker: Optional[ForegroundTracker] = None
        self.background_video_tracker: Optional[BackgroundVideoTracker] = None
        self.reporter = ActivityReporter(self.config, self.activity_logger)
        self.report_generator = ProfessionalReportGenerator(self.config)
        
        # Pass references for enhanced reporting
        self.report_generator.activity_logger = self.activity_logger
        self.report_generator.monitor = self
        
        self.running = False
        self._setup_shutdown_handlers()


    def _load_real_productivity_data(self, target_date: str) -> Optional[ProductivityData]:
        """Load real productivity data from a specific date"""
        try:
            historical_data = self.persistence.load_historical_data(target_date)
            
            if not historical_data:
                return None
            
            app_data = historical_data.get('app_data', {})
            bg_data = historical_data.get('bg_data', {})
            
            app_times = app_data.get('app_times', {}) if app_data else {}
            background_video_times = bg_data.get('background_video_times', {}) if bg_data else {}
            verified_playing_times = bg_data.get('verified_playing_times', {}) if bg_data else {}
            
            # Categorize the loaded apps
            productive_apps = {}
            unproductive_apps = {}
            uncategorized_apps = {}
            
            for app, secs in app_times.items():
                try:
                    category = AppCategorizer.categorize_app(app)
                    if category == Category.PRODUCTIVE:
                        productive_apps[app] = int(secs)
                    elif category == Category.UNPRODUCTIVE:
                        unproductive_apps[app] = int(secs)
                    elif category == Category.UNCATEGORIZED:
                        uncategorized_apps[app] = int(secs)
                except Exception as e:
                    self.activity_logger.debug_log(f"Error categorizing {app}: {e}")
                    uncategorized_apps[app] = int(secs)
            
            # Calculate totals
            productive_time = sum(productive_apps.values())
            unproductive_time = sum(unproductive_apps.values())
            background_video_time = sum(background_video_times.values()) if background_video_times else 0
            verified_playing_time = sum(verified_playing_times.values()) if verified_playing_times else 0
            
            try:
                system_info = self._collect_system_info()
            except:
                system_info = None
            
            self.activity_logger.debug_log(f"📊 Loaded real data for {target_date}:")
            self.activity_logger.debug_log(f"   Productive: {self._format_duration(productive_time)}")
            self.activity_logger.debug_log(f"   Unproductive: {self._format_duration(unproductive_time)}")
            self.activity_logger.debug_log(f"   Background video: {self._format_duration(background_video_time)}")
            self.activity_logger.debug_log(f"   Total apps: {len(app_times)}")
            
            return ProductivityData(
                productive_time=productive_time,
                unproductive_time=unproductive_time,
                background_video_time=background_video_time,
                verified_playing_time=verified_playing_time,
                productive_apps=productive_apps,
                unproductive_apps=unproductive_apps,
                uncategorized_apps=uncategorized_apps,
                background_videos=[],
                background_video_apps=background_video_times,
                verified_playing_apps=verified_playing_times,
                date=target_date,
                system_info=system_info
            )
            
        except Exception as e:
            self.activity_logger.debug_log(f"Error loading real productivity data for {target_date}: {e}")
            return None

    def _find_existing_report(self, target_date: str) -> Optional[str]:
        """Look for an existing report file that wasn't sent"""
        try:
            report_filename = f"productivity_report_{target_date}.txt"
            report_path = os.path.join(self.config.LOG_DIR, report_filename)
            
            if os.path.exists(report_path):
                self.activity_logger.debug_log(f"📄 Found existing report file: {report_filename}")
                return report_path
            
            return None
            
        except Exception as e:
            self.activity_logger.debug_log(f"Error looking for existing report: {e}")
            return None

    def _send_existing_report(self, missed_date: str, report_path: str) -> bool:
        """Send an existing report file that wasn't emailed"""
        try:
            subject = f"MISSED Daily Productivity Report - {missed_date}"
            body = f"""This is the productivity report for {missed_date} that was generated but not sent due to a system issue.

    The original report file has been recovered and is attached.

    Report generated on: {missed_date}
    Recovered on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"""
            
            success = self.email_manager.send_email_via_outlook(subject, body, report_path)
            
            if success:
                self.activity_logger.debug_log(f"📧 Existing report for {missed_date} sent successfully")
            
            return success
            
        except Exception as e:
            self.activity_logger.debug_log(f"Error sending existing report: {e}")
            return False

    def _create_minimal_report_data(self, date: str) -> ProductivityData:
        """Create minimal report data for missed dates"""
        try:
            system_info = self._collect_system_info()
        except:
            system_info = None
        
        return ProductivityData(
            productive_time=0,
            unproductive_time=0,
            background_video_time=0,
            verified_playing_time=0,
            productive_apps={},
            unproductive_apps={},
            uncategorized_apps={},
            background_videos=[],
            background_video_apps={},
            verified_playing_apps={},
            date=date,
            system_info=system_info
        )

    def _format_duration(self, seconds: int) -> str:
        """Format duration for email body"""
        if seconds < 60:
            return f"{seconds}s"
        elif seconds < 3600:
            minutes = seconds // 60
            return f"{minutes}m"
        else:
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            if minutes > 0:
                return f"{hours}h {minutes}m"
            else:
                return f"{hours}h"

    # UPDATE generate_and_email_daily_report method to mark reports as sent
    def generate_and_email_daily_report(self) -> bool:
        """Generate daily report and email it - ENHANCED with sent tracking"""
        productivity_data = self._collect_productivity_data()

        report_content = self.report_generator.generate_daily_report(productivity_data)
        report_path = self.report_generator.save_report_to_file(report_content, productivity_data.date)
        self.activity_logger.debug_log(f"Daily report generated and saved to: {report_path}")

        success = self.email_manager.send_email_via_outlook(
            "Daily Productivity Report",
            "Attached is the daily productivity report.",
            report_path
        )
        
        # Mark report as sent if successful
        if success:
            self.activity_logger.mark_report_sent(productivity_data.date, "daily")
            self.activity_logger.debug_log(f"✅ Report marked as sent for {productivity_data.date}")
        
        return success

    def _collect_system_info(self) -> Dict[str, Any]:
        """Collect system information for the report"""
        if not hasattr(self, 'system_info_collector'):
            self.system_info_collector = SystemInfoCollector()
        
        return self.system_info_collector.get_system_info()

    def _initialize_wmi_parallel(self):
        """Initialize WMI with improved error handling"""
        def wmi_init_worker():
            try:
                self.activity_logger.debug_log("Starting improved login/logout detection...")
                self.login_logout_poller = ImprovedLoginLogoutPoller(max_init_time=60)
                
                if self.login_logout_poller.initialization_success:
                    status = self.login_logout_poller.get_status()
                    self.activity_logger.debug_log(f"Login/logout detection ready - Status: {status}")
                    self.activity_logger.buffer_login_logout_event("System startup detected - monitoring started")
                else:
                    self.activity_logger.debug_log("Login/logout detection failed to initialize")
                    self.wmi_initialization_failed = True
                
                self.wmi_initialization_complete.set()
                
            except Exception as e:
                self.activity_logger.debug_log(f"Login/logout detection initialization failed: {e}")
                self.activity_logger.debug_log(f"Full error: {traceback.format_exc()}")
                self.wmi_initialization_failed = True
                self.wmi_initialization_complete.set()
        
        wmi_thread = threading.Thread(target=wmi_init_worker, daemon=True)
        wmi_thread.start()
        return wmi_thread

    def _setup_shutdown_handlers(self):
        def handle_exit_event(ctrl_type):
            if ctrl_type in (win32con.CTRL_LOGOFF_EVENT, win32con.CTRL_SHUTDOWN_EVENT):
                self.activity_logger.debug_log("Detected system logout or shutdown event.")
                self._perform_shutdown_tasks()
                return True
            return False

        def graceful_exit(*args):
            self._perform_shutdown_tasks()

        win32api.SetConsoleCtrlHandler(handle_exit_event, True)
        atexit.register(graceful_exit)
        signal.signal(signal.SIGTERM, graceful_exit)

    def _perform_shutdown_tasks(self):
        """Enhanced shutdown with data persistence"""
        self.activity_logger.buffer_login_logout_event("System shutdown detected")
        
        if self.tracker:
            self.tracker.stop()
            self.reporter.log_activity(self.tracker)
            self.activity_logger.flush_buffer()

        if self.background_video_tracker:
            self.background_video_tracker.stop()

        # Save data before generating final report
        self.persistence.save_tracking_data(self.tracker, self.background_video_tracker)

        self.generate_and_email_daily_report()
        self.activity_logger.buffer_login_logout_event("System logout or shutdown detected. Final productivity report emailed.")
        self.activity_logger.flush_buffer()

        self.activity_logger.debug_log("Graceful shutdown completed - activity logged and data saved.")

    def run(self):
        
        self.activity_logger.debug_log("Activity monitor started with complete enhanced logging and missed report recovery.")
        # self.check_and_send_missed_reports()
        self.activity_logger.debug_log("Activity monitor started with data persistence.")

        # Start core tracking immediately
        self.tracker = ForegroundTracker(self.config, self.activity_logger)
        self.background_video_tracker = BackgroundVideoTracker(self.config, self.activity_logger)
        
        self.tracker.start()
        self.background_video_tracker.start()
        self.activity_logger.debug_log("Core tracking started immediately.")

        # Load previous session data if available
        self.persistence.load_tracking_data(self.tracker, self.background_video_tracker)
        
        # Verify loaded data
        verification = self.persistence.verify_loaded_data(self.tracker, self.background_video_tracker)
        self.activity_logger.debug_log(f"📊 Data verification: {verification}")

        # Start improved WMI initialization in parallel
        wmi_thread = self._initialize_wmi_parallel()

        self.running = True

        try:
            last_email_time = self.email_manager.get_last_email_time()
            is_first_run = not os.path.exists(self.config.EMAIL_TRACK_FILE)

            # IMPROVED STARTUP LOGIC - Respect timing settings
            if is_first_run:
                # First time running - just set timestamp, don't send
                self.email_manager._update_email_timestamp()
                self.activity_logger.debug_log("🚀 First run detected - email timing initialized, no immediate report sent")
            else:
                # Check if we should send based on timing rules
                should_send, reason = self.email_manager.should_send_productivity_report(last_email_time)
                if should_send:
                    self.activity_logger.debug_log(f"📧 Startup report check: {reason}")
                    
                    # Only send if we have meaningful data or it's been a very long time
                    startup_data = self._collect_productivity_data()
                    total_startup_time = (startup_data.productive_time + 
                                        startup_data.unproductive_time + 
                                        startup_data.background_video_time)
                    
                    if total_startup_time > 300:  # Only send if >5 minutes of activity
                        if self.generate_and_email_daily_report():
                            last_email_time = time.time()
                            self.activity_logger.debug_log("📧 Startup report sent - meaningful activity detected")
                    else:
                        self.activity_logger.debug_log(f"📧 Startup report skipped - insufficient activity ({total_startup_time}s)")
                        # Update timestamp to prevent immediate sending in main loop
                        self.email_manager._update_email_timestamp()
                else:
                    self.activity_logger.debug_log(f"📧 No startup report needed: {reason}")

            loop_count = 0
            wmi_ready_logged = False
            last_save_time = time.time()
            save_interval = 60  # Save every 1 minute

            while self.running:
                try:
                    loop_count += 1
                    
                    # Wait for WMI on first few iterations
                    if loop_count <= 3 and not self.wmi_initialization_complete.is_set():
                        self.activity_logger.debug_log(f"Loop #{loop_count}: Waiting for WMI initialization...")
                        self.wmi_initialization_complete.wait(self.config.LOG_INTERVAL)
                    else:
                        time.sleep(self.config.LOG_INTERVAL)

                    # Log WMI status once
                    if not wmi_ready_logged and self.wmi_initialization_complete.is_set():
                        if self.wmi_initialization_failed:
                            self.activity_logger.debug_log("WMI initialization failed - login/logout tracking limited")
                        else:
                            self.activity_logger.debug_log("WMI ready - login/logout tracking enabled")
                        wmi_ready_logged = True

                    # Core tracking (always runs)
                    self.reporter.log_activity(self.tracker)
                    self.activity_logger.flush_buffer()

                    # Periodically save tracking data to disk
                    current_time = time.time()
                    if current_time - last_save_time >= save_interval:
                        self.persistence.save_tracking_data(self.tracker, self.background_video_tracker)
                        last_save_time = current_time

                   # Periodically verify data integrity AND cleanup old backups (every hour)
                    if loop_count % 60 == 0:  # Every hour
                        verification = self.persistence.verify_loaded_data(self.tracker, self.background_video_tracker)
                        self.activity_logger.debug_log(f"🔍 Hourly verification: {verification}")
                        
                        # Clean up old dated backups every 24 hours
                        if loop_count % 1440 == 0:  # Every 24 hours
                            self.persistence.cleanup_old_dated_backups()

                    # Reports
                    # Reports
                    self.generate_daily_report()

                    # Check if we should send email report
                    should_send, reason = self.email_manager.should_send_productivity_report(last_email_time)
                    if should_send: 
                        self.activity_logger.debug_log(f"📧 Sending report: {reason}")
                        if self.generate_and_email_daily_report():
                            last_email_time = time.time()
                    else:
                        # Optionally log why we're not sending (every 10th iteration to avoid spam)
                        if loop_count % 10 == 0:
                            self.activity_logger.debug_log(f"📧 Not sending report: {reason}")

                    # Login/logout polling with improved error handling
                    if (self.wmi_initialization_complete.is_set() and 
                        not self.wmi_initialization_failed and 
                        self.login_logout_poller):
                        self._poll_login_logout_events()

                except Exception as loop_error:
                    self.activity_logger.debug_log(f"ERROR in monitoring loop iteration #{loop_count}: {loop_error}")
                    time.sleep(5)

        except KeyboardInterrupt:
            self.activity_logger.debug_log("Received keyboard interrupt.")
        except Exception as e:
            self.activity_logger.debug_log(f"Unexpected error in main loop: {e}")
        finally:
            # Final save before shutdown
            self.persistence.save_tracking_data(self.tracker, self.background_video_tracker)
            
            self.running = False
            if self.tracker:
                self.tracker.stop()
            if self.background_video_tracker:
                self.background_video_tracker.stop()
            
            # Clean up WMI connection
            if self.login_logout_poller and hasattr(self.login_logout_poller, 'wmi_connection'):
                if hasattr(self.login_logout_poller.wmi_connection, '_cleanup_connection'):
                    self.login_logout_poller.wmi_connection._cleanup_connection()

    def _poll_login_logout_events(self):
        """Enhanced login/logout polling with status monitoring"""
        if not self.wmi_initialization_complete.is_set() or self.wmi_initialization_failed or not self.login_logout_poller:
            return
        
        try:
            events = self.login_logout_poller.poll_events()
            
            for event in events:
                timestamp = getattr(event, 'TimeGenerated', 'Unknown time')
                if hasattr(event, 'EventCode'):
                    if event.EventCode == 4624:
                        self.activity_logger.buffer_login_logout_event(f"User logged in at {timestamp}")
                    elif event.EventCode == 4634:
                        self.activity_logger.buffer_login_logout_event(f"User logged out at {timestamp}")
            
            if events:
                self.activity_logger.flush_buffer()
                
            # Periodically log status for debugging
            if not hasattr(self, 'last_status_log'):
                self.last_status_log = 0
            
            if time.time() - self.last_status_log > 300:  # Every 5 minutes
                status = self.login_logout_poller.get_status()
                self.activity_logger.debug_log(f"Login/logout poller status: {status}")
                self.last_status_log = time.time()
                
        except Exception as e:
            self.activity_logger.debug_log(f"Error in login/logout polling: {e}")

    def generate_daily_report(self):
        """Generate daily report and save locally (no emailing)"""
        productivity_data = self._collect_productivity_data()

        report_content = self.report_generator.generate_daily_report(productivity_data)
        report_path = self.report_generator.save_report_to_file(report_content, productivity_data.date)
        self.activity_logger.debug_log(f"Daily report generated and saved to: {report_path}")

    def generate_and_email_daily_report(self) -> bool:
        """Generate daily report and email it"""
        productivity_data = self._collect_productivity_data()

        report_content = self.report_generator.generate_daily_report(productivity_data)
        report_path = self.report_generator.save_report_to_file(report_content, productivity_data.date)
        self.activity_logger.debug_log(f"Daily report generated and saved to: {report_path}")

        return self.email_manager.send_email_with_timing_update(
            "Daily Productivity Report",
            "Attached is the daily productivity report.",
            report_path
        )

    def _collect_productivity_data(self) -> ProductivityData:
        """UPDATED: Collect all productivity data including uncategorized websites and background video time"""
        
        try:
            app_times = self.tracker.get_app_times()
            
            # Safely categorize apps with error handling
            productive_apps = {}
            unproductive_apps = {}
            uncategorized_apps = {}
            
            for app, secs in app_times.items():
                try:
                    category = AppCategorizer.categorize_app(app)
                    if category == Category.PRODUCTIVE:
                        productive_apps[app] = secs
                    elif category == Category.UNPRODUCTIVE:
                        unproductive_apps[app] = secs
                    elif category == Category.UNCATEGORIZED:
                        uncategorized_apps[app] = secs
                    # category == None is ignored (system apps, etc.)
                except Exception as e:
                    self.logger.warning(f"Error categorizing app '{app}': {e}")
                    # Default uncategorized if categorization fails
                    uncategorized_apps[app] = secs

            productive_time = sum(productive_apps.values())
            unproductive_time = sum(unproductive_apps.values())

            # Get background video data safely
            try:
                background_video_apps = self.background_video_tracker.get_background_video_times()
                background_video_time = self.background_video_tracker.get_total_background_video_time()
                verified_playing_apps = self.background_video_tracker.get_verified_playing_times()
                verified_playing_time = self.background_video_tracker.get_total_verified_playing_time()
            except Exception as e:
                self.logger.warning(f"Error getting background video data: {e}")
                background_video_apps = {}
                background_video_time = 0
                verified_playing_apps = {}
                verified_playing_time = 0

            # Create simplified background video sessions
            background_video_sessions = []
            for site_name, total_time in background_video_apps.items():
                session_id = f"{site_name}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                
                background_video_sessions.append(
                    BackgroundVideoSession(
                        browser="Unknown",
                        site=site_name,
                        title=site_name,
                        start_time=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        duration=total_time,
                        impact_level=self.report_generator._classify_video_impact(site_name),
                        session_id=session_id
                    )
                )

            # Get system information safely
            try:
                system_info = self._collect_system_info()
            except Exception as e:
                self.logger.warning(f"Error collecting system info: {e}")
                system_info = None

            return ProductivityData(
                productive_time=productive_time,
                unproductive_time=unproductive_time,
                background_video_time=background_video_time,
                verified_playing_time=verified_playing_time,
                productive_apps=productive_apps,
                unproductive_apps=unproductive_apps,
                uncategorized_apps=uncategorized_apps,  # This should never be None now
                background_videos=background_video_sessions,
                background_video_apps=background_video_apps,
                verified_playing_apps=verified_playing_apps,
                date=datetime.datetime.now().strftime('%Y-%m-%d'),
                system_info=system_info
            )
            
        except Exception as e:
            self.logger.error(f"Critical error in _collect_productivity_data: {e}")
            self.logger.error(f"Full traceback: {traceback.format_exc()}")
            
            # Return minimal data structure to prevent crashes
            return ProductivityData(
                productive_time=0,
                unproductive_time=0,
                background_video_time=0,
                verified_playing_time=0,
                productive_apps={},
                unproductive_apps={},
                uncategorized_apps={},  # Always provide this
                background_videos=[],
                background_video_apps={},
                verified_playing_apps={},
                date=datetime.datetime.now().strftime('%Y-%m-%d'),
                system_info=None
            )
        
class ChainedSessionTracker(SessionTracker):
    """Session tracker that creates session chains for long continuous usage"""
    
    def __init__(self):
        super().__init__()
        self.session_chunk_duration = 14400  # 4 hours per chunk
        self.inactivity_threshold = 1800     # 30 minutes = real session end
        
    def parse_login_logout_events(self, events: List[str]) -> List[LoginSession]:
        """Create session chains for long continuous usage"""
        sessions = []
        pending_login = None
        
        # Sort events by timestamp
        sorted_events = self._sort_events_by_time(events)
        
        for event in sorted_events:
            event_time, event_type = self._parse_event(event)
            
            if not event_time:
                continue
            
            if event_type == "login" or event_type == "startup":
                # Handle new login
                if pending_login:
                    # Previous login without logout - create session chain
                    sessions.extend(self._create_session_chain(pending_login, event_time))
                
                # Start new session
                session_type = "Startup" if event_type == "startup" else "Normal"
                pending_login = LoginSession(
                    login_time=event_time,
                    session_type=session_type
                )
            
            elif event_type == "logout" or event_type == "shutdown":
                if pending_login:
                    # Complete the session normally (may create chain if long)
                    sessions.extend(self._create_session_chain(pending_login, event_time))
                    pending_login = None
                else:
                    # Orphaned logout
                    orphaned_session = LoginSession(
                        login_time=event_time - datetime.timedelta(minutes=30),
                        logout_time=event_time,
                        session_type="Orphaned Logout"
                    )
                    orphaned_session.calculate_duration()
                    sessions.append(orphaned_session)
        
        # Handle ongoing session
        if pending_login:
            current_time = datetime.datetime.now()
            sessions.extend(self._create_session_chain(pending_login, current_time, ongoing=True))
        
        return sessions
    
    def _create_session_chain(self, login_session: LoginSession, end_time: datetime.datetime, ongoing: bool = False) -> List[LoginSession]:
        """Create a chain of 4-hour sessions for long continuous usage"""
        sessions = []
        
        start_time = login_session.login_time
        total_duration = (end_time - start_time).total_seconds()
        
        if total_duration <= self.session_chunk_duration and not ongoing:
            # Normal length session - return as-is
            login_session.logout_time = end_time
            login_session.calculate_duration()
            sessions.append(login_session)
            return sessions
        
        # Long session - break into chunks
        current_start = start_time
        chunk_number = 1
        
        while (end_time - current_start).total_seconds() > self.session_chunk_duration:
            # Create a 4-hour chunk
            chunk_end = current_start + datetime.timedelta(seconds=self.session_chunk_duration)
            
            chunk_session = LoginSession(
                login_time=current_start,
                logout_time=chunk_end,
                session_type=f"Continuous Work (Part {chunk_number})"
            )
            chunk_session.calculate_duration()
            sessions.append(chunk_session)
            
            # Move to next chunk
            current_start = chunk_end
            chunk_number += 1
        
        # Handle remaining time
        if current_start < end_time:
            remaining_duration = (end_time - current_start).total_seconds()
            
            if ongoing:
                # Ongoing session
                final_session = LoginSession(
                    login_time=current_start,
                    session_type=f"Ongoing (Part {chunk_number})"
                )
            else:
                # Final completed chunk
                final_session = LoginSession(
                    login_time=current_start,
                    logout_time=end_time,
                    session_type=f"Continuous Work (Part {chunk_number})"
                )
                final_session.calculate_duration()
            
            sessions.append(final_session)
        
        return sessions
    
    def get_session_summary(self, sessions: List[LoginSession]) -> dict:
        """Enhanced summary that accounts for session chains"""
        summary = super().get_session_summary(sessions)
        
        # Group chained sessions
        continuous_sessions = [s for s in sessions if "Continuous Work" in s.session_type or "Part" in s.session_type]
        normal_sessions = [s for s in sessions if "Continuous Work" not in s.session_type and "Part" not in s.session_type]
        
        # Calculate session chains
        session_chains = self._count_session_chains(sessions)
        
        summary.update({
            'session_chains': len(session_chains),
            'chained_sessions': len(continuous_sessions),
            'normal_sessions': len(normal_sessions),
            'longest_continuous_work': max([chain['total_duration'] for chain in session_chains], default=0),
        })
        
        return summary
    
    def _count_session_chains(self, sessions: List[LoginSession]) -> List[dict]:
        """Group sessions into chains and calculate chain statistics"""
        chains = []
        current_chain = []
        
        for session in sessions:
            if "Part" in session.session_type:
                current_chain.append(session)
            else:
                if current_chain:
                    # Close current chain
                    total_duration = sum(s.duration_seconds or 0 for s in current_chain)
                    chains.append({
                        'sessions': current_chain,
                        'total_duration': total_duration,
                        'start_time': current_chain[0].login_time,
                        'end_time': current_chain[-1].logout_time if current_chain[-1].logout_time else datetime.datetime.now()
                    })
                    current_chain = []
                
                # Add single session as its own "chain"
                if session.duration_seconds:
                    chains.append({
                        'sessions': [session],
                        'total_duration': session.duration_seconds,
                        'start_time': session.login_time,
                        'end_time': session.logout_time or datetime.datetime.now()
                    })
        
        # Handle ongoing chain
        if current_chain:
            total_duration = sum(s.duration_seconds or 0 for s in current_chain)
            chains.append({
                'sessions': current_chain,
                'total_duration': total_duration,
                'start_time': current_chain[0].login_time,
                'end_time': current_chain[-1].logout_time if current_chain[-1].logout_time else datetime.datetime.now()
            })
        
        return chains

# Enhanced session analysis that shows chains
    def _generate_session_analysis_with_chains(self, report_date: str) -> str:
        """Session analysis that properly displays session chains"""
        
        all_events = self._get_all_login_logout_events(report_date)
        
        if not all_events:
            return f"""

    📋 SESSION ANALYSIS
    {'─' * 50}
    No login/logout events recorded today."""

        # Use chained session tracker
        tracker = ChainedSessionTracker()
        sessions = tracker.parse_login_logout_events(all_events)
        chains = tracker._count_session_chains(sessions)
        
        section = f"""

    📋 SESSION ANALYSIS
    {'─' * 50}

    📅 Today's Work Sessions:"""

        for i, chain in enumerate(chains, 1):
            start_time = chain['start_time'].strftime('%H:%M:%S')
            end_time = chain['end_time'].strftime('%H:%M:%S') if chain['end_time'] != datetime.datetime.now() else "ACTIVE"
            total_duration = self._format_duration(chain['total_duration'])
            
            if len(chain['sessions']) == 1:
                # Single session
                session = chain['sessions'][0]
                if session.session_type == "Ongoing":
                    section += f"\n   {i}. {start_time} → {end_time} ({total_duration}) 🔄 Ongoing"
                else:
                    section += f"\n   {i}. {start_time} → {end_time} ({total_duration}) ✅ Complete"
            else:
                # Session chain
                parts = len(chain['sessions'])
                ongoing = any("Ongoing" in s.session_type for s in chain['sessions'])
                
                if ongoing:
                    section += f"\n   {i}. {start_time} → {end_time} ({total_duration}) 🔗 Long Session ({parts} parts, ongoing)"
                else:
                    section += f"\n   {i}. {start_time} → {end_time} ({total_duration}) 🔗 Long Session ({parts} parts)"
                
                # Show individual parts
                for j, part in enumerate(chain['sessions'], 1):
                    part_start = part.login_time.strftime('%H:%M:%S')
                    part_end = part.logout_time.strftime('%H:%M:%S') if part.logout_time else "ACTIVE"
                    part_duration = self._format_duration(part.duration_seconds or 0)
                    section += f"\n      └─ Part {j}: {part_start} → {part_end} ({part_duration})"
        
        # Add explanation
        long_sessions = [c for c in chains if len(c['sessions']) > 1]
        if long_sessions:
            section += f"""

    💡 NOTE: Long continuous work sessions are automatically split into 4-hour chunks
    for better time tracking. Total time is preserved across all parts."""
        
        return section
    
def test_active_hybrid_outlook():
    """Test the new active hybrid behavior with status monitoring"""
    print("🧪 TESTING ACTIVE HYBRID OUTLOOK BEHAVIOR")
    print("=" * 60)
    
    try:
        config = CompleteEnhancedConfig()
        hybrid_manager = HybridOutlookManager(config, "test@example.com")
        
        # Show initial status
        print("📊 Initial Store Outlook Status:")
        initial_status = hybrid_manager.get_store_outlook_status()
        print(f"   Available: {'✅' if initial_status['available'] else '❌'}")
        print(f"   Currently Running: {'✅' if initial_status['currently_running'] else '❌'}")
        print(f"   Process Count: {initial_status['process_count']}")
        
        # Test 1: Setup hybrid environment (should launch Store Outlook only if needed)
        print("\n🔧 Test 1: Setting up hybrid environment...")
        setup_results = hybrid_manager.setup_hybrid_environment()
        
        print(f"   Store UI Launched: {'✅' if setup_results['store_ui_launched'] else '❌'}")
        print(f"   Desktop Email Ready: {'✅' if setup_results['desktop_email_ready'] else '❌'}")
        print(f"   Hybrid Ready: {'✅' if setup_results['hybrid_ready'] else '❌'}")
        
        # Test 2: Check status after setup
        print(f"\n📊 Status After Setup:")
        after_status = hybrid_manager.get_store_outlook_status()
        print(f"   Currently Running: {'✅' if after_status['currently_running'] else '❌'}")
        print(f"   Process Count: {after_status['process_count']}")
        print(f"   Launched by Script: {'✅' if after_status['launched_by_script'] else '❌'}")
        
        # Test 3: Test sending email (should NOT launch another instance)
        print(f"\n📧 Test 3: Testing email sending (should not launch duplicate)...")
        
        if setup_results['desktop_email_ready']:
            # Create temporary test file
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
                f.write("Test attachment content")
                temp_file = f.name
            
            print(f"   Testing hybrid email process...")
            # This should NOT launch another Outlook instance
            print(f"   (This should detect existing Outlook and not launch duplicate)")
            
            # We won't actually send, just test the detection
            is_running_before = hybrid_manager.is_store_outlook_running()
            print(f"   Store Outlook running before: {'✅' if is_running_before else '❌'}")
            
            # Clean up
            os.unlink(temp_file)
        else:
            print(f"   Email test skipped: ❌ Desktop Outlook not ready")
        
        # Final status
        print(f"\n📊 Final Status:")
        final_status = hybrid_manager.get_store_outlook_status()
        print(f"   Process Count: {final_status['process_count']}")
        print(f"   Running Processes:")
        for proc in final_status['running_processes']:
            print(f"     • PID {proc['pid']}: {proc['name']}")
        
        print(f"\n🎯 SUMMARY:")
        if setup_results['store_ui_launched'] and setup_results['desktop_email_ready']:
            print(f"   🎉 PERFECT: Store Outlook opened + Email ready")
            print(f"   📱 User can see Outlook (Microsoft Store) interface")
            print(f"   📧 Emails will send automatically via desktop")
            print(f"   🚫 No duplicate instances will be created")
        elif setup_results['desktop_email_ready']:
            print(f"   📧 GOOD: Email functionality available")
            print(f"   ⚠️ Store Outlook UI not available")
        else:
            print(f"   ❌ ISSUES: Email functionality not working")
        
        return setup_results
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return None

def main():
    """Simple main entry point with Friday-only check"""
    
    # Initialize config to check Friday-only setting
    config = CompleteEnhancedConfig()
    config_manager = ConfigManager(config.CONFIG_PATH)
    
    # Check Friday-only mode FIRST
    if config_manager.is_friday_only_enabled():
        current_day = datetime.datetime.now().weekday()  # 0=Monday, 4=Friday
        day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        
        if current_day != 4:  # Not Friday
            print(f"🗓️ Friday-only mode enabled")
            print(f"📅 Today is {day_names[current_day]}")
            print(f"🚫 Script only runs on Fridays - exiting now")
            print(f"⏳ Will run again next Friday")
            return  # EXIT IMMEDIATELY
        else:
            print(f"🎉 Friday detected! Friday-only mode allowing execution")
    
    # If we get here, either friday_only=false OR it's Friday
    print("🚀 Starting productivity monitoring...")
    
    test_your_specific_outlook()
    test_active_hybrid_outlook()
    test_new_email_timing()
    test_friday_only_mode()

    monitor = ActivityMonitor()
    monitor.run()
    
def test_new_email_timing():
    """Test the new email timing system"""
    print("🧪 TESTING NEW EMAIL TIMING SYSTEM")
    print("=" * 50)
    
    try:
        config = CompleteEnhancedConfig()
        config_manager = EnhancedConfigManager(config.CONFIG_PATH)
        
        # Show current configuration
        status = config_manager.get_email_timing_status()
        print(f"📧 Current Configuration:")
        print(f"   Mode: {status['mode']}")
        print(f"   To Email: {status['to_email']}")
        
        if status['mode'] == 'interval':
            print(f"   Interval: {status['interval_formatted']}")
        elif status['mode'] == 'time_of_day':
            print(f"   Daily Time: {status['daily_time']}")
        
        print(f"   Next Email: {status['next_email_info']}")
        
        # Test timing logic
        email_manager = EnhancedHybridOutlookManager(config, config_manager)
        should_send, reason = email_manager.should_send_productivity_report(time.time() - 3600)
        
        print(f"📊 Timing Test:")
        print(f"   Should send now: {should_send}")
        print(f"   Reason: {reason}")
        
        print("✅ Email timing system test completed")
        return True
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_complete_enhanced_system():
    """Test the complete enhanced log management and missed report system"""
    print("🧪 TESTING COMPLETE ENHANCED SYSTEM")
    print("=" * 80)
    
    try:
        config = CompleteEnhancedConfig()
        logger = CompleteEnhancedActivityLogger(config)
        persistence = CompleteEnhancedProductivityDataPersistence(config)
        
        print(f"📂 CONFIGURATION:")
        print(f"   Log directory: {config.LOG_DIR}")
        print(f"   Log cleanup: {config.CLEANUP_DAYS_TO_KEEP} days")
        print(f"   Dated backups: {'Enabled' if config.ENABLE_DATED_BACKUPS else 'Disabled'}")
        print(f"   Missed reports: {'Enabled' if config.ENABLE_MISSED_REPORT_RECOVERY else 'Disabled'}")
        print(f"   Backup cleanup: {config.DATED_BACKUP_DAYS_TO_KEEP} days")
        print(f"   Max log size: {config.MAX_LOG_SIZE_MB}MB")
        
        print(f"\n📊 CURRENT STATUS:")
        
        # Check log files
        log_files = [
            (config.DEBUG_LOG, "Debug log"),
            (config.ACTIVITY_LOG, "Activity log"),
            (logger.sent_reports_file, "Sent reports"),
        ]
        
        for log_file, description in log_files:
            if os.path.exists(log_file):
                size_mb = os.path.getsize(log_file) / (1024 * 1024)
                print(f"   {description}: {size_mb:.2f}MB")
            else:
                print(f"   {description}: Not created yet")
        
        # Check dated backups
        if config.ENABLE_DATED_BACKUPS and os.path.exists(persistence.dated_backup_dir):
            backup_files = glob.glob(os.path.join(persistence.dated_backup_dir, "*.json"))
            print(f"   Dated backups: {len(backup_files)} files")
        else:
            print(f"   Dated backups: Directory not created")
        
        # Test sent reports tracking
        print(f"\n📧 MISSED REPORT TRACKING:")
        today = datetime.datetime.now().strftime('%Y-%m-%d')
        yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        
        print(f"   Today ({today}) sent: {'✅' if logger.was_report_sent(today) else '❌'}")
        print(f"   Yesterday ({yesterday}) sent: {'✅' if logger.was_report_sent(yesterday) else '❌'}")
        
        unsent_dates = logger.get_unsent_recent_dates()
        if unsent_dates:
            print(f"   Unsent recent dates: {unsent_dates}")
        else:
            print(f"   ✅ No unsent reports in last {config.MISSED_REPORT_DAYS_BACK} days")
        
        print(f"\n✅ COMPLETE ENHANCED SYSTEM TEST PASSED")
        return True
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False
    
def test_friday_only_mode():
    """Test the Friday-only functionality"""
    print("🧪 TESTING FRIDAY-ONLY MODE")
    print("=" * 40)
    
    config = CompleteEnhancedConfig()
    config_manager = ConfigManager(config.CONFIG_PATH)
    
    friday_only = config_manager.is_friday_only_enabled()
    current_day = datetime.datetime.now().weekday()
    day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    print(f"Config file exists: {os.path.exists(config.CONFIG_PATH)}")
    print(f"Friday-only enabled: {friday_only}")
    print(f"Today: {day_names[current_day]}")
    print(f"Is Friday: {current_day == 4}")
    
    if friday_only:
        if current_day == 4:
            print("✅ Would run (Friday-only mode + it's Friday)")
        else:
            print("🚫 Would exit immediately (Friday-only mode + not Friday)")
    else:
        print("✅ Would run (Friday-only mode disabled)")
    
    print("=" * 40)
    return friday_only, current_day == 4    

def test_your_specific_outlook():
    """Test function specifically for your Outlook installation"""
    print("🎯 TESTING YOUR SPECIFIC OUTLOOK INSTALLATION")
    print("=" * 60)
    
    detector = ImprovedStoreOutlookDetector()
    
    # Test the specific detection method
    result = detector._check_your_specific_outlook()
    
    if result:
        print(f"✅ SUCCESS: Found your Outlook installation!")
        print(f"   Path: {result['path']}")
        print(f"   Executable: {result['executable']}")
        print(f"   Executable Name: {result['executable_name']}")
        print(f"   Detection Method: {result['detection_method']}")
        
        # Test if it can be launched
        try:
            print(f"\n🚀 Testing if {result['executable_name']} can be accessed...")
            if os.access(result['executable'], os.X_OK):
                print("✅ Executable has proper permissions")
            else:
                print("❌ Executable lacks execution permissions")
        except Exception as e:
            print(f"❌ Error testing executable: {e}")
            
    else:
        print("❌ Could not find your specific Outlook installation")
        
        # Debug: Check if the directory exists
        exact_path = r"C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2025.522.100_x64__8wekyb3d8bbwe"
        print(f"\n🔍 Debug info:")
        print(f"   Exact path exists: {os.path.exists(exact_path)}")
        
        if os.path.exists(exact_path):
            print(f"   Directory contents:")
            try:
                files = os.listdir(exact_path)
                for file in files:
                    if file.lower().endswith('.exe'):
                        print(f"     • {file}")
            except PermissionError:
                print(f"     ❌ Permission denied to list directory")
        
        # Try pattern matching
        pattern = r"C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_*"
        matches = glob.glob(pattern)
        print(f"   Pattern matches: {len(matches)}")
        for match in matches:
            print(f"     • {match}")
    
    print("=" * 60)
    return result is not None

if __name__ == "__main__":
    main()
    test_complete_enhanced_system()