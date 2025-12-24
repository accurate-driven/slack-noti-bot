import os
import sys
import time
import json
import subprocess
import sqlite3
import socket
import locale
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# Load environment variables
# Handle both script execution and PyInstaller executable
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    script_dir = Path(sys.executable).parent
else:
    # Running as script
    script_dir = Path(__file__).parent.absolute()
env_path = script_dir / '.env'
load_dotenv(dotenv_path=env_path)

class WindowsNotificationMonitor:
    def __init__(self):
        self.slack_token = os.getenv('SLACK_BOT_TOKEN')
        self.slack_channel = os.getenv('SLACK_CHANNEL', '#notifications')
        self.slack_client = WebClient(token=self.slack_token)
        self.processed_notifications = set()
        # Get machine identifier (hostname or custom name from env)
        self.machine_name = os.getenv('MACHINE_NAME', socket.gethostname())
        # Get region/location
        self.region = self._get_region()
    
    def _get_region(self):
        """
        Detect region/location from various sources
        Priority: env variable > timezone > Windows region > locale
        """
        # Check if region is set in env
        env_region = os.getenv('REGION')
        if env_region:
            return env_region
        
        # Try to get from timezone
        try:
            import timezonefinder
            import pytz
            # Get local timezone
            local_tz = datetime.now().astimezone().tzinfo
            if hasattr(local_tz, 'zone'):
                tz_name = local_tz.zone
                # Convert timezone to region name (e.g., America/New_York -> US-East)
                if 'America' in tz_name:
                    if 'New_York' in tz_name or 'Toronto' in tz_name:
                        return 'US-East'
                    elif 'Chicago' in tz_name:
                        return 'US-Central'
                    elif 'Denver' in tz_name:
                        return 'US-Mountain'
                    elif 'Los_Angeles' in tz_name or 'Vancouver' in tz_name:
                        return 'US-West'
                    else:
                        return 'US'
                elif 'Europe' in tz_name:
                    return 'Europe'
                elif 'Asia' in tz_name:
                    if 'Tokyo' in tz_name:
                        return 'Japan'
                    elif 'Shanghai' in tz_name or 'Beijing' in tz_name:
                        return 'China'
                    elif 'Singapore' in tz_name:
                        return 'Singapore'
                    else:
                        return 'Asia'
                elif 'Australia' in tz_name:
                    return 'Australia'
                else:
                    return tz_name.split('/')[-1].replace('_', '-')
        except ImportError:
            pass
        except Exception:
            pass
        
        # Try Windows region settings
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International")
            geo_id = winreg.QueryValueEx(key, "GeoID")[0]
            winreg.CloseKey(key)
            
            # Map GeoID to region (common ones)
            geo_map = {
                244: 'US',  # United States
                39: 'CA',   # Canada
                234: 'GB',  # United Kingdom
                81: 'DE',   # Germany
                84: 'FR',   # France
                110: 'JP',  # Japan
                45: 'CN',   # China
                195: 'AU',  # Australia
            }
            if geo_id in geo_map:
                return geo_map[geo_id]
        except Exception:
            pass
        
        # Fallback to locale
        try:
            loc = locale.getdefaultlocale()[0]
            if loc:
                # Extract country code from locale (e.g., 'en_US' -> 'US')
                if '_' in loc:
                    return loc.split('_')[1]
                return loc
        except Exception:
            pass
        
        # Final fallback
        return 'Unknown'
        
    def get_notifications(self):
        """
        Get Windows notifications - try database first, then PowerShell fallback
        """
        # Try reading from SQLite database first (most reliable)
        notifications = self.get_notifications_from_database()
        if notifications:
            return notifications
        
        # Fallback to PowerShell method
        return self.get_notifications_powershell()
    
    def get_notifications_from_database(self):
        """
        Read notifications directly from Windows notification database (SQLite)
        This reads toast notifications from the Notification table
        """
        try:
            # Path to Windows notification database
            db_path = Path(os.environ.get('LOCALAPPDATA')) / 'Microsoft' / 'Windows' / 'Notifications' / 'wpndatabase.db'
            
            if not db_path.exists():
                return []
            
            # Connect to database (read-only)
            conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            notifications = []
            
            try:
                # Query for toast notifications (not tiles)
                # Join with NotificationHandler to get app information
                # Note: "Group" is a reserved keyword in SQLite, so we use [Group]
                query = """
                    SELECT 
                        n.Id,
                        n.Tag,
                        n.[Group],
                        n.Payload,
                        n.ArrivalTime,
                        h.PrimaryId as AppId
                    FROM Notification n
                    LEFT JOIN NotificationHandler h ON n.HandlerId = h.RecordId
                    WHERE n.Type = 'toast'
                    ORDER BY n.ArrivalTime DESC
                    LIMIT 100
                """
                
                cursor.execute(query)
                rows = cursor.fetchall()
                
                # Debug output
                if len(rows) > 0:
                    print(f"Found {len(rows)} toast notifications in database")
                else:
                    # Check if there are any notifications at all
                    cursor.execute("SELECT COUNT(*) as count FROM Notification WHERE Type = 'toast'")
                    toast_count = cursor.fetchone()['count']
                    if toast_count == 0:
                        print("No toast notifications found in database (this is normal if none exist)")
                    else:
                        print(f"Query found 0 rows but database has {toast_count} toast notifications - possible query issue")
                
                for row in rows:
                    # Group is a reserved keyword - access by column index (position 2 in SELECT)
                    # SELECT order: Id(0), Tag(1), Group(2), Payload(3), ArrivalTime(4), AppId(5)
                    group_val = row[2] if len(row) > 2 else ''
                    tag_val = row['Tag'] if 'Tag' in row.keys() else (row[1] if len(row) > 1 else '')
                    id_val = row['Id'] if 'Id' in row.keys() else (row[0] if len(row) > 0 else '')
                    notification_id = f"{id_val}_{group_val}_{tag_val}"
                    
                    if notification_id not in self.processed_notifications:
                        # Parse the Payload XML
                        payload = row['Payload']
                        title = 'Notification'
                        body = ''
                        app_id = row['AppId'] or 'Unknown'
                        
                        if payload:
                            try:
                                import xml.etree.ElementTree as ET
                                # Payload is bytes, decode it
                                if isinstance(payload, bytes):
                                    xml_str = payload.decode('utf-8', errors='ignore')
                                else:
                                    xml_str = str(payload)
                                
                                root = ET.fromstring(xml_str)
                                
                                # Extract text from toast XML structure
                                # Toast XML: <toast><visual><binding><text>Title</text><text>Body</text></binding></visual></toast>
                                text_elements = root.findall('.//{*}text')
                                if not text_elements:
                                    # Try without namespace
                                    text_elements = root.findall('.//text')
                                
                                if text_elements:
                                    title = text_elements[0].text if text_elements[0].text else 'Notification'
                                    if len(text_elements) > 1:
                                        body = text_elements[1].text if text_elements[1].text else ''
                                
                                # Try to get app name from visual/binding
                                visual = root.find('.//{*}visual')
                                if visual is None:
                                    visual = root.find('.//visual')
                                
                                if visual is not None:
                                    binding = visual.find('.//{*}binding')
                                    if binding is None:
                                        binding = visual.find('.//binding')
                                    if binding is not None and 'template' in binding.attrib:
                                        # Template might have app info
                                        pass
                                
                            except Exception as e:
                                # If XML parsing fails, try to extract any text
                                if isinstance(payload, bytes):
                                    payload_str = payload.decode('utf-8', errors='ignore')
                                else:
                                    payload_str = str(payload)
                                # Simple text extraction as fallback
                                if '<text>' in payload_str:
                                    import re
                                    matches = re.findall(r'<text[^>]*>(.*?)</text>', payload_str, re.DOTALL)
                                    if matches:
                                        title = matches[0].strip() if matches[0].strip() else 'Notification'
                                        if len(matches) > 1:
                                            body = matches[1].strip()
                        
                        # Add notification if we have content
                        if title and (title != 'Notification' or body):
                            notification_data = {
                                'id': notification_id,
                                'title': title if title != 'Notification' else (body[:50] if body else 'Notification'),
                                'body': body,
                                'timestamp': datetime.now().isoformat(),
                                'app_id': app_id
                            }
                            notifications.append(notification_data)
                            self.processed_notifications.add(notification_id)
                            print(f"  → Captured: {notification_data['title'][:50]} from {app_id}")
                        
            except Exception as e:
                # Query failed - print error for debugging
                print(f"Database query error: {e}")
                return []
            
            conn.close()
            return notifications
            
        except sqlite3.OperationalError as e:
            if "database is locked" in str(e).lower():
                # Database is locked by Windows - this is normal, fall back to PowerShell
                return []
            else:
                print(f"Database error: {e}")
                return []
        except Exception as e:
            # Database read failed, fall back to PowerShell
            return []
    
    def get_notifications_powershell(self):
        """
        Get Windows notifications using PowerShell
        Tries multiple methods to access notification history
        """
        # Try method 1: Standard GetHistory()
        notifications = self._try_get_history_standard()
        if notifications:
            return notifications
        
        # Try method 2: GetHistory() for specific apps (Slack, etc.)
        notifications = self._try_get_history_by_apps()
        if notifications:
            return notifications
        
        # If both fail, return empty
        return []
    
    def _try_get_history_standard(self):
        """Try standard GetHistory() method"""
        try:
            ps_script = """
            $ErrorActionPreference = "Stop"
            try {
                Add-Type -AssemblyName System.Runtime.WindowsRuntime
                [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
                
                $history = $null
                try {
                    $history = [Windows.UI.Notifications.ToastNotificationManager]::History.GetHistory()
                } catch {
                    $errorMsg = $_.Exception.Message
                    if ($errorMsg -like "*0x80070490*" -or $errorMsg -like "*Element not found*") {
                        Write-Output "[]"
                        exit 0
                    } else {
                        Write-Error "GetHistory failed: $errorMsg"
                        Write-Output "[]"
                        exit 0
                    }
                }
                
                if ($null -eq $history) {
                    Write-Output "[]"
                    exit 0
                }
                
                $results = @()
                
                foreach ($notif in $history) {
                    try {
                        $xml = $notif.Content.GetXml()
                        $xmlDoc = New-Object System.Xml.XmlDocument
                        $xmlDoc.LoadXml($xml)
                        
                        $textNodes = $xmlDoc.SelectNodes("//text")
                        $title = ""
                        $body = ""
                        
                        if ($textNodes -and $textNodes.Count -gt 0) {
                            $title = $textNodes[0].InnerText
                        }
                        if ($textNodes -and $textNodes.Count -gt 1) {
                            $body = $textNodes[1].InnerText
                        }
                        
                        $results += @{
                            Id = if ($notif.Id) { $notif.Id } else { "" }
                            Group = if ($notif.Group) { $notif.Group } else { "" }
                            Tag = if ($notif.Tag) { $notif.Tag } else { "" }
                            AppId = if ($notif.ApplicationId) { $notif.ApplicationId } else { "Unknown" }
                            Title = $title
                            Body = $body
                        }
                    } catch {
                        continue
                    }
                }
                
                if ($results.Count -eq 0) {
                    Write-Output "[]"
                } else {
                    $results | ConvertTo-Json -Compress
                }
            } catch {
                Write-Error "Script error: $($_.Exception.Message)"
                Write-Output "[]"
            }
            """
            return self._execute_powershell(ps_script)
        except Exception as e:
            return []
    
    def _try_get_history_by_apps(self):
        """Try getting history for specific apps like Slack"""
        try:
            # Common app IDs that send notifications
            app_ids = [
                "com.slack.Slack",
                "slack",
                "Microsoft.Windows.Shell.RunDialog",
                "Microsoft.Windows.Shell.StartMenuExperienceHost"
            ]
            
            ps_script = """
            $ErrorActionPreference = "Stop"
            try {
                Add-Type -AssemblyName System.Runtime.WindowsRuntime
                [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
                
                $allResults = @()
                $appIds = @('com.slack.Slack', 'slack', 'Microsoft.Windows.Shell.RunDialog', 'Microsoft.Windows.Shell.StartMenuExperienceHost')
                
                foreach ($appId in $appIds) {
                    try {
                        $history = [Windows.UI.Notifications.ToastNotificationManager]::History.GetHistory($appId)
                        if ($history -and $history.Count -gt 0) {
                            foreach ($notif in $history) {
                                try {
                                    $xml = $notif.Content.GetXml()
                                    $xmlDoc = New-Object System.Xml.XmlDocument
                                    $xmlDoc.LoadXml($xml)
                                    
                                    $textNodes = $xmlDoc.SelectNodes("//text")
                                    $title = ""
                                    $body = ""
                                    
                                    if ($textNodes -and $textNodes.Count -gt 0) {
                                        $title = $textNodes[0].InnerText
                                    }
                                    if ($textNodes -and $textNodes.Count -gt 1) {
                                        $body = $textNodes[1].InnerText
                                    }
                                    
                                    $allResults += @{
                                        Id = if ($notif.Id) { $notif.Id } else { "" }
                                        Group = if ($notif.Group) { $notif.Group } else { "" }
                                        Tag = if ($notif.Tag) { $notif.Tag } else { "" }
                                        AppId = if ($notif.ApplicationId) { $notif.ApplicationId } else { $appId }
                                        Title = $title
                                        Body = $body
                                    }
                                } catch {
                                    continue
                                }
                            }
                        }
                    } catch {
                        continue
                    }
                }
                
                if ($allResults.Count -eq 0) {
                    Write-Output "[]"
                } else {
                    $allResults | ConvertTo-Json -Compress
                }
            } catch {
                Write-Output "[]"
            }
            """
            return self._execute_powershell(ps_script)
        except Exception as e:
            return []
    
    def _execute_powershell(self, ps_script):
        """Execute PowerShell script and parse results"""
        try:
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            if result.returncode == 0:
                try:
                    output = result.stdout.strip()
                    if not output or output == "[]":
                        return []
                    
                    notifications_json = json.loads(output)
                    if not isinstance(notifications_json, list):
                        notifications_json = [notifications_json]
                    
                    current_notifications = []
                    for notif in notifications_json:
                        notification_id = f"{notif.get('Id', '')}_{notif.get('Group', '')}_{notif.get('Tag', '')}"
                        
                        if notification_id not in self.processed_notifications:
                            notification_data = {
                                'id': notification_id,
                                'title': notif.get('Title', 'Notification') or 'Notification',
                                'body': notif.get('Body', ''),
                                'timestamp': datetime.now().isoformat(),
                                'app_id': notif.get('AppId', 'Unknown')
                            }
                            current_notifications.append(notification_data)
                            self.processed_notifications.add(notification_id)
                    
                    return current_notifications
                except json.JSONDecodeError:
                    return []
            return []
        except Exception:
            return []
            
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            # Debug: Log errors if any
            if result.stderr and result.stderr.strip():
                # Only log unexpected errors (not the "Element not found" which is normal)
                if "Element not found" not in result.stderr and "0x80070490" not in result.stderr:
                    print(f"PowerShell warning: {result.stderr.strip()[:100]}")
            
            if result.returncode == 0:
                try:
                    output = result.stdout.strip()
                    if not output or output == "[]":
                        # No notifications or empty history - this is normal
                        return []
                    
                    notifications_json = json.loads(output)
                    if not isinstance(notifications_json, list):
                        notifications_json = [notifications_json]
                    
                    current_notifications = []
                    for notif in notifications_json:
                        notification_id = f"{notif.get('Id', '')}_{notif.get('Group', '')}_{notif.get('Tag', '')}"
                        
                        if notification_id not in self.processed_notifications:
                            notification_data = {
                                'id': notification_id,
                                'title': notif.get('Title', 'Notification') or 'Notification',
                                'body': notif.get('Body', ''),
                                'timestamp': datetime.now().isoformat(),
                                'app_id': notif.get('AppId', 'Unknown')
                            }
                            current_notifications.append(notification_data)
                            self.processed_notifications.add(notification_id)
                    
                    return current_notifications
                except json.JSONDecodeError as e:
                    # If output is not JSON, it might be an error message
                    if result.stdout.strip():
                        print(f"Error parsing PowerShell JSON output: {e}")
                        print(f"Output: {result.stdout[:200]}")
                    return []
            else:
                # PowerShell returned an error code, but we'll try to parse stderr
                if result.stderr and "Element not found" not in result.stderr:
                    # Only print if it's not the expected "empty history" error
                    print(f"PowerShell error: {result.stderr[:200]}")
                return []
                
        except subprocess.TimeoutExpired:
            print("PowerShell command timed out")
            return []
        except Exception as e:
            print(f"Error in PowerShell method: {e}")
            return []
    
    def _get_timezone_offset(self):
        """
        Get timezone offset in GMT format (e.g., GMT-5, GMT+3)
        """
        try:
            now = datetime.now()
            # Get UTC offset
            utc_offset = now.astimezone().utcoffset()
            if utc_offset:
                # Convert to hours
                offset_hours = utc_offset.total_seconds() / 3600
                # Format as GMT+X or GMT-X
                if offset_hours >= 0:
                    return f"GMT+{int(offset_hours)}"
                else:
                    return f"GMT{int(offset_hours)}"
        except Exception:
            pass
        return "GMT"
    
    def send_to_slack(self, notification):
        """
        Send notification to Slack
        """
        try:
            # Format the message in requested format
            # Format: ----- {desktop-name} | {location} | {app name} | {time}({timezone}) -----
            app_name = notification.get('app_id', 'Unknown')
            # Clean up app name (remove package prefixes)
            if '.' in app_name:
                app_name = app_name.split('.')[-1]
            app_name = app_name.replace('com.squirrel.', '').replace('com.', '').title()
            
            # Format time with timezone
            try:
                notif_time = datetime.fromisoformat(notification.get('timestamp', ''))
                time_str = notif_time.strftime('%H:%M:%S')
            except:
                time_str = datetime.now().strftime('%H:%M:%S')
            
            # Get timezone offset
            tz_offset = self._get_timezone_offset()
            time_with_tz = f"{time_str}({tz_offset})"
            
            # Header line
            header = f"----- {self.machine_name} | {self.region} | {app_name} | {time_with_tz} -----"
            
            # Notification content
            title = notification.get('title', 'Notification')
            body = notification.get('body', '')
            
            # Build message
            message = f"{header}\n"
            if title and title != 'Notification':
                message += f"{title}\n"
            if body:
                message += f"{body}\n"
            
            response = self.slack_client.chat_postMessage(
                channel=self.slack_channel,
                text=message
            )
            
            print(f"Sent notification to Slack: {notification.get('title', 'Unknown')}")
            return True
            
        except SlackApiError as e:
            error_code = e.response.get('error', 'unknown')
            if error_code == 'channel_not_found':
                print(f"❌ Slack Error: Channel '{self.slack_channel}' not found!")
                print(f"   Make sure:")
                print(f"   1. The channel ID/name is correct")
                print(f"   2. The bot is invited to the channel")
                print(f"   3. Or use channel name with # prefix (e.g., #general)")
            else:
                print(f"Slack API Error: {error_code}")
            return False
        except Exception as e:
            print(f"Error sending to Slack: {e}")
            return False
    
    def check_notification_history_access(self):
        """
        Check if notification history is accessible
        """
        try:
            ps_script = """
            try {
                Add-Type -AssemblyName System.Runtime.WindowsRuntime
                [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
                $history = [Windows.UI.Notifications.ToastNotificationManager]::History.GetHistory()
                Write-Output "OK"
            } catch {
                Write-Output "ERROR"
            }
            """
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            return result.stdout.strip() == "OK"
        except:
            return False
    
    def monitor(self, interval=5):
        """
        Monitor notifications and send to Slack
        """
        print(f"Starting Windows notification monitor...")
        print(f"Monitoring every {interval} seconds")
        print(f"Sending to Slack channel: {self.slack_channel}")
        
        # Check if notification history is accessible
        print("\nChecking notification history access...")
        if not self.check_notification_history_access():
            print("\n⚠️  WARNING: Cannot access Windows notification history!")
            print("\nThis usually means:")
            print("  1. Notification history is empty (no notifications stored yet)")
            print("  2. Notification history feature is disabled")
            print("  3. Windows settings need to be configured")
            print("\nTo fix this:")
            print("  1. Open Windows Settings (Win + I)")
            print("  2. Go to System > Notifications")
            print("  3. Make sure notifications are enabled")
            print("  4. Enable 'Get notifications from apps and other senders'")
            print("\nThe bot will continue running and will capture notifications once they appear in history.")
            print("Trigger a test notification (like a Slack message) to test.\n")
        else:
            print("✓ Notification history is accessible\n")
        
        print("Press Ctrl+C to stop\n")
        
        try:
            check_count = 0
            while True:
                notifications = self.get_notifications()
                
                for notification in notifications:
                    self.send_to_slack(notification)
                
                if notifications:
                    print(f"✓ Processed {len(notifications)} new notification(s)")
                else:
                    check_count += 1
                    # Print status every 12 checks (1 minute) to show it's working
                    if check_count % 12 == 0:
                        print(f"Checking... (no new notifications found)")
                
                time.sleep(interval)
                
        except KeyboardInterrupt:
            print("\nStopping monitor...")
        except Exception as e:
            print(f"Error in monitor loop: {e}")

def main():
    # Validate environment variables
    if not os.getenv('SLACK_BOT_TOKEN'):
        print("Error: SLACK_BOT_TOKEN not found in .env file")
        return
    
    monitor = WindowsNotificationMonitor()
    monitor.monitor(interval=5)  # Check every 5 seconds

if __name__ == "__main__":
    main()

