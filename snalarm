import win32com.client
import os
import time
import pygame

# ---------------------------------------------
# Audio setup (loads Ring08.wav from the script folder)
# ---------------------------------------------
def init_audio_and_load_sound():
    pygame.init()
    pygame.mixer.init()

    # Resolve Ring08.wav path relative to this script; fallback to CWD if needed
    if "__file__" in globals():
        base_dir = os.path.dirname(os.path.abspath(__file__))
    else:
        base_dir = os.getcwd()
    sound_path = os.path.join(base_dir, "Ring08.wav")

    if not os.path.isfile(sound_path):
        print("ERROR: 'Ring08.wav' not found at:", sound_path)
        print("Place Ring08.wav next to this script or update 'sound_path'. Exiting.")
        raise SystemExit(1)

    try:
        sound = pygame.mixer.Sound(sound_path)
    except Exception as e:
        print("ERROR: Unable to load 'Ring08.wav':", e)
        raise SystemExit(1)

    return sound

# ---------------------------------------------
# Outlook helpers (read-only)
# ---------------------------------------------
def get_outlook_inbox():
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return ns.GetDefaultFolder(6)  # 6 -> olFolderInbox

def is_mail_item(item):
    try:
        return getattr(item, "Class", None) == 43  # 43 -> olMail
    except Exception:
        return False

def from_servicenow(item):
    domain_flag = "servicenow.com"
    try:
        addr = (getattr(item, "SenderEmailAddress", "") or "").lower()
        if domain_flag in addr:
            return True
    except Exception:
        pass
    try:
        sender = getattr(item, "Sender", None)
        if sender:
            try:
                ex_user = sender.GetExchangeUser()
            except Exception:
                ex_user = None
            if ex_user:
                smtp = (getattr(ex_user, "PrimarySmtpAddress", "") or "").lower()
                if domain_flag in smtp:
                    return True
    except Exception:
        pass
    return False

# ---------------------------------------------
# Monitor loop (no de-duplication)
# ---------------------------------------------
def monitor_inbox_and_alarm(poll_seconds=300):
    print("Connecting to Outlook…")
    inbox = get_outlook_inbox()
    print(f"Connected. Monitoring unread emails from 'servicenow.com' every {poll_seconds} seconds.")
    print("If a matching unread email exists, the alarm plays until you press Enter;")
    print("if it remains unread, you'll hear the alarm again on the next poll.")

    alarm_sound = init_audio_and_load_sound()

    try:
        while True:
            # Restrict to unread (read-only)
            try:
                items = inbox.Items.Restrict("[Unread] = True")
            except Exception:
                items = inbox.Items

            # Sort newest first (read-only)
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            # Trigger alarm if any unread from servicenow.com exists
            found_match = False
            for item in items:
                if not is_mail_item(item):
                    continue
                if from_servicenow(item):
                    found_match = True
                    break

            if found_match:
                print(">>> Unread email detected from 'servicenow.com'. Starting alarm…")
                print(">>> Press Enter to stop the alarm; monitoring will continue.")
                alarm_sound.play(loops=-1)

                # Wait for Enter to silence alarm
                try:
                    input()
                except KeyboardInterrupt:
                    print("\nStopping alarm and exiting…")
                    break
                finally:
                    pygame.mixer.stop()
                    print("Alarm silenced. Continuing to monitor…")

            time.sleep(poll_seconds)

    except KeyboardInterrupt:
        print("\nExiting…")
    finally:
        try:
            pygame.mixer.stop()
            pygame.mixer.quit()
            pygame.quit()
        except Exception:
            pass

if __name__ == "__main__":
    monitor_inbox_and_alarm(poll_seconds=300)  # 5 minutes