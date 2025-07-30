# snalarm

# Outlook ServiceNow Alarm

## Overview
This Python script monitors your **Outlook Inbox** for **unread** messages from the **`servicenow.com`** domain.  
When at least one matching message is found, it plays an alarm sound (`Ring08.wav`) on loop until you **press Enter**.  
If the email remains unread, it will alarm again on the next poll.

- **Checks interval:** every 5 minutes (configurable)
- **Sound:** `pygame.mixer.Sound('Ring08.wav')` (must be present)
- **Outlook access:** COM via `Dispatch("Outlook.Application").GetNamespace("MAPI")`
- **Imports used:** `win32com.client`, `os`, `time`, `pygame`

---

## Features
- Detects **unread** emails from **`servicenow.com`**.
- **Audible alarm** loops until you press **Enter**.
- **Re-alarms** on the next poll if the email remains unread.
- **Read-only:** does not modify messages or EntryIDs.

---

## Requirements
- **Windows** with **Microsoft Outlook** installed and configured.
- **Python 3.8+**  
- Python packages:
  - `pywin32` (for `win32com.client`)
  - `pygame`

Install the required packages:
```bash
pip install pywin32 pygame

---

## Files & Setup

* Place the script (e.g., `sn_alarm.py`) and the **`Ring08.wav`** sound file in the same folder.

Example:

```
C:\Scripts\sn_alarm\
  ├─ sn_alarm.py
  └─ Ring08.wav
```

---

## How It Works

1. Connects to Outlook:

   ```python
   ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
   inbox = ns.GetDefaultFolder(6)  # Inbox
   ```
2. Every `poll_seconds` (default 5 min), retrieves **unread** items:

   ```python
   items = inbox.Items.Restrict("[Unread] = True")
   ```
3. Scans for senders containing `servicenow.com`.
4. If found:

   * Plays `Ring08.wav` on loop:

     ```python
     alarm_sound.play(loops=-1)
     ```
   * Waits for you to press **Enter**, then stops the alarm and continues monitoring.

---

## Running the Script

Run the script from the folder containing `sn_alarm.py` and `Ring08.wav`:

```bash
python sn_alarm.py
```

To silence the alarm: **press Enter**
To exit the program: **Ctrl+C**

---

## Configuration

### Poll interval

Change the interval in seconds:

```python
if __name__ == "__main__":
    monitor_inbox_and_alarm(poll_seconds=300)  # default 5 min
```

### Domain filter

Update the domain if needed:

```python
domain_flag = "servicenow.com"
```

### Sound file location

Ensure `Ring08.wav` is in the same folder as the script.
To use a different path:

```python
sound_path = r"C:\Path\To\MySound.wav"
```

---

## Behavior

* **Re-alarm:** If the message is still unread on the next poll, the alarm will sound again.
* **Multiple matches:** Any single unread message from the domain triggers the alarm.
* **Read-only:** It does not mark items as read or alter Outlook data.

---

## Troubleshooting

### `Ring08.wav` not found

* Ensure the file is in the same folder as the script.

### No sound / mixer error

* Verify your audio device works and is not locked by another application.
* On Remote Desktop sessions, ensure audio redirection is enabled.

### Outlook COM errors

* Install `pywin32`:

  ```bash
  pip install pywin32
  ```
* Confirm Outlook is installed and accessible under the current user profile.

### Script never triggers

* Confirm the message is **Unread**.
* Ensure the sender address contains `servicenow.com`.
* Only the default Inbox is monitored; check rules moving emails to subfolders.

---

## Limitations

* **Inbox only:** Subfolders are not monitored.
* **Domain match:** Simple substring match on `servicenow.com`.
* **Blocking alarm:** While the alarm plays, the script waits for Enter.

---

## FAQ

**Q:** Will it alarm again if I don’t read the email?
**A:** Yes, the script re-triggers the alarm on the next poll if the email remains unread.

**Q:** Does it mark the email as read or change anything in Outlook?
**A:** No, it is strictly read-only.

**Q:** Can I use another WAV file?
**A:** Yes, place a compatible WAV and update the filename in the code if needed.

---

## Quick Reference

* **Start:** `python sn_alarm.py`
* **Silence alarm:** Press **Enter**
* **Exit program:** **Ctrl+C**
* **Poll interval:** change `poll_seconds` (e.g., `monitor_inbox_and_alarm(120)`)
