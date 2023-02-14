from pathlib import Path
# In home directory
BACKUP_FOLDER_LOCATION = Path.home() / 'Desktop' / 'InDesign Backups'
BACKUP_FOLDER_LOCATION.mkdir(parents=True, exist_ok=True)
LOG_LOCATION = Path.home() / "InDesign Saver"
LOG_LOCATION.mkdir(parents=True, exist_ok=True)