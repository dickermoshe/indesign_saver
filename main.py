import psutil
import logging
import time
from pathlib import Path
import re
import ctypes
from ctypes import wintypes
import shutil
from pyautogui import  hotkey
from threading import Thread as Th


# In home directory
BACKUP_FOLDER_LOCATION = Path.home() / 'Desktop' / 'InDesign Backups'
BACKUP_FOLDER_LOCATION.mkdir(parents=True,exist_ok=True)


class InDesignNotRunning(Exception):
    pass


logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

class InDesign:
    def __init__(self,pid) -> None:
        # Get the InDesign process
        self.pid = pid
        self.proc = psutil.Process(pid)
        logging.info(f'InDesign process found: {self.proc}')
    
    def start(self):
        while True:
            try:
                self.run()
            except InDesignNotRunning:
                logger.info(f"InDesign process {self.pid} has stopped running")
                break
            except:
                logger.exception("An error occurred")
            time.sleep(10)

    def run(self):
        self.sent_save = False
        # Check if the InDesign process is running
        if not self.proc.is_running():
            raise InDesignNotRunning(f"InDesign process {self.pid} has stopped running")
        
        # Check if indesign is active
        if not self.is_active():
            logger.info(f"InDesign process {self.pid} is not active")
            return
        
        # Get the files locked by InDesign
        files = self.files()
        
        if not files:
            logger.info(f"InDesign process {self.pid} has no files open")
            return
        
        logger.info(f"InDesign process {self.pid} has {len(files)} files open")
        
        for file in files:
            self.backup(file)
        
    
    def backup(self, open_file: Path):
        # Iterate through the files in the backup folder
        latest_version_number = 0
        latest_version = None
        for backup_file in BACKUP_FOLDER_LOCATION.iterdir():
            try:
                stem,n,ext = re.match(r'(.*)_(\d+)(.*)',backup_file.name).groups()
                backup_filename = stem + ext
                if backup_filename == open_file.name:
                    if int(n) > latest_version_number:
                        latest_version_number = int(n)
                        latest_version = backup_file
                
            except AttributeError:
                continue
        new_version_number = latest_version_number + 1
        new_version = BACKUP_FOLDER_LOCATION / f'{open_file.stem}_{new_version_number}{open_file.suffix}'

        logger.info(f"Latest Backup File: {latest_version}")
        logger.info(f"New Backup File: {new_version}")

        if latest_version:
            # Check if the file has the same size
            while latest_version.stat().st_size == open_file.stat().st_size:
                if not self.sent_save:
                    self.send_ctrl_s()
                    self.sent_save = True
                    time.sleep(10)
                    continue
                else:
                    logger.info(f"File {open_file.name} has not changed")
                    return

        logger.info(f"Backing up {open_file.name} to {new_version.name}")
        shutil.copy(open_file,new_version)
            
    def files(self) -> list[Path]:
        """
        Get the files locked by InDesign
        """
        try:
            from win32com.client import gencache
            gencache.EnsureModule('{F8817376-B287-45BB-8139-1BF5BFC1FD02}', 0, 1, 0)
            from win32com.client import Dispatch
            app = Dispatch("InDesign.Application.2023")
            return Path(app.ActiveDocument.FullName)
        except:
            return []
        
    def is_active(self):
        # Get the active window PID
        user32 = ctypes.windll.user32
        h_wnd = user32.GetForegroundWindow()
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(h_wnd, ctypes.byref(pid))
        return pid.value == self.proc.pid
    
    @classmethod
    def get_indesigns(cls) -> list["InDesign"]:
        """
        Get all the InDesign process IDs
        """
        indesigns = []
        for proc in psutil.process_iter():
            if proc.name().lower() == 'indesign.exe':
                indesigns.append(cls(proc.pid))
        return indesigns

    def send_ctrl_s(self):
        logger.info(f"Sending Ctrl+S to InDesign process {self.pid}")
        hotkey('ctrl','s')
        
def main():
    threads : dict[int,Th] = {}
    while True:
        # Get all the InDesign processes
        indesigns = InDesign.get_indesigns()
        for indesign in indesigns:
            if indesign.pid not in threads:
                threads[indesign.pid] = Th(target=indesign.start)
                threads[indesign.pid].start()
            else:
                if not threads[indesign.pid].is_alive():
                    # Remove the thread if it has stopped running
                    threads.pop(indesign.pid)
        time.sleep(10)


if __name__ == '__main__':
    main()