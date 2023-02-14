import logging
import time
from win32com.client import Dispatch
from threading import Thread as Th
from queue import Queue
from pathlib import Path
import shutil
import re
from .dirs import BACKUP_FOLDER_LOCATION


logger = logging.getLogger(__name__)

TIME_TO_WAIT_FOR_INDESIGN = 10
TIME_TO_WAIT_BETWEEN_BACKUPS = 10

logger.info(f"Saving backups to {BACKUP_FOLDER_LOCATION}")

class InDesign:
    def __init__(self,app) -> None:
        self.app = app
        import pythoncom
        pythoncom.CoInitialize()


    @classmethod
    def get_app(cls):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            from win32com.client import gencache
            gencache.EnsureModule('{F8817376-B287-45BB-8139-1BF5BFC1FD02}', 0, 1, 0)
            app = Dispatch("InDesign.Application.2023")
            return cls(app)
        except:
            logger.exception("Timed out getting InDesign process")
            return None
    
    def get_open_document(self) -> Path|None:
        try:
            return Path(self.app.ActiveDocument.FullName)
        except Exception as e:
            # Get the message from the exception
            if e.args[1] == 'The RPC server is unavailable.':
                logger.info("InDesign is not running")
                raise TimeoutError("InDesign is not running")
            logger.exception("No open document")
            return None

    def run(self):
        while True:
            try:
                file = self.get_open_document()
            except TimeoutError:
                return
            if not file:
                logger.info("Waiting for open document")
                time.sleep(1)
                continue
            
            logger.info(f"Backing up {file}")
            self.backup(file)

            logger.info(f"Waiting {TIME_TO_WAIT_BETWEEN_BACKUPS} until next check")
            time.sleep(TIME_TO_WAIT_BETWEEN_BACKUPS)
    
    def backup(self, file:Path):
        latest_version, latest_version_number = self.get_latest_version(file)
        
        if latest_version:
            logger.info(f"Latest Backup File: {latest_version}")
            if self.same_file(file, latest_version):
                logger.info(f"File {file} has not changed since last backup.")
                self.save()
            if self.same_file(file, latest_version):
                logger.info(f"File has not been edited since last backup.")
                return
        
        logger.info(f"Backing up {file}")
        new_version_number = latest_version_number + 1
        new_version = BACKUP_FOLDER_LOCATION / f"{file.stem}_{new_version_number}{file.suffix}"
        shutil.copyfile(file, new_version)
        logger.info(f"Backup saved to {new_version}")

    def same_file(self, file1:Path, file2:Path) -> bool:
        s1 = file1.stat().st_size
        s2 = file2.stat().st_size
        logger.info(f"File sizes: {s1} {s2}")
        return s1 == s2

    def get_latest_version(self, file: Path) -> tuple[Path|None,int]:
        latest_version_number = 0
        latest_version = None
        for backup_file in BACKUP_FOLDER_LOCATION.iterdir():
            try:
                stem,n,ext = re.match(r'(.*)_(\d+)(.*)',backup_file.name).groups()
                backup_filename = stem + ext
                if backup_filename == file.name:
                    if int(n) > latest_version_number:
                        latest_version_number = int(n)
                        latest_version = backup_file
            except:
                pass
        return latest_version, latest_version_number
    
    def save(self):
        try:
            self.app.ActiveDocument.Save()
        except:
            logger.error("Failed to save document")

def main():
    while True:
        indesign = InDesign.get_app()
        if indesign:
            indesign.run()
        time.sleep(1)
    