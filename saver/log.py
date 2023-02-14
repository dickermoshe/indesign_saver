import logging
import logging.handlers
from .dirs import LOG_LOCATION

fh = logging.handlers.RotatingFileHandler(
    str(LOG_LOCATION / 'indesign_saver.log'),
    backupCount=1,
    maxBytes=50_000_000,
    encoding='utf-8'
)
fh.setLevel(logging.DEBUG)
fh.setFormatter(logging.Formatter(
    '[%(levelname)s] %(name)s: [%(threadName)s] %(message)s'))

sh = logging.StreamHandler()
sh.setLevel(logging.DEBUG)

logger = logging.getLogger(__name__.split('.')[0])
logger.setLevel(logging.DEBUG)
logger.addHandler(fh)
logger.addHandler(sh)
