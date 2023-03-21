from logging import getLogger,StreamHandler,FileHandler,INFO,DEBUG,WARN,ERROR,CRITICAL
from logging import Formatter
import datetime
class ExcelGrepLogger:    
    def __init__(self):

        dt = datetime.datetime.now()
        # dtStr = dt.strftime("%Y%m%d")

        self.logger = getLogger(__name__)
        # Formatter
        formatter = Formatter(
            '[%(levelname)s] %(asctime)s - %(message)s (%(filename)s)'
        )

        # Handler
        self.debHandler = StreamHandler()
        self.debHandler.setLevel(DEBUG)
        self.debHandler.setFormatter(formatter)
        
#        self.errHanler = FileHandler(f"err_{dtStr}.log")
#        self.errHanler.setLevel(ERROR)
#        self.errHanler.setFormatter(formatter)
                
        self.logger.setLevel(DEBUG)
        
        self.logger.addHandler(self.debHandler)
#        self.logger.addHandler(self.errHanler)

    def outDebug(self, msg):
        self.logger.debug(msg)

    def outInfo(self, msg):
        self.logger.info(msg)        

    def outWarning(self, msg):
        self.logger.warning(msg)
        
    def outError(self, msg):
        self.logger.error(msg)
    
    def outCritical(self, msg):
        self.logger.critical(msg)        