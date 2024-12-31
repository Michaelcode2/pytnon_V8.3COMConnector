import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import logging
from logging.handlers import TimedRotatingFileHandler
import os
from datetime import datetime, timedelta
from flask import Flask, jsonify
from waitress import serve
import threading
from win32com.client.dynamic import Dispatch
import re

app = Flask(__name__)

# Configure logging with rotation
LOG_DIR = 'C:\\ServiceLogs'
LOG_FILE = os.path.join(LOG_DIR, 'api_service.log')

def setup_logging():
    # Create logs directory if it doesn't exist
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    
    # Set up the rotating log handler
    handler = TimedRotatingFileHandler(
        LOG_FILE,
        when="midnight",
        interval=1,
        backupCount=10,
        encoding='utf-8'
    )
    
    # Set the log format
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    
    # Configure the root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(handler)

def cleanup_old_logs():
    """Delete log files older than 10 days"""
    try:
        current_time = datetime.now()
        for file in os.listdir(LOG_DIR):
            if not file.startswith('api_service.log'):
                continue
            
            file_path = os.path.join(LOG_DIR, file)
            file_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
            
            if current_time - file_modified > timedelta(days=10):
                os.remove(file_path)
                logging.info(f"Deleted old log file: {file}")
    except Exception as e:
        logging.error(f"Error cleaning up logs: {str(e)}")

class ProductService:
    def __init__(self):
        self.v8_connector = None
        self.app = None
        self.initialize_connector()

    def initialize_connector(self):
        try:
            connection_string = r'File="C:\common\DB\83Empty8310";'
            self.v8_connector = Dispatch('V83.COMConnector')
            self.app = self.v8_connector.Connect(connection_string)
            logging.info("COM connector initialized successfully")
        except Exception as e:
            logging.error(f"Failed to initialize COM connector: {str(e)}")
            raise

    def get_product_by_scan_code(self, scan_code):
        try:
            # TODO: Replace this with actual COM connector query
            # This is a placeholder that will be updated with actual COM integration
            # Example: product_data = self.app.GetProductByScanCode(scan_code)
            
            # Placeholder response
            product_data = {
                "name": "Product Name",
                "measurement": "pcs",
                "price": 10.99,
                "discountPrice": 8.99
            }
            return product_data
        except Exception as e:
            logging.error(f"Error fetching product data: {str(e)}")
            raise

# Create global product service instance
product_service = None

def is_valid_ean13(barcode):
    """
    Validate EAN13 barcode format and checksum
    """
    if not re.match(r'^\d{13}$', barcode):
        return False
    
    # Calculate checksum
    total = 0
    for i in range(12):
        digit = int(barcode[i])
        total += digit if i % 2 == 0 else digit * 3
    
    check_digit = (10 - (total % 10)) % 10
    return check_digit == int(barcode[-1])

@app.route('/api/health')
def health_check():
    try:
        # Check if COM connector is alive
        if product_service and product_service.app:
            return {'status': 'healthy', 'comConnector': 'connected'}
        return {'status': 'degraded', 'comConnector': 'disconnected'}
    except Exception as e:
        logging.error(f"Health check failed: {str(e)}")
        return {'status': 'unhealthy', 'error': str(e)}, 500

@app.route('/products/<scan_code>')
def get_product(scan_code):
    try:
        if not product_service:
            return {'error': 'Service not initialized'}, 503
        
        # Validate EAN13 format
        if not is_valid_ean13(barcode):
            logging.warning(f"Invalid EAN13 barcode format: {barcode}")
            return {
                'error': 'Invalid barcode format',
                'details': 'Barcode must be a valid EAN13 format (13 digits with valid checksum)'
            }, 400

        product_data = product_service.get_product_by_scan_code(scan_code)
        return jsonify(product_data)
    except Exception as e:
        logging.error(f"Error processing request for scan_code {scan_code}: {str(e)}")
        return {'error': 'Failed to fetch product data'}, 500

class APIService(win32serviceutil.ServiceFramework):
    _svc_name_ = "ProductAPIService"
    _svc_display_name_ = "Product API Service"
    _svc_description_ = "Windows Service running a Product API with COM Connector"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        self.stop_requested = True

    def SvcDoRun(self):
        try:
            global product_service
            
            # Initialize logging
            setup_logging()
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )
            
            # Initialize the product service
            product_service = ProductService()
            
            # Run Flask app in a separate thread
            def run_server():
                serve(app, host='0.0.0.0', port=8880)
            
            server_thread = threading.Thread(target=run_server)
            server_thread.daemon = True
            server_thread.start()
            
            # Run log cleanup daily
            def cleanup_logs():
                while not self.stop_requested:
                    cleanup_old_logs()
                    # Sleep for 24 hours
                    win32event.WaitForSingleObject(self.stop_event, 24 * 60 * 60 * 1000)
            
            cleanup_thread = threading.Thread(target=cleanup_logs)
            cleanup_thread.daemon = True
            cleanup_thread.start()
            
            # Keep the service running
            while not self.stop_requested:
                win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)
                
        except Exception as e:
            logging.error(f"Service error: {str(e)}")
            servicemanager.LogErrorMsg(f"Service error: {str(e)}")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(APIService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(APIService)