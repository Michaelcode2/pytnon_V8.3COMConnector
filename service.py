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
import pythoncom

app = Flask(__name__)

# Configure logging with rotation
def get_log_dir():
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, 'logs')

LOG_DIR = get_log_dir()
LOG_FILE = os.path.join(LOG_DIR, 'api_service.log')

DEBUG_MODE = False  # Set to False when building for production

def get_settings_path():
    if getattr(sys, 'frozen', False):
        # First try to find settings next to the executable
        exe_dir = os.path.dirname(sys.executable)
        exe_settings = os.path.join(exe_dir, 'settings')
        if os.path.exists(exe_settings):
            return exe_settings
        # Fall back to bundled settings
        return os.path.join(sys._MEIPASS, 'settings')
    else:
        # Development mode
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings')

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
        pythoncom.CoInitialize()
        self.initialize_connector()

    def __del__(self):
        try:
            pythoncom.CoUninitialize()  # Cleanup COM
        except:
            pass

    def initialize_connector(self):
        try:
            # Read connection string from config
            config_path = os.path.join(get_settings_path(), 'connection.cfg')
            with open(config_path, 'r', encoding='utf-8') as f:
                connection_string = f.read().strip()
            
            # Add debug logging
            logging.info(f"Using connection string from: {config_path}")
            logging.info(f"Connection string: {connection_string}")
            
            self.v8_connector = Dispatch('V83.COMConnector')
            self.app = self.v8_connector.Connect(connection_string)
            logging.info("COM connector initialized successfully")
        except Exception as e:
            logging.error(f"Failed to initialize COM connector: {str(e)}")
            raise

    def get_product_by_scan_code(self, scan_code):
        try:
            # Updated query filename
            query_text = QueryHandler.load_query('query_price.sql')
            
            # Execute query
            query = self.app.NewObject("Query")
            query.Text = query_text
            query.SetParameter("barcode", scan_code)
            result = query.Execute()
            selection = result.Choose()
            
            if selection.Next():
                # Format response maintaining existing contract
                return {
                    "name": str(selection.Номенклатура if selection.Номенклатура else ""),
                    "measurement": str(selection.ЕдиницаИзмерения if selection.ЕдиницаИзмерения else "шт"),
                    "price": float(selection.Цена or 0),
                    "discountPrice": None
                }
                
            logging.warning(f"No product found for barcode {scan_code}")
            return None
            
        except Exception as e:
            logging.error(f"Error fetching product data: {str(e)}")
            raise

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
        logging.info(f"Received request for scan_code: {scan_code}")
        if not product_service:
            logging.error("Service not initialized")
            return {'error': 'Service not initialized'}, 503
        
        # Validate EAN13 format
        if not is_valid_ean13(scan_code):
            logging.warning(f"Invalid EAN13 barcode format: {scan_code}")
            return {
                'error': 'Invalid barcode format',
                'details': 'Barcode must be a valid EAN13 format (13 digits with valid checksum)'
            }, 400

        product_data = product_service.get_product_by_scan_code(scan_code)
        logging.info(f"Product data retrieved for {scan_code}: {product_data}")
        return jsonify(product_data)
    except Exception as e:
        logging.error(f"Error processing request for scan_code {scan_code}: {str(e)}")
        return {'error': 'Failed to fetch product data'}, 500

class APIService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PriceCheckerService"
    _svc_display_name_ = "PriceChecker Service"
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

class QueryHandler:
    @staticmethod
    def load_query(filename):
        query_path = os.path.join(get_settings_path(), filename)
        with open(query_path, 'r', encoding='utf-8') as f:
            return f.read()

if __name__ == '__main__':
    if DEBUG_MODE:
        try:
            # Initialize logging
            setup_logging()
            logging.info("Starting in DEBUG mode")
            
            product_service = ProductService()
            
            # Run Flask app directly in debug mode
            app.run(host='0.0.0.0', port=8880, debug=True)
            
        except Exception as e:
            logging.error(f"Debug mode error: {str(e)}")
            sys.exit(1)
    else:
        # Normal service mode
        if len(sys.argv) == 1:
            try:
                servicemanager.Initialize()
                servicemanager.PrepareToHostSingle(APIService)
                servicemanager.StartServiceCtrlDispatcher()
            except Exception as e:
                logging.error(f"Service failed to start: {str(e)}")
                sys.exit(1)
        else:
            win32serviceutil.HandleCommandLine(APIService)