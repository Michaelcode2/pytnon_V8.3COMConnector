import win32com.client
import sys

def test_1c_connection():
    try:
        print("Python version:", sys.version)
        print("Python architecture:", "32 bit" if sys.maxsize <= 2**32 else "64 bit")
        
        connection_string = r'File="C:\Users\Michael\Documents\DB_Retail";Usr="admin";Pwd="";'
        
        # Try different COM connector names
        try:
            v8_connector = win32com.client.Dispatch('V83.COMConnector')  # For 8.3
        except:
            try:
                v8_connector = win32com.client.Dispatch('V82.COMConnector')  # For 8.2
            except:
                v8_connector = win32com.client.Dispatch('V1CEnterprise.Connect')  # Legacy
        
        app = v8_connector.Connect(connection_string)
        
        print("Successfully connected to 1C!")
        return True
    except Exception as e:
        print(f"Connection failed: {str(e)}")
        return False

if __name__ == "__main__":
    test_1c_connection() 