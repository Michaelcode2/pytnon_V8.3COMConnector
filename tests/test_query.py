import win32com.client
from query_handler import QueryHandler

def inspect_selection(selection):
    """Inspect all available values in the selection"""
    try:
        print("\n=== Selection Details ===")
        
        if selection.Next():
            print("\nRow Values:")
            # Try to get values for our known columns
            try:
                print(f"Штрихкод: {selection.Штрихкод}")
                print(f"Цена: {selection.Цена}")
                print(f"Номенклатура: {selection.Номенклатура}")
                print(f"ЕдиницаИзмерения: {selection.ЕдиницаИзмерения}")
 
            except Exception as e:
                print(f"Error accessing column: {str(e)}")
        else:
            print("No data found")

    except Exception as e:
        print(f"Error inspecting selection: {str(e)}")

def test_query():
    try:
        # Connect to 1C
        connection_string = r'File="C:\Users\Michael\Documents\DB_Retail";Usr="admin";Pwd="";'
        v8_connector = win32com.client.Dispatch('V83.COMConnector')
        connection = v8_connector.Connect(connection_string)
        
        # Test barcode
        test_barcode = "4820000195447"  # Using your barcode
        
        # Load query
        query_text = QueryHandler.load_query('product_by_barcode.sql')
        
        # Create and execute query directly to inspect results
        query = connection.NewObject("Query")
        query.Text = query_text
        query.SetParameter("barcode", test_barcode)
        
        # Execute query and get selection
        result = query.Execute()
        
        # Inspect the selection object
        inspect_selection(result.Choose())
            
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    test_query() 