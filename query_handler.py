import os

class QueryHandler:
    QUERY_DIR = 'queries'
    
    @staticmethod
    def load_query(filename):
        """Load query from file"""
        query_path = os.path.join(QueryHandler.QUERY_DIR, filename)
        with open(query_path, 'r', encoding='utf-8') as file:
            return file.read()
    
    @staticmethod
    def execute_query(connection, query_text, parameters=None):
        """Execute 1C query and return results"""
        try:
            # Create query object
            query = connection.NewObject("Query")
            query.Text = query_text
            
            # Set parameters if provided
            if parameters:
                for param_name, param_value in parameters.items():
                    query.SetParameter(param_name, param_value)
            
            # Execute query
            result = query.Execute()
            
            # Convert to list of dictionaries
            selection = result.Choose()
            data = []
            while selection.Next():
                row = {
                    "barcode": str(selection.Штрихкод),
                    "price": float(selection.Цена or 0),
                    "nomenclature": str(selection.Номенклатура)
                }
                data.append(row)
            
            return data
            
        except Exception as e:
            raise Exception(f"Query execution error: {str(e)}") 