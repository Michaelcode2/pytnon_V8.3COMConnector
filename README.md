# Windows Service Product API with 1C COM Integration

This project implements a Windows service that exposes a REST API for product information retrieval from 1C database using COM connector. The service handles EAN13 barcode requests and returns product information in JSON format.

## Features

- Windows service implementation using Python
- REST API endpoints for product price checking
- EAN13 barcode validation
- 1C COM connector integration
- Daily log rotation with 10-day retention
- Health check endpoint
- Configurable connection settings
- Debug mode for development
- Conversion to standalone .exe file

## Prerequisites

- Windows OS
- Python 3.7+
- Administrative privileges
- 1C:Enterprise platform with COM connector
- Registered 1C COM connector (`regsvr32 comcntr.dll`)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/product-api-service.git
cd product-api-service
```

2. Install required packages:
```bash
pip install pywin32 flask waitress pyinstaller
```

3. Create the following directory structure:

```
   PriceChecker/
   ├── settings/
   │   ├── connection.cfg    # 1C connection string
   │   └── query_price.sql   # SQL query for price lookup
   ├── service.py
   └── build_service.py
```

4. Make sure 1C COM connector is properly registered:
```bash
# Run as Administrator
regsvr32 comcntr.dll
```

## Building the Executable

1. Run the build script:
```bash
python build_service.py
```

The executable will be created in the `dist` directory as `APIService.exe`.

## Service Installation

Run these commands in an elevated (Administrator) command prompt:

```bash
# Install the service
PriceCheckerService.exe install

# Start the service
PriceCheckerService.exe start

# Stop the service
PriceCheckerService.exe stop

# Remove the service
PriceCheckerService.exe remove
```


## Configuration


### Connection String (settings/connection.cfg)
```
File="path_to_your_1C_database";Usr="username";Pwd="password";
```

### SQL Query (settings/query_price.sql)

```sql
ВЫБРАТЬ
	Штрихкоды.Штрихкод КАК Штрихкод,
	ЦеныНоменклатурыСрезПоследних.Цена КАК Цена,
	Штрихкоды.ЕдиницаИзмерения.Наименование КАК ЕдиницаИзмерения,
	Штрихкоды.Владелец.НаименованиеПолное КАК Номенклатура
ИЗ
	РегистрСведений.Штрихкоды КАК Штрихкоды
		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЦеныНоменклатуры.СрезПоследних КАК ЦеныНоменклатурыСрезПоследних
		ПО Штрихкоды.Владелец = ЦеныНоменклатурыСрезПоследних.Номенклатура
ГДЕ
	Штрихкоды.Штрихкод = &barcode
```

## API Endpoints

### Product Information
```
GET /products/{barcode}
```
Returns product information for a valid EAN13 barcode.

Example response:
```json
{
 "name": "Product Name",
 "measurement": "pcs",
 "price": 10.99,
 "discountPrice": 8.99
}
```

### Health Check
```
GET /api/health
```
Returns service health status including COM connector state.

Example response:
```json
{
    "status": "healthy",
    "comConnector": "connected"
}
```


## Development

1. Set `DEBUG_MODE = True` in service.py

2. Run directly with Python:
   ```bash
   python service.py
   ```


## Service Management

- Install: `PriceCheckerService.exe install`

- Start: `PriceCheckerService.exe start`

- Stop: `PriceCheckerService.exe stop`

- Remove: `PriceCheckerService.exe remove`  

- Status: `PriceCheckerService.exe status`

## Logging

Logs are stored in the `logs` directory next to the executable:

- Main log file: `logs/api_service.log`

- Daily rotation at midnight

- Automatic cleanup of logs older than 10 days

- Format: `api_service.log.YYYY-MM-DD`

## Error Handling

The service includes comprehensive error handling for:
- Invalid EAN13 barcodes
- COM connector issues
- Service initialization problems
- Database connection errors

## Security Notes

- The service requires administrative privileges for installation
- Ensure proper firewall configuration for port 8880
- Implement appropriate authentication if exposed beyond localhost
- Secure the COM connector access

## Troubleshooting

1. COM Connector Issues:
   - Verify `comcntr.dll` registration
   - Check 1C database path
   - Review logs for connection errors

2. Service Installation Problems:
   - Run command prompt as Administrator
   - Check Windows Event Viewer
   - Verify Python dependencies

3. Log Access Issues:

   - Ensure `logs` directory exists
   - Check permissions on logs directory
   - Verify service account permissions


## Notes

- Configuration files can be modified without rebuilding by placing them in the `settings` folder next to the executable

- Debug mode provides additional logging and direct console output

- Service runs on port 8880 by default

- Logs are stored in the `logs` directory next to the executable

- Daily rotation at midnight

- Automatic cleanup of logs older than 10 days

- Format: `api_service.log.YYYY-MM-DD`

## Contributing

[Your contribution guidelines]