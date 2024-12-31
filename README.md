
# Windows Service Product API with 1C COM Integration

This project implements a Windows service that exposes a REST API for product information retrieval from 1C database using COM connector. The service handles EAN13 barcode requests and returns product information in JSON format.

## Features

- Windows service implementation using Python
- REST API endpoint for product information retrieval
- EAN13 barcode validation
- 1C COM connector integration
- Daily log rotation with 10-day retention
- Health check endpoint
- Conversion to standalone .exe file

## Prerequisites

- Windows OS
- Python 3.7+
- Administrative privileges
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

3. Make sure 1C COM connector is properly registered:
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
APIService.exe install

# Start the service
APIService.exe start

# Stop the service
APIService.exe stop

# Remove the service
APIService.exe remove
```

## API Endpoints

### Product Information
```
GET /products/ean13/{barcode}
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

## Configuration

The service uses the following default settings:
- Port: 8880
- Logs directory: `C:\ServiceLogs`
- Log rotation: Daily
- Log retention: 10 days
- 1C database path: `C:\common\DB\83Empty8310`

## Logging

Logs are stored in `C:\ServiceLogs` with the following features:
- Daily rotation at midnight
- Automatic cleanup of logs older than 10 days
- Format: `api_service.log.YYYY-MM-DD`
- Current log: `api_service.log`

## Error Handling

The service includes comprehensive error handling for:
- Invalid EAN13 barcodes
- COM connector issues
- Service initialization problems
- Database connection errors

## Development

To modify the service:

1. Update the COM connector query in `ProductService.get_product_by_scan_code()`
2. Modify log settings in `setup_logging()`
3. Add new API endpoints in the Flask application
4. Rebuild the executable using `build_service.py`

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
   - Ensure `C:\ServiceLogs` directory exists
   - Check permissions on logs directory
   - Verify service account permissions

## License

[Your chosen license]

## Contributing

[Your contribution guidelines]