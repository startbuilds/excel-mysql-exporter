# Excel to MySQL Data Export Tool

A professional PHP solution for automated data export from Excel to MySQL with advanced duplicate detection, batch processing, and incremental updates.

## Features

✅ **Large Dataset Support** - Handles 900K+ records efficiently  
✅ **Duplicate Detection** - Prevents redundant data uploads  
✅ **Batch Processing** - Memory-efficient processing of large datasets  
✅ **Incremental Updates** - Export only new/updated records  
✅ **Error Handling** - Comprehensive error logging and recovery  
✅ **Performance Optimized** - Built for speed and reliability  
✅ **CLI Interface** - Easy automation and scheduling  
✅ **Detailed Logging** - Full audit trail of export operations  

## Requirements

- PHP 7.4 or higher
- MySQL 5.7 or higher
- Composer
- PHP Extensions: PDO, PDO_MySQL
- Memory limit: 512MB+ recommended for large datasets

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-company/excel-mysql-exporter.git
cd excel-mysql-exporter
```

2. Install dependencies:
```bash
composer install
```

3. Configure database settings in the script or create a config file

## Usage

### Command Line Interface

**Full Export (All Data):**
```bash
php export.php -f your_excel_file.xlsx -t full
```

**Incremental Export (New Records Only):**
```bash
php export.php -f your_excel_file.xlsx -t incremental
```

**Help:**
```bash
php export.php -h
```

### Programmatic Usage

```php
require_once 'vendor/autoload.php';

$config = [
    'mysql' => [
        'host' => 'localhost',
        'database' => 'your_database',
        'username' => 'your_username',
        'password' => 'your_password'
    ],
    'batch_size' => 1000,
    'duplicate_check_columns' => ['id', 'name', 'email']
];

$exporter = new ExcelToMySQLExporter($config);

// Full export
$result = $exporter->exportAllData('data.xlsx');

// Incremental export
$result = $exporter->exportIncrementalData('data.xlsx');

if ($result['success']) {
    echo "Export completed in " . $result['execution_time'] . " seconds";
} else {
    echo "Export failed: " . $result['error'];
}
```

## Configuration

### Database Configuration
```php
'mysql' => [
    'host' => 'localhost',
    'port' => 3306,
    'database' => 'your_database',
    'username' => 'your_username',
    'password' => 'your_password',
    'charset' => 'utf8mb4'
]
```

### Table and Sheet Mapping
```php
'tables' => [
    'primary' => 'main_data',
    'secondary' => 'supplementary_data'
],
'sheets' => [
    'primary' => 'MainData',
    'secondary' => 'SupplementaryData'
]
```

### Duplicate Detection
```php
'duplicate_check_columns' => ['id', 'name', 'date_created']
```

## Performance Optimization

- **Batch Processing**: Processes data in configurable batches (default: 1000 records)
- **Memory Management**: Automatic garbage collection for large datasets
- **Prepared Statements**: Prevents SQL injection and improves performance
- **Connection Pooling**: Efficient database connection management
- **Indexing**: Automatic table creation with proper indexes

## Error Handling

- Comprehensive error logging
- Graceful failure recovery
- Detailed execution reports
- Memory usage monitoring
- Progress tracking for large datasets

## Logging

All operations are logged with timestamps and severity levels:
- Export start/completion times
- Batch processing progress
- Duplicate detection results
- Error messages and stack traces
- Performance metrics

## Database Schema

The tool automatically creates tables with the following structure:
```sql
CREATE TABLE IF NOT EXISTS table_name (
    id INT AUTO_INCREMENT PRIMARY KEY,
    -- Dynamic columns based on Excel headers --
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);
```

## Monitoring and Statistics

Get export statistics:
```php
$stats = $exporter->getExportStats();
foreach ($stats as $table => $data) {
    echo "$table: " . number_format($data['record_count']) . " records\n";
}
```

## Scheduling

Set up automated exports using cron:
```bash
# Full export every night at 2 AM
0 2 * * * /usr/bin/php /path/to/export.php -f /path/to/data.xlsx -t full

# Incremental export every hour
0 * * * * /usr/bin/php /path/to/export.php -f /path/to/data.xlsx -t incremental
```

## Security Features

- Prepared statements prevent SQL injection
- Input validation and sanitization
- Secure connection handling
- Error message sanitization
- Configurable access controls

## Testing

Run the test suite:
```bash
composer test
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For technical support and questions:
- Email: dev@yourcompany.com
- Issues: GitHub Issues
- Documentation: Full API documentation available

## Changelog

### v1.0.0
- Initial release
- Full and incremental export functionality
- Duplicate detection
- Batch processing
- CLI interface
- Comprehensive logging

---

**Built with ❤️ by Naveen**

*The solution is derived from a proven script already handling 900K+ records reliably in production.*