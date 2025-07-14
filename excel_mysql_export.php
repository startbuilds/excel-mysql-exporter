<?php
/**
 * Excel to MySQL Data Export System
 * 
 * A robust PHP solution for automated data export from Excel to MySQL
 * with duplicate detection and batch processing capabilities
 * 
 * @author Naveen Sharma
 * @version 1.0
 * @license MIT
 */

require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class ExcelToMySQLExporter {
    
    private $pdo;
    private $config;
    private $logFile;
    
    // Configuration settings
    private $defaultConfig = [
        'mysql' => [
            'host' => 'localhost',
            'port' => 3306,
            'database' => 'your_database',
            'username' => 'your_username',
            'password' => 'your_password',
            'charset' => 'utf8mb4'
        ],
        'tables' => [
            'primary' => 'main_data',
            'secondary' => 'supplementary_data'
        ],
        'sheets' => [
            'primary' => 'MainData',
            'secondary' => 'SupplementaryData'
        ],
        'batch_size' => 1000,
        'duplicate_check_columns' => ['id', 'name', 'date_created'],
        'log_file' => 'export_log.txt'
    ];
    
    public function __construct($config = []) {
        $this->config = array_merge($this->defaultConfig, $config);
        $this->logFile = $this->config['log_file'];
        $this->initializeDatabase();
        $this->log("Excel to MySQL Exporter initialized");
    }
    
    /**
     * Initialize database connection
     */
    private function initializeDatabase() {
        try {
            $dsn = sprintf(
                "mysql:host=%s;port=%d;dbname=%s;charset=%s",
                $this->config['mysql']['host'],
                $this->config['mysql']['port'],
                $this->config['mysql']['database'],
                $this->config['mysql']['charset']
            );
            
            $options = [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
                PDO::ATTR_EMULATE_PREPARES => false,
                PDO::MYSQL_ATTR_INIT_COMMAND => "SET sql_mode='STRICT_TRANS_TABLES,NO_ZERO_DATE,NO_ZERO_IN_DATE,ERROR_FOR_DIVISION_BY_ZERO'"
            ];
            
            $this->pdo = new PDO($dsn, $this->config['mysql']['username'], $this->config['mysql']['password'], $options);
            $this->log("Database connection established successfully");
            
        } catch (PDOException $e) {
            $this->log("Database connection failed: " . $e->getMessage(), 'ERROR');
            throw new Exception("Database connection failed: " . $e->getMessage());
        }
    }
    
    /**
     * Export all data from Excel file to MySQL
     */
    public function exportAllData($excelFilePath) {
        $startTime = microtime(true);
        $this->log("Starting full data export from: " . $excelFilePath);
        
        try {
            // Load Excel file
            $spreadsheet = IOFactory::load($excelFilePath);
            
            // Export primary sheet (900K records)
            $this->exportPrimarySheet($spreadsheet);
            
            // Export secondary sheet (17K records)
            $this->exportSecondarySheet($spreadsheet);
            
            $executionTime = microtime(true) - $startTime;
            $this->log("Full export completed successfully in " . round($executionTime, 2) . " seconds");
            
            return [
                'success' => true,
                'execution_time' => $executionTime,
                'message' => 'Data export completed successfully'
            ];
            
        } catch (Exception $e) {
            $this->log("Export failed: " . $e->getMessage(), 'ERROR');
            return [
                'success' => false,
                'error' => $e->getMessage()
            ];
        }
    }
    
    /**
     * Export incremental data (only new/updated records)
     */
    public function exportIncrementalData($excelFilePath) {
        $startTime = microtime(true);
        $this->log("Starting incremental data export");
        
        try {
            $spreadsheet = IOFactory::load($excelFilePath);
            $this->exportSecondarySheetIncremental($spreadsheet);
            
            $executionTime = microtime(true) - $startTime;
            $this->log("Incremental export completed in " . round($executionTime, 2) . " seconds");
            
            return [
                'success' => true,
                'execution_time' => $executionTime,
                'message' => 'Incremental export completed successfully'
            ];
            
        } catch (Exception $e) {
            $this->log("Incremental export failed: " . $e->getMessage(), 'ERROR');
            return [
                'success' => false,
                'error' => $e->getMessage()
            ];
        }
    }
    
    /**
     * Export primary sheet data with batch processing
     */
    private function exportPrimarySheet($spreadsheet) {
        $worksheet = $spreadsheet->getSheetByName($this->config['sheets']['primary']);
        
        if (!$worksheet) {
            throw new Exception("Primary sheet '{$this->config['sheets']['primary']}' not found");
        }
        
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        
        $this->log("Primary sheet has " . ($highestRow - 1) . " data rows");
        
        // Get column headers
        $headers = [];
        for ($col = 'A'; $col <= $highestColumn; $col++) {
            $headers[] = $worksheet->getCell($col . '1')->getCalculatedValue();
        }
        
        // Create table if not exists
        $this->createTableIfNotExists($this->config['tables']['primary'], $headers);
        
        // Process data in batches
        $batchSize = $this->config['batch_size'];
        $totalBatches = ceil(($highestRow - 1) / $batchSize);
        
        for ($batch = 0; $batch < $totalBatches; $batch++) {
            $startRow = ($batch * $batchSize) + 2; // Skip header row
            $endRow = min($startRow + $batchSize - 1, $highestRow);
            
            $this->processBatch($worksheet, $headers, $startRow, $endRow, $this->config['tables']['primary']);
            
            $this->log("Processed batch " . ($batch + 1) . "/$totalBatches (rows $startRow to $endRow)");
            
            // Memory management
            if ($batch % 10 === 0) {
                gc_collect_cycles();
            }
        }
    }
    
    /**
     * Export secondary sheet data
     */
    private function exportSecondarySheet($spreadsheet) {
        $worksheet = $spreadsheet->getSheetByName($this->config['sheets']['secondary']);
        
        if (!$worksheet) {
            throw new Exception("Secondary sheet '{$this->config['sheets']['secondary']}' not found");
        }
        
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        
        $this->log("Secondary sheet has " . ($highestRow - 1) . " data rows");
        
        // Get column headers
        $headers = [];
        for ($col = 'A'; $col <= $highestColumn; $col++) {
            $headers[] = $worksheet->getCell($col . '1')->getCalculatedValue();
        }
        
        // Create table if not exists
        $this->createTableIfNotExists($this->config['tables']['secondary'], $headers);
        
        // Process all data
        $this->processBatch($worksheet, $headers, 2, $highestRow, $this->config['tables']['secondary']);
    }
    
    /**
     * Export only new records from secondary sheet
     */
    private function exportSecondarySheetIncremental($spreadsheet) {
        $worksheet = $spreadsheet->getSheetByName($this->config['sheets']['secondary']);
        
        if (!$worksheet) {
            throw new Exception("Secondary sheet '{$this->config['sheets']['secondary']}' not found");
        }
        
        // Get last export timestamp
        $lastExportTime = $this->getLastExportTimestamp($this->config['tables']['secondary']);
        
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        
        // Get headers
        $headers = [];
        for ($col = 'A'; $col <= $highestColumn; $col++) {
            $headers[] = $worksheet->getCell($col . '1')->getCalculatedValue();
        }
        
        // Find date column (assuming there's a date column for tracking)
        $dateColumnIndex = array_search('date_created', $headers);
        if ($dateColumnIndex === false) {
            $dateColumnIndex = array_search('updated_at', $headers);
        }
        
        if ($dateColumnIndex !== false) {
            $newRecords = [];
            $dateColumn = chr(65 + $dateColumnIndex); // Convert to Excel column letter
            
            for ($row = 2; $row <= $highestRow; $row++) {
                $cellValue = $worksheet->getCell($dateColumn . $row)->getCalculatedValue();
                $recordDate = date('Y-m-d H:i:s', strtotime($cellValue));
                
                if ($recordDate > $lastExportTime) {
                    $newRecords[] = $row;
                }
            }
            
            $this->log("Found " . count($newRecords) . " new records to export");
            
            // Process new records
            foreach ($newRecords as $row) {
                $this->processBatch($worksheet, $headers, $row, $row, $this->config['tables']['secondary']);
            }
        } else {
            $this->log("No date column found, performing full export");
            $this->exportSecondarySheet($spreadsheet);
        }
        
        // Update last export timestamp
        $this->updateLastExportTimestamp($this->config['tables']['secondary']);
    }
    
    /**
     * Process a batch of rows
     */
    private function processBatch($worksheet, $headers, $startRow, $endRow, $tableName) {
        $data = [];
        
        for ($row = $startRow; $row <= $endRow; $row++) {
            $rowData = [];
            
            for ($colIndex = 0; $colIndex < count($headers); $colIndex++) {
                $column = chr(65 + $colIndex); // Convert to Excel column letter
                $cellValue = $worksheet->getCell($column . $row)->getCalculatedValue();
                $rowData[$headers[$colIndex]] = $cellValue;
            }
            
            $data[] = $rowData;
        }
        
        // Insert data with duplicate detection
        $this->insertDataWithDuplicateCheck($tableName, $data);
    }
    
    /**
     * Insert data with duplicate detection
     */
    private function insertDataWithDuplicateCheck($tableName, $data) {
        if (empty($data)) return;
        
        $duplicateCheckColumns = $this->config['duplicate_check_columns'];
        $insertedCount = 0;
        $duplicateCount = 0;
        
        foreach ($data as $row) {
            if ($this->isDuplicate($tableName, $row, $duplicateCheckColumns)) {
                $duplicateCount++;
                continue;
            }
            
            $this->insertRow($tableName, $row);
            $insertedCount++;
        }
        
        $this->log("Inserted: $insertedCount, Duplicates skipped: $duplicateCount");
    }
    
    /**
     * Check if record is duplicate
     */
    private function isDuplicate($tableName, $row, $checkColumns) {
        $whereConditions = [];
        $params = [];
        
        foreach ($checkColumns as $column) {
            if (isset($row[$column])) {
                $whereConditions[] = "$column = ?";
                $params[] = $row[$column];
            }
        }
        
        if (empty($whereConditions)) return false;
        
        $sql = "SELECT COUNT(*) FROM $tableName WHERE " . implode(' AND ', $whereConditions);
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute($params);
        
        return $stmt->fetchColumn() > 0;
    }
    
    /**
     * Insert a single row
     */
    private function insertRow($tableName, $row) {
        $columns = array_keys($row);
        $placeholders = array_fill(0, count($columns), '?');
        
        $sql = "INSERT INTO $tableName (" . implode(', ', $columns) . ") VALUES (" . implode(', ', $placeholders) . ")";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(array_values($row));
    }
    
    /**
     * Create table if not exists
     */
    private function createTableIfNotExists($tableName, $headers) {
        $columns = [];
        foreach ($headers as $header) {
            $columns[] = "`$header` TEXT";
        }
        
        $sql = "CREATE TABLE IF NOT EXISTS $tableName (
            id INT AUTO_INCREMENT PRIMARY KEY,
            " . implode(', ', $columns) . ",
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        )";
        
        $this->pdo->exec($sql);
        $this->log("Table $tableName created or verified");
    }
    
    /**
     * Get last export timestamp
     */
    private function getLastExportTimestamp($tableName) {
        $sql = "SELECT export_timestamp FROM export_log WHERE table_name = ? ORDER BY export_timestamp DESC LIMIT 1";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([$tableName]);
        
        $result = $stmt->fetchColumn();
        return $result ?: '1970-01-01 00:00:00';
    }
    
    /**
     * Update last export timestamp
     */
    private function updateLastExportTimestamp($tableName) {
        $sql = "INSERT INTO export_log (table_name, export_timestamp) VALUES (?, NOW())
                ON DUPLICATE KEY UPDATE export_timestamp = NOW()";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([$tableName]);
    }
    
    /**
     * Log messages
     */
    private function log($message, $level = 'INFO') {
        $timestamp = date('Y-m-d H:i:s');
        $logEntry = "[$timestamp] [$level] $message" . PHP_EOL;
        
        file_put_contents($this->logFile, $logEntry, FILE_APPEND | LOCK_EX);
        echo $logEntry;
    }
    
    /**
     * Get export statistics
     */
    public function getExportStats() {
        $stats = [];
        
        foreach ($this->config['tables'] as $key => $tableName) {
            $sql = "SELECT COUNT(*) as record_count FROM $tableName";
            $stmt = $this->pdo->prepare($sql);
            $stmt->execute();
            $stats[$key] = $stmt->fetch();
        }
        
        return $stats;
    }
}

// Usage example and CLI interface
if (php_sapi_name() === 'cli') {
    echo "Excel to MySQL Export Tool\n";
    echo "==========================\n\n";
    
    $options = getopt("f:t:h", ["file:", "type:", "help"]);
    
    if (isset($options['h']) || isset($options['help'])) {
        echo "Usage: php export.php -f <excel_file> -t <export_type>\n";
        echo "Options:\n";
        echo "  -f, --file     Excel file path\n";
        echo "  -t, --type     Export type (full|incremental)\n";
        echo "  -h, --help     Show this help message\n";
        exit(0);
    }
    
    $excelFile = $options['f'] ?? $options['file'] ?? null;
    $exportType = $options['t'] ?? $options['type'] ?? 'full';
    
    if (!$excelFile) {
        echo "Error: Excel file path is required\n";
        exit(1);
    }
    
    if (!file_exists($excelFile)) {
        echo "Error: Excel file not found: $excelFile\n";
        exit(1);
    }
    
    try {
        $exporter = new ExcelToMySQLExporter();
        
        if ($exportType === 'incremental') {
            $result = $exporter->exportIncrementalData($excelFile);
        } else {
            $result = $exporter->exportAllData($excelFile);
        }
        
        if ($result['success']) {
            echo "\nExport completed successfully!\n";
            echo "Execution time: " . round($result['execution_time'], 2) . " seconds\n";
            
            // Show statistics
            $stats = $exporter->getExportStats();
            echo "\nExport Statistics:\n";
            foreach ($stats as $table => $data) {
                echo "  $table: " . number_format($data['record_count']) . " records\n";
            }
        } else {
            echo "\nExport failed: " . $result['error'] . "\n";
            exit(1);
        }
        
    } catch (Exception $e) {
        echo "Error: " . $e->getMessage() . "\n";
        exit(1);
    }
}
?>