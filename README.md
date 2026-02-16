# ATT2000.MDB Export Utility

This utility exports attendance data from Microsoft Access ATT2000.MDB database to MariaDB/MySQL.

## Database Structure

### Source: ATT2000.MDB (Microsoft Access)

#### Tables:
- **userinfo** - User information table
  - `userid` - User ID (primary key)
  - `BadgeNumber` - Employee badge number
  - `att` - Attendance flag (1 = active for attendance)

- **CHECKINOUT** - Check-in/check-out records
  - `userid` - Foreign key to userinfo.userid
  - `CHECKTIME` - Date/time of check-in/out

### Target: MariaDB/MySQL

#### Table: `timecard`
```sql
CREATE TABLE timecard (
    scanCode VARCHAR(5) NOT NULL,    -- Badge number from userinfo
    scanAt DATETIME NOT NULL,        -- Check-in/out timestamp
    PRIMARY KEY (scanCode, scanAt)   -- Composite key with INSERT IGNORE
);
```

## Configuration

### Database Connection
- **Source**: Microsoft Access via OLEDB
- **Target**: MariaDB/MySQL with connection:
  - Host: Provided as parameter
  - Port: 3306
  - User: `pr-user`
  - Password: `pr-user`
  - Database: `payroll`
  - Timezone: `+07:00`

## Usage

### Prerequisites
1. **Node.js 12+** for basic compatibility
2. **Node.js 24+** for optimal performance (recommended)
3. Install dependencies: `npm install`
4. Compile TypeScript: `npm run build`

### Node.js 24+ Setup (Recommended)

For optimal performance with Node.js 24+, use the automated setup script:

#### Quick Start
```bash
# Navigate to Node.js 24 setup directory
cd node24

# Run the automated installation and build
./install.bat
```

#### Manual Setup for Node.js 24+
```bash
# Copy source files
cp ../attexport.ts .

# Install dependencies
npm run install

# Build with Node.js 24 optimizations
npm run build

# Run the application
npm run serve
```

### Legacy Node.js Setup

#### Node.js 24+ Automated Script

The `/node24/install.bat` script provides a complete automated setup:

**What it does:**
1. Copies `attexport.ts` from parent directory
2. Installs all npm dependencies
3. Builds TypeScript with Node.js 24 optimizations
4. Starts the application with default parameters

**Usage:**
```bash
# Navigate to setup directory
cd node24

# Run automated setup
./install.bat
```

**Custom Parameters:**
To use custom parameters, modify the `npm run serve` command in `package.json` or run manually:
```bash
node dist/attexport.js <mdb_path> <host_ip> [date]
```

#### Using Batch File (Windows)
```batch
attexport.bat <source_mdb_path> <target_host_ip> [optional_date]
```

Examples:
```batch
# Default export from last record in target
attexport.bat "D:\PAYROLL\ATT2000.MDB" "localhost"

# Export from specific date
attexport.bat "D:\PAYROLL\ATT2000.MDB" "192.168.1.100" "2024-01-01"
```

#### Using Node.js Directly
```bash
node dist/attexport.js <mdb_path> <host_ip> [date]
```

### Parameters
1. **mdb_path** - Full path to ATT2000.MDB file
2. **host_ip** - MariaDB/MySQL server IP address
3. **date** (optional) - Export from this date (YYYY-MM-DD format)

## Export Logic

### Data Selection
- Exports records from `userinfo` where `att = 1`
- Joins with `CHECKINOUT` table
- Filters by date range:
  - Start: Last record in target OR specified date OR 2023-01-01
  - End: Start date + 1 year
- Orders by `CHECKTIME` ascending

### Processing
- Filters badge numbers ≤ 5 characters
- Processes records sequentially by date
- Batches inserts by day for performance
- Uses `INSERT IGNORE` to prevent duplicates
- Shows progress with record counts and percentages

### Output Format
- Console progress tracking with timing
- Format: `DATE Records TOTAL %`
- Example: `2024-01-15     124     1500   8.3%`

## Error Handling

- Validates input parameters
- Checks for compiled JavaScript file
- Handles database connection errors
- Graceful cleanup of database connections

## Performance

- Connection pooling (max 5 connections)
- Batch processing by date
- Progress tracking with timing information
- Efficient memory usage with streaming processing
