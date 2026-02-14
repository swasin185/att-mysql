import * as mysql from "mysql2/promise"

// Run command "node ./dist/attexport.js path host-ip"
// Path to MDB in first first parameter
const MDB_FILE_PATH: string = process.argv[2] || "D:\\PAYROLL\\ATT2000.MDB"

// builin ms-access connect string
const ADODB_CONNECTION_STRING: string = `Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${MDB_FILE_PATH};`

// mariadb connect config get host-ip from second parameter
const MARIADB_CONFIG: mysql.PoolOptions = {
    host: process.argv[3] || "localhost",
    port: 3306,
    user: "pr-user",
    password: "pr-user",
    database: "payroll",
    connectionLimit: 5,
    waitForConnections: true,
    queueLimit: 0,
    timezone: "+07:00",
}

import ADODB from "node-adodb"
import moment from "moment"
var insertCount = 0
var batch: [string, string][] = []

// Progress tracking variables
var totalRecords = 0
var processedRecords = 0
var currentDay = ""
var dayRecords = 0
var startTime = Date.now()

// Function to format aligned console output
function logProgress(elapsed: string, date: string, transfered: number, total: number, percent: string) {
    const elapsedStr = elapsed.padEnd(8)
    const dateStr = date.padEnd(12)
    const transferedStr = transfered.toString().padStart(8)
    const totalStr = total.toString().padStart(8)
    const percentStr = percent.padStart(7)
    
    console.info(`${elapsedStr} ${dateStr} ${transferedStr} ${totalStr} ${percentStr}`)
}

// Function to format elapsed time
function formatElapsed(ms: number): string {
    const seconds = Math.floor(ms / 1000)
    const minutes = Math.floor(seconds / 60)
    const hours = Math.floor(minutes / 60)
    
    if (hours > 0) {
        return `${hours}h${minutes % 60}m`
    } else if (minutes > 0) {
        return `${minutes}m${seconds % 60}s`
    } else {
        return `${seconds}s`
    }
}

var paramDate = process.argv[4]

async function main() {
    console.time("Import")
    var mariadbPool: mysql.Pool | null = null
    try {
        // connect to ms-access DB
        const adodbConnection = ADODB.open(ADODB_CONNECTION_STRING)
        
        // connect to mariadb DB
        mariadbPool = mysql.createPool(MARIADB_CONFIG)
        
        // check last date data in target mariadb
        const dateQuery: string = "SELECT MAX(scanAt) AS maxDate FROM timecard"
        const [result] = await mariadbPool.query<mysql.RowDataPacket[]>(dateQuery)

        const maxDateStr = result[0].maxDate
        const exportDate: Date = paramDate ? new Date(paramDate) : maxDateStr ? new Date(maxDateStr) : new Date("2023-01-01")

        const exportDateStr = moment(exportDate).format("YYYY-MM-DD")
        const untilDateStr = moment(exportDate).add(1, "year").format("YYYY-MM-DD")

        const accessQuery: string = `
            SELECT
                userinfo.BadgeNumber,
                CHECKINOUT.CHECKTIME
            FROM
                userinfo, CHECKINOUT
            WHERE
                CHECKINOUT.CHECKTIME >= #${exportDateStr}# AND
                CHECKINOUT.CHECKTIME <= #${untilDateStr}# AND
                userinfo.att = 1 AND
                userinfo.userid = CHECKINOUT.userid
            ORDER BY CHECKINOUT.CHECKTIME`

        console.time("MS-Access Query")
        const checkInOutRecords: any[] = await adodbConnection.query(accessQuery)
        console.timeEnd("MS-Access Query")
        
        // Initialize progress tracking
        totalRecords = checkInOutRecords.length
        processedRecords = 0
        currentDay = ""
        dayRecords = 0
        startTime = Date.now()
        
        // Print header
        console.info("")
        console.info("ELAPSE   DATE         TRANSFERED    TOTAL   PERCENT")
        console.info("-------- ------------ -------- -------- -------")
        
        console.time("MariaDB Insertion")
        batch = []
        
        // Process records sequentially (already sorted by day)
        for (const record of checkInOutRecords) {
            if (record.BadgeNumber.length <= 5) {
                const iso = new Date(record.CHECKTIME)
                const checkTime = moment(iso)
                const timeTxt = checkTime.format("YYYY-MM-DD HH:mm")
                const recordDate = checkTime.format("YYYY-MM-DD")
                
                // Check if we've moved to a new day
                if (currentDay !== "" && currentDay !== recordDate) {
                    // Insert all records for previous day
                    if (batch.length > 0) {
                        await insertBatch(mariadbPool)
                    }
                    
                    // Log progress for completed day
                    const elapsed = formatElapsed(Date.now() - startTime)
                    const percent = ((processedRecords / totalRecords) * 100).toFixed(1)
                    logProgress(elapsed, currentDay, dayRecords, processedRecords, percent + "%")
                    dayRecords = 0
                    batch = []
                }
                
                currentDay = recordDate
                dayRecords++
                processedRecords++

                batch.push([record.BadgeNumber, timeTxt])
            }
        }
        
        // Insert and log final day
        if (batch.length > 0) {
            await insertBatch(mariadbPool)
            const elapsed = formatElapsed(Date.now() - startTime)
            const percent = ((processedRecords / totalRecords) * 100).toFixed(1)
            logProgress(elapsed, currentDay, dayRecords, processedRecords, percent + "%")
        }
            
        console.timeEnd("MariaDB Insertion")
        console.timeEnd("Import")
        console.info("")
        console.info(`✅ Export completed: ${insertCount} records transferred successfully.`)
        console.info("")
    } catch (e) {
        console.error("Error:", (e as Error).message)
    } finally {
        if (mariadbPool) await mariadbPool.end()
    }
}

async function insertBatch(conn: mysql.Pool) {
    const params = batch.flat()
    const placeholders = new Array(batch.length).fill("(?, ?)").join(", ")
    const sql = `
        INSERT IGNORE INTO timecard (scanCode, scanAt)
        VALUES ${placeholders}`
    await conn.execute(sql, params)
    insertCount += batch.length
    batch = []
}

main().catch(console.error)
