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
import moment from "moment";
var insertCount = 0
var batch: [string, string][] = []
const BATCH_SIZE = 1000

const paramDate = process.argv[4]

async function main() {
    console.time("Total Execution")
    let mariadbPool: mysql.Pool | null = null

    console.log("Starting KEEHIN ATT2000 Data Transfer...")

    try {
        // connect to ms-access DB
        const adodbConnection = ADODB.open(ADODB_CONNECTION_STRING)
        console.log(`Connected to MS Access MDB: ${MDB_FILE_PATH}`)

        // connect to mariadb DB
        mariadbPool = mysql.createPool(MARIADB_CONFIG)
        console.log(`Connected to MariaDB: ${MARIADB_CONFIG.database}`)
        console.timeLog("Total Execution", "Connected to DBs")

        // check last date data in target mariadb
        const dateQuery: string = "SELECT MAX(scanAt) AS maxDate FROM timecard"
        const [result] = await mariadbPool.query<mysql.RowDataPacket[]>(dateQuery)

        const maxDateStr = result[0].maxDate
        const exportDate: Date = paramDate ? new Date(paramDate) : maxDateStr ? new Date(maxDateStr) : new Date("2023-01-01")

        const exportDateStr = moment(exportDate).format("YYYY-MM-DD")
        const untilDateStr = moment(exportDate).add(1, "month").format("YYYY-MM-DD")
        console.log(`Exporting records with CHECKTIME >= ${exportDateStr}`)
        console.timeLog("Total Execution", "Retrieved Max Date")

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


        console.time("Access Query")
        const checkInOutRecords: any[] = await adodbConnection.query(accessQuery)
        console.timeEnd("Access Query")
        console.timeLog("Total Execution", "Access Query Completed")
        console.log("MS-Access records", checkInOutRecords.length)

        console.time("Batch Insertion")

        for (const record of checkInOutRecords)
            if (record.BadgeNumber.length <= 5) {
                const badgeNumber: string = record.BadgeNumber
                const iso = new Date(record.CHECKTIME)
                const checkTime = moment(iso)
                // const dateTxt = checkTime.format("YYYY-MM-DD")
                const timeTxt = checkTime.format("YYYY-MM-DD HH:mm")
                batch.push([badgeNumber, timeTxt])
                if (batch.length >= BATCH_SIZE) {
                    await insertBatch(mariadbPool, batch)
                }
            }

        if (batch.length > 0) {
            await insertBatch(mariadbPool, batch)
        }
        console.timeEnd("Batch Insertion")
        console.timeLog("Total Execution", "Batch Insertion Completed")

        console.log(`✅ Time Export **${insertCount}** records successfully.`)
        console.timeEnd("Total Execution")
    } catch (e) {
        console.error("An error occurred during data transfer:", (e as Error).message)
    } finally {
        if (mariadbPool) await mariadbPool.end()
        console.log("Database connections closed.")
    }
}

async function insertBatch(conn: mysql.Pool, batch: [string, string][]) {
    const params = batch.flat()
    const placeholders = new Array(batch.length).fill("(?, ?)").join(", ")
    const sql = `
        INSERT INTO timecard (scanCode, scanAt)
        VALUES ${placeholders}
        ON DUPLICATE KEY UPDATE scanAt = VALUES(scanAt)`
    await conn.execute(sql, params)
    insertCount += batch.length
    batch = []
}

await main()
