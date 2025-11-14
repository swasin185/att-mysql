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

// convert to date "YYYY-MM-DD"
function formatDate(date: Date): string {
    return date.toLocaleDateString("sv-SE")
}

// convert to date "YYYY-MM-DD HH:MM:SS"
function formatDateTime(date: Date): string {
    return date.toLocaleString("sv-SE")
}

import ADODB from "node-adodb"

async function main() {
    let mariadbPool: mysql.Pool | null = null

    console.log("Starting KEEHIN ATT2000 Data Transfer...")

    try {
        // connect to ms-access DB
        const adodbConnection = ADODB.open(ADODB_CONNECTION_STRING)
        console.log(`Connected to MS Access MDB: ${MDB_FILE_PATH}`)

        // connect to mariadb DB
        mariadbPool = mysql.createPool(MARIADB_CONFIG)
        console.log(`Connected to MariaDB: ${MARIADB_CONFIG.database}`)

        // check last date data in target mariadb
        const dateQuery: string = "SELECT MAX(dateTxt) AS maxDate FROM timecard"
        const [result] = await mariadbPool.query<mysql.RowDataPacket[]>(dateQuery)

        // get last export date or 2023-01-01
        const maxDateStr = result[0].maxDate
        const exportDate: Date = maxDateStr ? new Date(maxDateStr) : new Date("2023-01-01")

        const exportDateStr = formatDate(exportDate)
        console.log(`Exporting records with CHECKTIME >= ${exportDateStr}`)

        const accessQuery: string = `
            SELECT
                userinfo.BadgeNumber,
                CHECKINOUT.CHECKTIME
            FROM
                userinfo, CHECKINOUT
            WHERE
                CHECKINOUT.CHECKTIME >= #${exportDateStr}# AND
                userinfo.att = 1 AND
                userinfo.userid = CHECKINOUT.userid`

        const checkInOutRecords: any[] = await adodbConnection.query(accessQuery)
        console.log("MS-Access records", checkInOutRecords.length)
        let insertCount = 0
        let batch: [string, string, string][] = []
        const BATCH_SIZE = 1000
        const TIME_ZONE_OFFSET = 7 * 60 * 60 * 1000

        for (const record of checkInOutRecords)
            if (record.BadgeNumber <= "99999") {
                const badgeNumber: string = record.BadgeNumber
                const localTime = record.CHECKTIME + TIME_ZONE_OFFSET
                const checkTime: string = formatDateTime(localTime)
                const dateTxt = checkTime.substring(0, 10)
                const timeTxt = checkTime.substring(11, 19)
                batch.push([dateTxt, badgeNumber, timeTxt])
                if (batch.length >= BATCH_SIZE) {
                    await insertBatch(mariadbPool, batch)
                    insertCount += batch.length
                    batch = []
                }
            }

        if (batch.length > 0) {
            await insertBatch(mariadbPool, batch)
            insertCount += batch.length
        }

        console.log(`✅ Time Export **${insertCount}** records successfully.`)
    } catch (e) {
        console.error("An error occurred during data transfer:", (e as Error).message)
    } finally {
        if (mariadbPool) await mariadbPool.end()
        console.log("Database connections closed.")
    }
}

async function insertBatch(conn: mysql.Pool, batch: [string, string, string][]) {
    try {
        const params = batch.flat()
        const placeholders = new Array(batch.length).fill("(?, ?, ?)").join(", ")
        const sql = `
            INSERT INTO timecard (dateTxt, scanCode, timeTxt)
            VALUES ${placeholders}
            ON DUPLICATE KEY UPDATE timeTxt = VALUES(timeTxt)`
        await conn.execute(sql, params)
    } catch (e) {
        console.error("Batch insert failed (some records might be duplicates):", e)
    }
}

main().catch(() => {
    console.log("finish!")
})
