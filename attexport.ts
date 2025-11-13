import * as mysql from "mysql2/promise"

const MDB_FILE_PATH: string = process.argv[2] || "D:\\PAYROLL\\ATT2000.MDB"

const ADODB_CONNECTION_STRING: string = `Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${MDB_FILE_PATH};`

const MARIADB_CONFIG: mysql.PoolOptions = {
    host: "192.168.1.9",
    user: "pr-user",
    password: "pr-user",
    database: "payroll",
    port: 3306,
    connectionLimit: 5,
}

function formatDate(date: Date): string {
    return date.toISOString().substring(0, 10) // YYYY-MM-DD
}

function subtractMonths(date: Date, months: number): Date {
    const d = new Date(date)
    d.setMonth(d.getMonth() - months)
    return d
}

const ADODB = await import("node-adodb")

async function main() {
    let mariadbPool: mysql.Pool | null = null
    let mariadbConn: mysql.PoolConnection | null = null

    console.log("Starting KEEHIN ATT2000 Data Transfer...")

    try {
        const adodbConnection = ADODB.open(ADODB_CONNECTION_STRING)
        console.log(`Connected to MS Access MDB: ${MDB_FILE_PATH}`)

        mariadbPool = mysql.createPool(MARIADB_CONFIG)
        mariadbConn = await mariadbPool.getConnection()
        console.log(`Connected to MariaDB: ${MARIADB_CONFIG.database}`)

        const dateQuery: string =
            "SELECT MAX(dateTxt) AS maxDate FROM timecard WHERE dateTxt <= CURDATE()"
        const result = await mariadbConn.query<mysql.RowDataPacket[]>(dateQuery)

        let exportDate: Date
        const maxDateStr = (result[0] as any).maxDate

        if (maxDateStr) {
            exportDate = new Date(maxDateStr)
        } else {
            exportDate = subtractMonths(new Date(), 12)
        }

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

        await mariadbConn.beginTransaction()

        let insertCount = 0
        let batch: [string, string, string][] = []
        const BATCH_SIZE = 100

        for (const record of checkInOutRecords) {
            const badgeNumber: string = record.BadgeNumber
            const checkTime: string = record.CHECKTIME.toString() // Format: YYYY-MM-DD HH:mm:ss

            const dateTxt = checkTime.substring(0, 10)
            const timeTxt = checkTime.substring(11, 19)

            batch.push([badgeNumber, dateTxt, timeTxt])

            if (batch.length >= BATCH_SIZE) {
                await insertBatch(mariadbConn, batch)
                insertCount += batch.length
                batch = []
            }
        }

        if (batch.length > 0) {
            await insertBatch(mariadbConn, batch)
            insertCount += batch.length
        }

        await mariadbConn.commit()
        console.log(`✅ Time Export **${insertCount}** records successfully.`)
    } catch (e) {
        if (mariadbConn) {
            try {
                await mariadbConn.rollback()
                console.log("Transaction rolled back due to error.")
            } catch (rbError) {
                console.error("Rollback failed:", rbError)
            }
        }
        console.error("An error occurred during data transfer:", (e as Error).message)
    } finally {
        if (mariadbConn) mariadbConn.release()
        if (mariadbPool) await mariadbPool.end()

        console.log("Database connections closed.")
    }
}

async function insertBatch(conn: mysql.PoolConnection, batch: [string, string, string][]) {
    try {
        const params = batch.flat()
        const placeholders = new Array(batch.length).fill("(?, ?, ?)").join(", ")
        const sql = `
          INSERT INTO timecard (BadgeNumber, dateTxt, TimeTxt)
          VALUES ${placeholders}
          ON DUPLICATE KEY UPDATE BadgeNumber = VALUES(BadgeNumber)`
        await conn.execute(sql, params)
    } catch (e) {
        console.error("Batch insert failed (some records might be duplicates):", e)
    }
}

main().catch(() => {
    console.log("finish!")
})
