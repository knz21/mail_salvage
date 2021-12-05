const sheetId = ''
const headers = ['id', 'date', 'from', 'to', 'cc', 'bcc', 'subject', 'plainBody']
const defaultFetchSize = 500
const defaultMaxRowSize = 5000
const defaultConfig = [['query', ''], ['fetch_size', defaultFetchSize], ['max_row_size', defaultMaxRowSize]]
const maxCellTextLength = 50000
const format = 'yyyy/mm/dd hh:mm:ss'
const configSheetName = 'config'
const resumablesSheetName = 'resumables'

const salvage = () => {
    const sheet = SpreadsheetApp.openById(sheetId)
    const configSheet = sheet.getSheetByName(configSheetName)
    if (!configSheet) {
        Logger.log('The sheet "config" is necessary.')
        sheet.insertSheet(configSheetName).getRange(1, 1, defaultConfig.length, defaultConfig[0].length).setValues(defaultConfig)
        return
    }
    const configs = configSheet.getRange(1, 1, configSheet.getLastRow(), 2).getValues()
    const query = configs.find(row => row[0] == 'query')?.[1]
    if (!query) {
        Logger.log('The value "query" is necessary in config.')
        return
    }
    const fetchSize = configs.find(row => row[0] == 'fetch_size')?.[1] || defaultFetchSize
    const maxRowSize = configs.find(row => row[0] == 'max_row_size')?.[1] || defaultMaxRowSize

    let queryResultSheet = sheet.getSheetByName(query) || sheet.insertSheet(query)
    queryResultSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    let lastRow = queryResultSheet.getLastRow()
    const existsIds = queryResultSheet.getRange(2, 1, lastRow).getValues().map(row => row[0])

    const resumablesSheet = sheet.getSheetByName(resumablesSheetName) || sheet.insertSheet(resumablesSheetName)
    const lastResumableRow = resumablesSheet.getLastRow()
    const resumables = resumablesSheet.getRange(1, 1, lastResumableRow > 0 ? lastResumableRow : 1, 2).getValues()
    const startIndex = resumables.find(row => row[0] == query)?.[1] || 0
    const threads: GoogleAppsScript.Gmail.GmailThread[] = GmailApp.search(query, startIndex, fetchSize)
    Logger.log(`threads: ${threads.length}`)

    const messages: any[][] = []
    threads.forEach(thread => thread.getMessages().forEach((message: GoogleAppsScript.Gmail.GmailMessage) => {
        const id = message.getId()
        if (existsIds.includes(id)) return
        const body = message.getPlainBody()
        messages.push([
            message.getId(),
            message.getDate(),
            message.getFrom(),
            message.getTo(),
            message.getCc(),
            message.getBcc(),
            message.getSubject(),
            body.length < maxCellTextLength ? body : `${body.substring(0, maxCellTextLength - 50)}... Omitted because of max characters(${maxCellTextLength}).`
        ])
    }))

    if (lastRow > 1 && lastRow + messages.length > maxRowSize) {
        const numbers = sheet.getSheets()
            .map(sheet => {
                const name = sheet.getName()
                const found = name.match(/#[0-9]+?$/)
                if (name.startsWith(query) && found != null) {
                    return parseInt(found[0].substring(1))
                } else {
                    return 0
                }
            })
        const nextNumber = Math.max(...numbers) + 1
        queryResultSheet.setName(`${query}#${nextNumber}`)
        queryResultSheet = sheet.insertSheet(query)
        queryResultSheet.getRange(1, 1, 1, headers.length).setValues([headers])
        lastRow = 1
    }

    if (messages.length == 0) {
        Logger.log('No results')
    } else {
        Logger.log(`messages: ${messages.length}`)
        queryResultSheet.getRange(lastRow + 1, 1, messages.length, headers.length).setValues(messages)
        queryResultSheet.setRowHeightsForced(1, queryResultSheet.getLastRow(), 21)
        const formats = new Array(messages.length).fill([format])
        queryResultSheet.getRange(queryResultSheet.getLastRow() - messages.length + 1, 2, messages.length, 1).setNumberFormats(formats)
    }

    const resumableRow = resumables.findIndex(row => row[0] == query) + 1
    resumablesSheet.getRange(resumableRow > 0 ? resumableRow : 1, 1, 1, 2).setValues([[query, startIndex + threads.length]])
}