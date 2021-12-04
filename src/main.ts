const query = ''
const sheetId = ''
const headers = ['id', 'date', 'from', 'to', 'cc', 'bcc', 'subject', 'plainBody', 'starred', 'unread']
const fetchSize = 500
const cellMaxLength = 50000
const format = 'yyyy/mm/dd hh:mm:ss'
const resumablesSheetName = 'resumables'

const salvage = () => {
    const sheet = SpreadsheetApp.openById(sheetId)
    const queryResultSheet = sheet.getSheetByName(query) || sheet.insertSheet(query)
    queryResultSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    const lastRow = queryResultSheet.getLastRow()
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
            body.length < cellMaxLength ? body : `${body.substring(0, cellMaxLength - 20)}...${cellMaxLength}字を超えたので省略`,
            message.isStarred() ? '⭐️' : '',
            message.isUnread() ? '✉️' : ''
        ])
    }))

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