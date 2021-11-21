const query = ''
const sheetId = ''
const headers = ['id', 'date', 'from', 'to', 'cc', 'bcc', 'subject', 'plainBody', 'starred', 'unread']
const fetchSize = 500

const salvage = () => {
    const sheet = SpreadsheetApp.openById(sheetId)
    const queryResultSheet = sheet.getSheetByName(query) || sheet.insertSheet(query)
    queryResultSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    const lastRow = queryResultSheet.getLastRow()
    const existsIds = queryResultSheet.getRange(2, 1, lastRow).getValues().map(row => row[0])

    const threads: GoogleAppsScript.Gmail.GmailThread[] = GmailApp.search(query, lastRow - 1, fetchSize)

    Logger.log(`threads: ${threads.length}`)

    const messages: any[][] = []
    threads.forEach(thread => thread.getMessages().forEach((message: GoogleAppsScript.Gmail.GmailMessage) => {
        const id = message.getId()
        if (existsIds.includes(id)) return
        messages.push([
            message.getId(),
            message.getDate(),
            message.getFrom(),
            message.getTo(),
            message.getCc(),
            message.getBcc(),
            message.getSubject(),
            message.getPlainBody(),
            message.isStarred() ? '⭐️' : '',
            message.isUnread() ? '✉️' : ''
        ])
    }))

    if (messages.length == 0) {
        Logger.log('No results')
        return
    }

    Logger.log(`messages: ${messages.length}`)
    queryResultSheet.getRange(lastRow + 1, 1, messages.length, headers.length).setValues(messages)
}