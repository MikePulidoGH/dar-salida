function actualizarDiaEnRemitos() {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sheets.forEach(sheet => {
        if (sheet.getName().includes('Remito')) {
            // Update D1 with the current date
            sheet.getRange('D1').setValue(formattedDate);

            // Clear any dropdown validation in the sheet
            const range = sheet.getDataRange();
            range.clearDataValidations();

            // Maintaining daily updates could be done with a time-driven trigger.
        }
    });
}