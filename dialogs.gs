function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Протоколы')
    .addItem('Создать протокол встречи', 'showMeetingTypeSelector')
    .addItem('Окно записей', 'showRecordDialog')
    .addSeparator()
    .addItem('Обновить кэш сотрудников', 'clearEmployeeCache') // Новая кнопка
    .addToUi();
}

function showMeetingDialog() {
  var html = HtmlService.createHtmlOutputFromFile('meetingForm')
    .setWidth(600)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Новая встреча');
}

function showRecordDialog(meetingId, meetingNumber) {
  var html = HtmlService.createHtmlOutputFromFile('recordForm')
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Протокол №' + (meetingNumber || ''));
  PropertiesService.getScriptProperties().setProperty('currentMeetingId', meetingId);
}

function showCalendarEventsModal() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('calendarEventsModal')
      .setWidth(500)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Выбор события из календаря');
  } catch(e) {
    console.error('Ошибка открытия окна:', e);
    throw new Error('Не удалось открыть окно календаря');
  }
}

function showMeetingTypeSelector() {
  var html = HtmlService.createHtmlOutputFromFile('meetingTypeSelector')
    .setWidth(650)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Создание встречи');
}