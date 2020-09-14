const onHomepage = (e) => {
  return createCard()
}

const createCard = (sheetUrl) => {
  const startTimePicker = CardService.newDateTimePicker()
    .setTitle('start time')
    .setFieldName('start')
    .setValueInMsSinceEpoch(new Date().getTime() - 30 * 24 * 60 * 60 * 1000)

  const endTimePicker = CardService.newDateTimePicker()
    .setTitle('end time')
    .setFieldName('end')
    .setValueInMsSinceEpoch(new Date().getTime())

  const action = CardService.newAction().setFunctionName('handleExport')

  const button = CardService.newTextButton()
    .setText('Export')
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.TEXT)

  const buttonSet = CardService.newButtonSet().addButton(button)

  const section = CardService.newCardSection()
    .addWidget(startTimePicker)
    .addWidget(endTimePicker)
    .addWidget(buttonSet)

  const card = CardService.newCardBuilder().addSection(section)

  if (sheetUrl) {
    const footer = CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Open Spreadsheet')
        .setOpenLink(CardService.newOpenLink().setUrl(sheetUrl))
    )
    card.setFixedFooter(footer)
  }

  return card.build()
}

const handleExport = (e) => {
  const { start, end } = e.formInput
  const startTime = new Date(start.msSinceEpoch)
  const endTime = new Date(end.msSinceEpoch)
  const events = CalendarApp.getDefaultCalendar()
    .getEvents(startTime, endTime)
    .map((e) => [
      e.getId(),
      e.getTitle(),
      e.getDescription(),
      e.getStartTime(),
      e.getEndTime(),
      e.getLocation(),
      e.getCreators().join(','),
      e.getGuestList().length,
    ])
  const keys = [
    'id',
    'title',
    'description',
    'startTime',
    'endTime',
    'location',
    'creator',
    'guests',
  ]
  const values = [keys].concat(events)

  const spreadsheet = SpreadsheetApp.create('Calendar Event List')
  const sheet = spreadsheet.getSheets()[0]
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values)
  const sheetUrl = spreadsheet.getUrl()

  const card = createCard(sheetUrl)
  const navigation = CardService.newNavigation().updateCard(card)
  const actionResponse = CardService.newActionResponseBuilder().setNavigation(navigation)
  return actionResponse.build()
}
