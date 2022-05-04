// TODO: sort short tasks in one date
// TODO: check minutes and seconds

import xlsx from 'node-xlsx';
import { createObjectCsvWriter } from 'csv-writer';

import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const FILENAME = 'timesheet.xlsx';
const YANDEX_WORKSHEET = 'Табель (по тикетам)';
const LAST_TEXT = 'Итого, ч';

const csvWriter = createObjectCsvWriter({
  path: 'toggl.csv',
  header: [
    { id: 'User', title: 'User' },
    { id: 'Email', title: 'Email' },
    { id: 'Client', title: 'Client' },
    { id: 'Project', title: 'Project' },
    { id: 'Description', title: 'Description' },
    { id: 'Start date', title: 'Start date' },
    { id: 'Start time', title: 'Start time' },
    { id: 'Duration', title: 'Duration' },
    { id: 'Tags', title: 'Tags' },
  ],
});

const zeroPad = (num, places) => String(num).padStart(places, '0');

const createTaskRow = ({ taskName, taskDuration, taskStartDate }) => {
  const row = {
    User: 'Никита Парфенов',
    Email: 'n1koss.boy@gmail.com',
    Project: 'Yandex.Plus',
    Description: taskName,
    'Start date': taskStartDate.toISOString().split('T')[0], // TODO YYYY-MM-DD
    'Start time': '11:00:00', // TODO HH:MM:SS
    Duration: `${zeroPad(taskDuration, 2)}:00:00`,
  };

  return row;
};

const getYandexSheet = (exportedYandexSheet) => {
  return exportedYandexSheet.find(
    ({ name }) => name === YANDEX_WORKSHEET
  ).data;
};

const excelDateToDate = (excelSerialDate) => {
  return new Date(Date.UTC(0, 0, excelSerialDate - 1));
};

const getLastRow = (sheet) => {
  return sheet.findIndex((row) => row[0] === LAST_TEXT);
};

const getLastColumn = (sheet, row) => {
  return sheet[row].findIndex((column) => column === LAST_TEXT);
};

const exportedYandexSheet = xlsx.parse(`${__dirname}/${FILENAME}`);
const yandexSheet = getYandexSheet(exportedYandexSheet);

const NAME_COLUMN = 3;
const KEY_COLUMN = 2;
const START_DURATION_COLUMN = 7;
const START_ROW = 4;
const START_DATE_ROW = START_ROW - 2;
const START_DAY_ROW = START_DATE_ROW - 1;
const LAST_ROW = getLastRow(yandexSheet);
const LAST_DURATION_COLUMN = getLastColumn(
  yandexSheet,
  START_DAY_ROW
);

const records = [];
for (
  let durationColumn = START_DURATION_COLUMN;
  durationColumn < LAST_DURATION_COLUMN;
  durationColumn++
) {
  let prevTaskName = null;

  for (let row = START_ROW; row < LAST_ROW; row++) {
    const key = yandexSheet[row][KEY_COLUMN];
    const name = yandexSheet[row][NAME_COLUMN];

    let taskName = key && name ? `${key}: ${name}` : prevTaskName;
    prevTaskName = taskName;

    const date = yandexSheet[START_DATE_ROW][durationColumn];
    const taskStartDate = excelDateToDate(date);

    const taskDuration = yandexSheet[row][durationColumn];
    if (!taskDuration) continue;

    const taskRow = createTaskRow({
      taskName,
      taskDuration,
      taskStartDate,
    });
    records.push(taskRow);
  }
}

csvWriter.writeRecords(records).then(() => {
  console.log('Done!');
});
