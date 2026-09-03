const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const functions = [
  'getAvailableShifts', 'normalizeImportedShift', 'getShiftForDay',
  'resetImportedShiftOverrides', 'normalizeHeader', 'pick', 'excelDateToDate',
  'formatDate', 'monthKey', 'timeText', 'durationToHours', 'parseClockMinutes',
  'estimateHours', 'calculateShiftTotalHours', 'hasMissingEndPunch',
  'applyOvertimeRule', 'parseOvertime', 'normalizeRows', 'datesInMonth',
  'getReport', 'shiftHours', 'calculateShiftOvertime', 'formatHours',
  'formatDayTotalHours', 'getReportDayStatus', 'isCalendarRestDay',
  'withHolidayFill', 'thinBorder', 'formatClockMinutes', 'getInspectionStartTime',
  'getInspectionStatus', 'getAdjustedInspectionEnd', 'loadWorkbook'
];

function harness() {
  const callbacks = {};
  const exports = [];
  const defaults = [
    { name: 'A班', start: '08:00', end: '16:00' },
    { name: '正常班', start: '08:00', end: '17:30' }
  ];
  const context = vm.createContext({
    Date, Blob, console,
    defaultShifts: defaults, shifts: defaults.map((s) => ({ ...s })),
    shiftOverrides: {}, rawRows: [], selectedEmployeeKey: '',
    monthSelect: { value: '2026-09' }, overtimeLimitInput: { value: 46 },
    floorOvertimeCheckbox: { checked: true }, weekdays: ['日', '一', '二', '三', '四', '五', '六'],
    getHolidayForDate: () => null, updateMonthOptions() {}, render() {},
    localStorage: { setItem() {} },
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL() {} },
    document: { body: { appendChild() {}, removeChild() {} }, createElement: () => ({ click() {} }) },
    exportButton: { addEventListener: (_, cb) => { callbacks.B = cb; } },
    complianceButton: { addEventListener: (_, cb) => { callbacks.A = cb; } },
    XLSX: {
      utils: {
        sheet_to_json: (sheet, options) => options.header === 1 ? [Object.keys(sheet[0])] : sheet,
        book_new: () => ({ sheets: [] }),
        aoa_to_sheet: (rows) => {
          const sheet = {};
          rows.forEach((row, r) => row.forEach((v, c) => { sheet[`${r}:${c}`] = { v }; }));
          return sheet;
        },
        decode_cell: (address) => {
          const [r, c] = address.split(':').map(Number);
          return { r, c };
        },
        book_append_sheet: (book, sheet) => book.sheets.push(sheet)
      },
      write: (book) => { exports.push(book); return ''; }
    }
  });
  for (const name of functions) {
    const start = html.indexOf(`    function ${name}(`);
    assert.notEqual(start, -1, name);
    const end = html.indexOf('\n    }', start) + '\n    }'.length;
    vm.runInContext(html.slice(start, end), context);
  }
  for (const [startMarker, endMarker] of [
    ['    exportButton.addEventListener("click"', '    function formatClockMinutes'],
    ['    complianceButton.addEventListener("click"', '    function thinBorder']
  ]) {
    const start = html.indexOf(startMarker);
    vm.runInContext(html.slice(start, html.indexOf(endMarker, start)), context);
  }
  return { context, callbacks, exports };
}

function input(day, period) {
  const row = { 姓名: '測試員工', 工號: 'T1', 日期: new Date(2026, 8, day), 上班1: '08:00', 下班1: '18:00' };
  if (period !== undefined) row[' 對應時段 '] = period;
  return row;
}

test('import mapping drives report calculations and both export shift columns', () => {
  const { context: c, callbacks, exports } = harness();
  const inputs = [input(1, '正常班'), input(2, ' A班 '), input(3, '其他'), input(4, ''), input(5)];
  const expected = ['正常班', 'A班', '正常班', '正常班', '正常班'];
  c.loadWorkbook({ SheetNames: ['data'], Sheets: { data: inputs } });
  assert.deepEqual(Array.from(c.rawRows, (r) => r.shiftName), expected);
  const days = c.getReport().people[0].days;
  assert.deepEqual(Array.from(days.slice(0, 5), (d) => d.shift.name), expected);
  assert.equal(days[0].overtimeHours, 0);
  assert.equal(days[1].overtimeHours, 2);
  assert.equal(days[5].shift.name, '正常班');
  callbacks.B();
  callbacks.A();
  for (const book of exports) {
    const sheet = book.sheets[0];
    assert.equal(sheet['1:2'].v, '班別');
    expected.forEach((name, i) => assert.equal(sheet[`${i + 2}:2`].v, name));
  }
});

test('re-import replaces affected month overrides but permits later manual changes', () => {
  const { context: c } = harness();
  c.shiftOverrides = {
    'T1|測試員工|2026/9/1': 'A班', 'T1|測試員工|2026/9/30': 'A班',
    'T1|測試員工|2026/8/1': 'A班', 'other|2026/9/1': 'A班'
  };
  const book = { SheetNames: ['data'], Sheets: { data: [input(1, '正常班')] } };
  c.loadWorkbook(book);
  assert.equal(c.shiftOverrides['T1|測試員工|2026/9/30'], undefined);
  assert.equal(c.shiftOverrides['T1|測試員工|2026/8/1'], 'A班');
  assert.equal(c.shiftOverrides['other|2026/9/1'], 'A班');
  assert.equal(c.getReport().people[0].days[0].shift.name, '正常班');
  c.shiftOverrides['T1|測試員工|2026/9/1'] = 'A班';
  assert.equal(c.getReport().people[0].days[0].shift.name, 'A班');
  c.loadWorkbook(book);
  assert.equal(c.getReport().people[0].days[0].shift.name, '正常班');
});

test('custom times are preserved and missing built-in shifts remain selectable', () => {
  const { context: c } = harness();
  c.shifts = [{ name: '正常班', start: '09:00', end: '18:00' }];
  c.loadWorkbook({ SheetNames: ['data'], Sheets: { data: [input(1, '正常班'), input(2, 'A班')] } });
  const days = c.getReport().people[0].days;
  assert.equal(days[0].shift.start, '09:00');
  assert.equal(days[1].shift.end, '16:00');
  assert.ok(c.getAvailableShifts().some((s) => s.name === 'A班'));
});

test('embedded JavaScript parses', () => {
  new vm.Script(html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')));
});
