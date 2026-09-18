// --- CONFIGURATION ---

const DEFAULT_BUILDING = 'OMS';

const BUILDING_CONFIG = {
  'OMS': {
    name: 'Orono Middle School',
    carryOverMax: 12,
    scheduleType: 'periods',
    periods: [
      "Period 1 - 8:10 - 8:57",
      "Period 2 - 9:01 - 9:48",
      "Period 3 - 9:52 - 10:39",
      "Period 4 - 10:43 - 11:09",
      "Period 5 - 11:11 - 11:37",
      "Period 4/5 - 10:30 - 11:37",
      "Period 6 - 11:40 - 12:06",
      "Period 7 - 12:08 - 12:34",
      "Period 6/7 - 11:40 - 12:34",
      "Period 8 - 12:37 - 1:08",
      "Period 9 - 1:12 - 1:59",
      "Period 10 - 2:03 - 2:50"
    ],
    coverageTypes: [
      { label: 'Full Period', value: 1 },
      { label: 'Half Period', value: 0.5 }
    ],
    // Special rules can be defined here if needed, e.g. "Period 6/7 is always 0.5"
    // The frontend currently handles simple selection. We can enforce rules in validation.
  },
  'OHS': {
    name: 'Orono High School',
    carryOverMax: 12,
    scheduleType: 'periods',
    // Coverage slots only: the bell chart's 10-minute Break is deliberately not
    // here, since nobody earns TST covering a passing break and it would sit in
    // every teacher's availability grid all year.
    //
    // Listed in Mon/Wed/Fri order, which is three days out of five. Spartan Hour
    // only runs Tue/Thu, so it has no default time — assigning it on a Monday is
    // reported as "no time set for that day" rather than quietly guessing one.
    periods: [
      "Period 1",
      "Spartan Hour",
      "Period 2",
      "Period 3",
      "Period 4",
      "Period 5A",
      "Period 5B",
      "Period 5C",
      "Period 6",
      "Period 7"
    ],
    // Mon/Wed/Fri — the default schedule.
    periodTimes: {
      "Period 1":  { start: "08:00", end: "08:48" },
      "Period 2":  { start: "08:52", end: "09:40" },
      "Period 3":  { start: "09:54", end: "10:42" },
      "Period 4":  { start: "10:46", end: "11:34" },
      "Period 5A": { start: "11:38", end: "12:02" },
      "Period 5B": { start: "12:05", end: "12:29" },
      "Period 5C": { start: "12:32", end: "12:56" },
      "Period 6":  { start: "13:00", end: "13:48" },
      "Period 7":  { start: "13:52", end: "14:40" }
    },
    // Tue/Thu runs shorter periods to make room for Spartan Hour.
    dayGroups: [
      {
        name: "TTh",
        days: ["Tue", "Thu"],
        times: {
          "Period 1":     { start: "08:00", end: "08:41" },
          "Spartan Hour": { start: "08:45", end: "09:25" },
          "Period 2":     { start: "09:40", end: "10:20" },
          "Period 3":     { start: "10:24", end: "11:04" },
          "Period 4":     { start: "11:08", end: "11:48" },
          "Period 5A":    { start: "11:52", end: "12:16" },
          "Period 5B":    { start: "12:20", end: "12:44" },
          "Period 5C":    { start: "12:48", end: "13:12" },
          "Period 6":     { start: "13:16", end: "13:56" },
          "Period 7":     { start: "14:00", end: "14:40" }
        }
      }
    ],
    coverageTypes: [
      { label: 'Full Period', value: 1 }
    ]
  },
  'OIS': {
    name: 'Orono Intermediate School',
    carryOverMax: 12,
    scheduleType: 'time_range',
    increment: 15,
    coverageTypes: [
      { label: 'Time Duration', value: 'custom' }
    ]
  },
  'SE': {
    name: 'Schumann Elementary School',
    carryOverMax: 12,
    scheduleType: 'time_range',
    increment: 15,
    coverageTypes: [
      { label: 'Time Duration', value: 'custom' }
    ]
  }
};
