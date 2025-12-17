// --- CONFIGURATION ---

const BUILDING_CONFIG = {
  'OMS': {
    name: 'Orono Middle School',
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
    scheduleType: 'periods',
    periods: [
      "Period 1",
      "Period 2",
      "Period 3",
      "Period 4"
      // Add more as known or generic
    ],
    coverageTypes: [
      { label: 'Full Period', value: 1 }
    ]
  },
  'OIS': {
    name: 'Orono Intermediate School',
    scheduleType: 'time_range',
    increment: 15,
    coverageTypes: [
      { label: 'Time Duration', value: 'custom' }
    ]
  },
  'SES': {
    name: 'Schumann Elementary School',
    scheduleType: 'time_range',
    increment: 15,
    coverageTypes: [
      { label: 'Time Duration', value: 'custom' }
    ]
  }
};

/**
 * Helper to get building config.
 * Defaults to OMS if not found.
 */
function getBuildingConfig(buildingCode) {
  return BUILDING_CONFIG[buildingCode] || BUILDING_CONFIG['OMS'];
}
