/**
 * ==============================================================================================
 * CHAFFEY COLLEGE | ACADEMIC SCHEDULE GRID & ROOM CONFLICT ANALYZER
 * ==============================================================================================
 * Author: Antigravity Frontend Web Developer & UI/UX Designer
 * Platform: Pure Client-Side Web Application (100% GitHub Pages Ready)
 * Libraries: SheetJS (XLSX), html2canvas, jsPDF
 *
 * Core Technical Highlights & Data Pipeline:
 * 1. SECTION UNIQUE ID GROUPING:
 *    Academic sections frequently span multiple schedule rows (e.g. Lecture + Lab meetings).
 *    Rows sharing the same `Section Unique ID` and `Section Name` are aggregated into a single
 *    parent Course Section object containing an array of schedule meeting blocks.
 *
 * 2. SCHEDULE - DAYS NUMERICAL MAPPING:
 *    The `Schedule - Days` column uses a numerical format (e.g., "1", "2", "2,4").
 *    Values are comma-split and mapped to day integers:
 *    1 = Monday, 2 = Tuesday, 3 = Wednesday, 4 = Thursday, 5 = Friday, 6 = Saturday, 7 = Sunday.
 *
 * 3. SECTION TRAITS COMMA SPLITTING:
 *    The `Section Traits` column contains comma-separated values (e.g. "Not Basic Skills,In-Person").
 *    Values are trimmed and split into distinct arrays to populate individual filter tags.
 *
 * 4. CONFLICT DETECTION & STACKED CLASSES ENGINE:
 *    - STACKED CLASSES: Two different Section Unique IDs scheduled in the exact same Room on the
 *      same day/time with the SAME instructor. These are cross-listed concurrent courses and are
 *      NOT counted as room conflicts.
 *    - REAL ROOM CONFLICTS: Two different Section Unique IDs in the same Room on the same day/time
 *      with DIFFERENT instructors. These are flagged as scheduling errors.
 *    - DEFAULT TO ISOLATE CONFLICTS ONLY: When "View Room Conflicts" is activated, the grid
 *      immediately isolates only the conflicting sections.
 *
 * 5. VERTICALLY STACKED ROOM MATRIX:
 *    In Room Matrix view, instead of crowding all rooms onto one grid, each Building/Room expands
 *    into its own dedicated weekly calendar grid stacked vertically for clear room-by-room auditing.
 *
 * 6. GENERAL EDUCATION (GE) COVERAGE AUDIT VIEW:
 *    Evaluates GE Pattern and Area coverage for each campus location. Shows only GE sections with
 *    color-coded GE Area badges and identifies unscheduled/missing GE areas.
 * ==============================================================================================
 */

// --- 1. CONSTANTS & DAY MAPPINGS ---
/**
 * Numerical Day Mapping based on College scheduling specifications:
 * 1 = Monday, 2 = Tuesday, 3 = Wednesday, 4 = Thursday,
 * 5 = Friday, 6 = Saturday, 7 = Sunday.
 */
const DAY_MAP = {
  1: { id: 1, name: 'Monday', short: 'Mon' },
  2: { id: 2, name: 'Tuesday', short: 'Tue' },
  3: { id: 3, name: 'Wednesday', short: 'Wed' },
  4: { id: 4, name: 'Thursday', short: 'Thu' },
  5: { id: 5, name: 'Friday', short: 'Fri' },
  6: { id: 6, name: 'Saturday', short: 'Sat' },
  7: { id: 7, name: 'Sunday', short: 'Sun' }
};

// Standard California General Education Areas for Academic Auditing
// Valid Patterns: CalGETC (11 audited areas; 5C hidden) and Chaffey Local (7 audited areas)
const GE_STANDARD_AREAS = {
  'CALGETC': [
    { code: '1A', name: 'English Composition', desc: 'Written Communication', matchAreas: ['1A'] },
    { code: '1B', name: 'Critical Thinking and Composition', desc: 'Critical Thinking', matchAreas: ['1B'] },
    { code: '1C', name: 'Oral Communication', desc: 'Oral Communication / Public Speaking', matchAreas: ['1C'] },
    { code: '2',  name: 'Mathematical Concepts & Quantitative Reasoning', desc: 'Mathematics / Statistics', matchAreas: ['2'] },
    { code: '3A', name: 'Arts', desc: 'Arts, Cinema, Dance, Music, Theater', matchAreas: ['3A'] },
    { code: '3B', name: 'Humanities', desc: 'Literature, Philosophy, World Languages', matchAreas: ['3B'] },
    { code: '4',  name: 'Social and Behavioral Sciences', desc: 'History, Psychology, Sociology, Poli Sci', matchAreas: ['4'] },
    { code: '5A', name: 'Physical Sciences', desc: 'Chemistry, Physics, Astronomy, Geology', matchAreas: ['5A'] },
    { code: '5B', name: 'Biological Sciences', desc: 'Biology, Anatomy, Microbiology, Ecology', matchAreas: ['5B'] },
    { code: '6',  name: 'Ethnic Studies', desc: 'Ethnic Studies', matchAreas: ['6'] },
    { code: 'US1', name: 'American Institutions', desc: 'U.S. History, Constitution & American Ideals', matchAreas: ['US1'] }
  ],
  'CHAFFEY': [
    { code: '1A', name: 'English Composition', desc: 'Language & Rationality - English Composition', matchAreas: ['1A'] },
    { code: '1B or 1C', name: 'Communication & Analytical Thinking', desc: 'Critical Thinking (1B) or Oral Communication (1C)', matchAreas: ['1B', '1C'] },
    { code: '2',  name: 'Mathematics', desc: 'Mathematics Competency', matchAreas: ['2'] },
    { code: '3A or 3B', name: 'Arts & Humanities', desc: 'Arts (3A) or Humanities (3B)', matchAreas: ['3A', '3B'] },
    { code: '4',  name: 'Social and Behavioral Sciences', desc: 'Social & Behavioral Sciences', matchAreas: ['4'] },
    { code: '5',  name: 'Natural Sciences', desc: 'Physical and Life Sciences', matchAreas: ['5'] },
    { code: '6',  name: 'Ethnic Studies', desc: 'Ethnic Studies', matchAreas: ['6'] }
  ]
};

// Canonical sorting order for General Education areas
const GE_AREA_ORDER = ['1A', '1B', '1C', '2', '3A', '3B', '4', '5', '5A', '5B', '6', 'US1'];

function sortGeAreas(areas) {
  return Array.from(areas).sort((a, b) => {
    const idxA = GE_AREA_ORDER.indexOf(a);
    const idxB = GE_AREA_ORDER.indexOf(b);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b);
  });
}

function formatGePatternLabel(pattern) {
  if (pattern === 'CALGETC') return 'CalGETC';
  if (pattern === 'CHAFFEY') return 'Chaffey Local';
  return pattern;
}

function formatGeAreaLabel(area) {
  return `Area ${area}`;
}

/**
 * Format a section's sanitized GE Pattern & Area associations for popovers and modals,
 * suppressing any references to CSU, CSUGE, or IGETC.
 */
function formatSectionGeInfo(sec) {
  if (!sec || !sec.geAssociations || sec.geAssociations.length === 0) {
    return {
      patternsFormatted: 'None',
      areasFormatted: 'None',
      combinedFormatted: 'None'
    };
  }

  const calAreas = Array.from((sec.gePatternAreaMap && sec.gePatternAreaMap['CALGETC']) || []);
  const chafAreas = Array.from((sec.gePatternAreaMap && sec.gePatternAreaMap['CHAFFEY']) || []);

  const patternsList = [];
  const parts = [];

  if (calAreas.length > 0) {
    patternsList.push('CalGETC');
    parts.push(`CalGETC (Area ${sortGeAreas(calAreas).join(', ')})`);
  }
  if (chafAreas.length > 0) {
    patternsList.push('Chaffey Local');
    parts.push(`Chaffey Local (Area ${sortGeAreas(chafAreas).join(', ')})`);
  }

  const uniqueAreas = sortGeAreas(new Set([...calAreas, ...chafAreas]));

  return {
    patternsFormatted: patternsList.join(', ') || 'None',
    areasFormatted: uniqueAreas.map(a => `Area ${a}`).join(', ') || 'None',
    combinedFormatted: parts.join(' • ') || 'None'
  };
}

/**
 * DATA NORMALIZATION HELPERS:
 * Consistent normalization for Location (Campuses), Schedule - Buildings, and ACCs.
 */

// 1. Location (Campus) Normalization:
// "Chino" -> "Chino Campus", "Fontana Center" -> "Fontana Campus", "Rancho Campus" -> "Rancho Campus"
function normalizeLocation(rawLoc) {
  if (!rawLoc) return '';
  const s = String(rawLoc).trim();
  const lower = s.toLowerCase();
  if (lower === 'chino' || lower === 'chino campus') return 'Chino Campus';
  if (lower === 'fontana center' || lower === 'fontana' || lower === 'fontana campus') return 'Fontana Campus';
  if (lower === 'rancho' || lower === 'rancho campus') return 'Rancho Campus';
  return s;
}

// 2. Schedule - Building Normalization:
// "G" -> "Gym" (desc: "Gymnasium"), "Via" -> "Canvas" (desc: "Online")
function normalizeBuilding(bldgRaw, descRaw) {
  let bldg = String(bldgRaw || '').trim();
  let desc = String(descRaw || '').trim();

  const bldgUpper = bldg.toUpperCase();
  if (bldgUpper === 'G') {
    bldg = 'Gym';
    if (!desc || desc.toUpperCase() === 'G') desc = 'Gymnasium';
  } else if (bldgUpper === 'VIA') {
    bldg = 'Canvas';
    if (!desc || desc.toUpperCase() === 'VIA') desc = 'Online';
  }

  return { building: bldg, buildingDesc: desc };
}

// Format Building Dropdown and filter labels (e.g. "BE - Business Education")
function formatBuildingDropdownLabel(bldg, desc) {
  if (!bldg) return '';
  if (desc && desc.toLowerCase() !== bldg.toLowerCase()) {
    return `${bldg} - ${desc}`;
  }
  return bldg;
}

// 3. ACC (Academic & Career Community) Normalization:
// Normalizes variations (including "-Chino" variants) to the 8 standard categories
const STANDARD_ACCS = [
  'ACD: Arts, Communication, and Design',
  'BTH: Business, Technology, and Hospitality',
  'HW: Health and Wellness',
  'MIT: Manufacturing, Industrial Design, and Transportation',
  'PCS: Public Service, Culture, and Society',
  'STEM: Science, Technology, Engineering, and Mathematics',
  'Counseling and Matriculation',
  'Athletics'
];

function normalizeAcc(rawAcc) {
  if (!rawAcc) return '';
  let s = String(rawAcc).trim();
  const parts = s.split(',').map(p => p.trim()).filter(Boolean);
  if (parts.length > 0) {
    s = parts[0];
  }

  const sl = s.toLowerCase();

  // ACD: Arts, Communication, and Design (including -Chino variant)
  if (sl.includes('art') || sl.includes('communication') || sl.startsWith('acd')) {
    return 'ACD: Arts, Communication, and Design';
  }

  // BTH: Business, Technology, and Hospitality (including -Chino variant)
  if (sl.includes('business') || sl.includes('hosp') || sl.startsWith('bth')) {
    return 'BTH: Business, Technology, and Hospitality';
  }

  // HW: Health and Wellness
  if (sl.includes('health') || sl.includes('wellness') || sl.startsWith('hw')) {
    return 'HW: Health and Wellness';
  }

  // MIT: Manufacturing, Industrial Design, and Transportation
  if (sl.includes('manufacturing') || sl.includes('transportation') || sl.includes('ind design') || sl.startsWith('mit')) {
    return 'MIT: Manufacturing, Industrial Design, and Transportation';
  }

  // PCS: Public Service, Culture, and Society
  if (sl.includes('public service') || sl.includes('society') || sl.includes('culture') || sl.startsWith('pcs')) {
    return 'PCS: Public Service, Culture, and Society';
  }

  // STEM: Science, Technology, Engineering, and Mathematics
  if (sl.includes('stem') || sl.includes('science') || sl.includes('math') || sl.includes('engineering')) {
    return 'STEM: Science, Technology, Engineering, and Mathematics';
  }

  // Counseling and Matriculation
  if (sl.includes('counseling') || sl.includes('matriculation')) {
    return 'Counseling and Matriculation';
  }

  // Athletics
  if (sl.includes('athletic') || sl.includes('kinesiology')) {
    return 'Athletics';
  }

  return s;
}

function sortAccs(accList) {
  return Array.from(accList).sort((a, b) => {
    const idxA = STANDARD_ACCS.indexOf(a);
    const idxB = STANDARD_ACCS.indexOf(b);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b);
  });
}

// Built-in Dataset matching Schedule Sections Building-Room Grids.xlsx
// Enriched with stacked classes and multi-campus GE sections for full feature demonstration
const BUILT_IN_DATASET = [
  {
    "Section Unique ID": "210928",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48898",
    "Section Number": "48898",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Rhodes, Linda",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "12:00:00 PM",
    "Schedule - End Time": "12:50:00 PM",
    "Schedule - Days": "2",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "210928",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48898",
    "Section Number": "48898",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Rhodes, Linda",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "1:00:00 PM",
    "Schedule - End Time": "3:50:00 PM",
    "Schedule - Days": "2",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  // STACKED CLASS EXAMPLE: FASHD-45 meets at the same time & room as FASHD-40 with SAME instructor (Rhodes, Linda)
  {
    "Section Unique ID": "210929",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-45",
    "Section Name": "FASHD-45-48900",
    "Section Number": "48900",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Rhodes, Linda",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "1:00:00 PM",
    "Schedule - End Time": "3:50:00 PM",
    "Schedule - Days": "2",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "210954",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48924",
    "Section Number": "48924",
    "Location": "Chino Campus",
    "Day or Evening": "E",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Zero Textbook Cost ($0),ZTC; Inst mat free to stu",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Kim, Hyo Jeong",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "6:00:00 PM",
    "Schedule - End Time": "6:50:00 PM",
    "Schedule - Days": "4",
    "Schedule - Start Date": "Aug 20, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "210954",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48924",
    "Section Number": "48924",
    "Location": "Chino Campus",
    "Day or Evening": "E",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Zero Textbook Cost ($0),ZTC; Inst mat free to stu",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Kim, Hyo Jeong",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "7:00:00 PM",
    "Schedule - End Time": "9:50:00 PM",
    "Schedule - Days": "4",
    "Schedule - Start Date": "Aug 20, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "210963",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-42",
    "Section Name": "FASHD-42-48933",
    "Section Number": "48933",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Green, Jasmine",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "9:30:00 AM",
    "Schedule - End Time": "10:20:00 AM",
    "Schedule - Days": "5",
    "Schedule - Start Date": "Aug 21, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "210963",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-42",
    "Section Name": "FASHD-42-48933",
    "Section Number": "48933",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Green, Jasmine",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "10:30:00 AM",
    "Schedule - End Time": "1:20:00 PM",
    "Schedule - Days": "5",
    "Schedule - Start Date": "Aug 21, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "210994",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-470",
    "Section Name": "FASHD-470-48964",
    "Section Number": "48964",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Gomez, Lylia",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "206",
    "Schedule - Start Time": "12:00:00 PM",
    "Schedule - End Time": "2:50:00 PM",
    "Schedule - Days": "4",
    "Schedule - Start Date": "Aug 20, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "211007",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-61",
    "Section Name": "FASHD-61-48977",
    "Section Number": "48977",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Zero Textbook Cost ($0),ZTC; No assc inst materia",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Rhodes, Linda",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "9:30:00 AM",
    "Schedule - End Time": "11:20:00 AM",
    "Schedule - Days": "1",
    "Schedule - Start Date": "Aug 17, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "211007",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-61",
    "Section Name": "FASHD-61-48977",
    "Section Number": "48977",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Zero Textbook Cost ($0),ZTC; No assc inst materia",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Rhodes, Linda",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "11:30:00 AM",
    "Schedule - End Time": "2:20:00 PM",
    "Schedule - Days": "1",
    "Schedule - Start Date": "Aug 17, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "211058",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Merchandising",
    "Course": "FASHM-60",
    "Section Name": "FASHM-60-49027",
    "Section Number": "49027",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Early 14-Week",
    "Modality": "In-Person",
    "Number of Weeks": 14,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,In-Person,Traditional Textbook,Early 14-Week Class",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Johnson, Tara",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "206",
    "Schedule - Start Time": "9:30:00 AM",
    "Schedule - End Time": "11:05:00 AM",
    "Schedule - Days": "2,4",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Nov 10, 2026"
  },
  {
    "Section Unique ID": "213779",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-72",
    "Section Name": "FASHD-72-51708",
    "Section Number": "51708",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "14-Week",
    "Modality": "In-Person",
    "Number of Weeks": 14,
    "Section Traits": "14-Week Class,In-Person,Not Basic Skills,Not Coop Work Experience,Zero Textbook Cost ($0),ZTC; No-cost OER",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Green, Jasmine",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "3:00:00 PM",
    "Schedule - End Time": "4:05:00 PM",
    "Schedule - Days": "1",
    "Schedule - Start Date": "Sep 9, 2026",
    "Schedule - End Date": "Dec 9, 2026"
  },
  {
    "Section Unique ID": "213779",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-72",
    "Section Name": "FASHD-72-51708",
    "Section Number": "51708",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "14-Week",
    "Modality": "In-Person",
    "Number of Weeks": 14,
    "Section Traits": "14-Week Class,In-Person,Not Basic Skills,Not Coop Work Experience,Zero Textbook Cost ($0),ZTC; No-cost OER",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Green, Jasmine",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "4:05:00 PM",
    "Schedule - End Time": "5:20:00 PM",
    "Schedule - Days": "1",
    "Schedule - Start Date": "Sep 9, 2026",
    "Schedule - End Date": "Dec 9, 2026"
  },
  {
    "Section Unique ID": "213779",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Fashion Design",
    "Course": "FASHD-72",
    "Section Name": "FASHD-72-51708",
    "Section Number": "51708",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "14-Week",
    "Modality": "In-Person",
    "Number of Weeks": 14,
    "Section Traits": "14-Week Class,In-Person,Not Basic Skills,Not Coop Work Experience,Zero Textbook Cost ($0),ZTC; No-cost OER",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Green, Jasmine",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHCM",
    "Schedule - Building Description": "Chino Community Center Bldg",
    "Schedule - Room": "203",
    "Schedule - Start Time": "3:00:00 PM",
    "Schedule - End Time": "5:20:00 PM",
    "Schedule - Days": "3",
    "Schedule - Start Date": "Sep 9, 2026",
    "Schedule - End Date": "Dec 9, 2026"
  },
  // Chino Campus GE Courses (Covering Area 1C, Area 1A, Area 3A)
  {
    "Section Unique ID": "209114",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Communication Studies",
    "Course": "COMM-C1000",
    "Section Name": "COMM-C1000-47109",
    "Section Number": "47109",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "14-Week",
    "Modality": "In-Person",
    "Number of Weeks": 14,
    "Section Traits": "Not Basic Skills,Not Coop Work Experience,Local GE Math Comp Only,In-Person,14-Week Class,Zero Textbook Cost ($0),ZTC; Mix of No-Cost",
    "General Ed Pattern": "CHAFFEY,CALGETC",
    "General Ed Area": "1C,1C",
    "Faculty Name": "Ith, Phatana",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instruction Bldg",
    "Schedule - Room": "241",
    "Schedule - Start Time": "9:00:00 AM",
    "Schedule - End Time": "10:50:00 AM",
    "Schedule - Days": "2,4",
    "Schedule - Start Date": "Sep 8, 2026",
    "Schedule - End Date": "Dec 11, 2026"
  },
  {
    "Section Unique ID": "209115",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Art",
    "Course": "ART-10",
    "Section Name": "ART-10-47200",
    "Section Number": "47200",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Zero Textbook Cost ($0)",
    "General Ed Pattern": "CHAFFEY,CALGETC",
    "General Ed Area": "3A,3A",
    "Faculty Name": "Gale, Robert",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instruction Bldg",
    "Schedule - Room": "120",
    "Schedule - Start Time": "11:00:00 AM",
    "Schedule - End Time": "12:20:00 PM",
    "Schedule - Days": "2,4",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "209116",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "English",
    "Course": "ENGL-1A",
    "Section Name": "ENGL-1A-47300",
    "Section Number": "47300",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "CHAFFEY,CALGETC",
    "General Ed Area": "1A,1A",
    "Faculty Name": "Alvarez, Maria",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instruction Bldg",
    "Schedule - Room": "240",
    "Schedule - Start Time": "9:00:00 AM",
    "Schedule - End Time": "10:50:00 AM",
    "Schedule - Days": "1,3",
    "Schedule - Start Date": "Aug 17, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  // Rancho Campus GE Sections
  {
    "Section Unique ID": "220101",
    "Term": "2026/FA",
    "ACC": "STEM: Science, Technology, Engineering, and Mathematics",
    "Subject": "Mathematics",
    "Course": "MATH-25",
    "Section Name": "MATH-25-50111",
    "Section Number": "50111",
    "Location": "Rancho Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,In-Person,Transferable to UC/CSU,Zero Textbook Cost ($0)",
    "General Ed Pattern": "CHAFFEY,CALGETC",
    "General Ed Area": "2,2",
    "Faculty Name": "Brouwer, Peter",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "VSS",
    "Schedule - Building Description": "Vocational & Student Support",
    "Schedule - Room": "101",
    "Schedule - Start Time": "10:00:00 AM",
    "Schedule - End Time": "11:50:00 AM",
    "Schedule - Days": "1,3",
    "Schedule - Start Date": "Aug 17, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "220102",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "English",
    "Course": "ENGL-1A",
    "Section Name": "ENGL-1A-50222",
    "Section Number": "50222",
    "Location": "Rancho Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "CALGETC",
    "General Ed Area": "1A",
    "Faculty Name": "Miller, Sarah",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "WH",
    "Schedule - Building Description": "Wargin Hall",
    "Schedule - Room": "140",
    "Schedule - Start Time": "1:00:00 PM",
    "Schedule - End Time": "2:50:00 PM",
    "Schedule - Days": "2,4",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  // Evening & Weekend Sections Demo
  {
    "Section Unique ID": "224501",
    "Term": "2026/FA",
    "ACC": "PCS: Public Service, Culture, and Society",
    "Subject": "Psychology",
    "Course": "PSYCH-1",
    "Section Name": "PSYCH-1-51204",
    "Section Number": "51204",
    "Location": "Chino Campus",
    "Day or Evening": "E",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Evening Class,Zero Textbook Cost ($0)",
    "General Ed Pattern": "CALGETC",
    "General Ed Area": "CalGETC,4",
    "Faculty Name": "Alvarez, Marco",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instructional Bldg",
    "Schedule - Room": "101",
    "Schedule - Start Time": "6:00:00 PM",
    "Schedule - End Time": "8:50:00 PM",
    "Schedule - Days": "1,3",
    "Schedule - Start Date": "Aug 17, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "224502",
    "Term": "2026/FA",
    "ACC": "ACD: Arts, Communication, and Design",
    "Subject": "Art",
    "Course": "ART-10",
    "Section Name": "ART-10-51402",
    "Section Number": "51402",
    "Location": "Chino Campus",
    "Day or Evening": "E",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Friday Class,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "CALGETC",
    "General Ed Area": "CalGETC,3A",
    "Faculty Name": "Ramirez, Elena",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instructional Bldg",
    "Schedule - Room": "102",
    "Schedule - Start Time": "4:30:00 PM",
    "Schedule - End Time": "7:20:00 PM",
    "Schedule - Days": "5",
    "Schedule - Start Date": "Aug 21, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "224503",
    "Term": "2026/FA",
    "ACC": "STEM: Science, Technology, Engineering, and Mathematics",
    "Subject": "Biology",
    "Course": "BIOL-10",
    "Section Name": "BIOL-10-51610",
    "Section Number": "51610",
    "Location": "Rancho Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Weekend College,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "CALGETC",
    "General Ed Area": "CalGETC,5B",
    "Faculty Name": "Chen, David",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "WH",
    "Schedule - Building Description": "Wargin Hall",
    "Schedule - Room": "140",
    "Schedule - Start Time": "9:00:00 AM",
    "Schedule - End Time": "12:10:00 PM",
    "Schedule - Days": "6",
    "Schedule - Start Date": "Aug 22, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  // Multi-Instructor Program Stacking Demo (DENTAL Stack vs RADTEC Conflict)
  {
    "Section Unique ID": "225101",
    "Term": "2026/FA",
    "ACC": "STEM: Science, Technology, Engineering, and Mathematics",
    "Subject": "Dental Assisting",
    "Course": "DENTAL-100",
    "Section Name": "DENTAL-100-52010",
    "Section Number": "52010",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Dental Clinic Cohort,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Warren, James",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instructional Bldg",
    "Schedule - Room": "101",
    "Schedule - Start Time": "2:00:00 PM",
    "Schedule - End Time": "3:50:00 PM",
    "Schedule - Days": "2",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "225102",
    "Term": "2026/FA",
    "ACC": "STEM: Science, Technology, Engineering, and Mathematics",
    "Subject": "Dental Assisting",
    "Course": "DENTAL-102",
    "Section Name": "DENTAL-102-52012",
    "Section Number": "52012",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Dental Clinic Cohort,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Nguyen, Tina",
    "Schedule - Instruction Type": "Lab; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instructional Bldg",
    "Schedule - Room": "101",
    "Schedule - Start Time": "2:00:00 PM",
    "Schedule - End Time": "3:50:00 PM",
    "Schedule - Days": "2",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  },
  {
    "Section Unique ID": "225103",
    "Term": "2026/FA",
    "ACC": "STEM: Science, Technology, Engineering, and Mathematics",
    "Subject": "Radiologic Technology",
    "Course": "RADTEC-10",
    "Section Name": "RADTEC-10-52020",
    "Section Number": "52020",
    "Location": "Chino Campus",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Radiologic Clinic",
    "General Ed Pattern": "",
    "General Ed Area": "",
    "Faculty Name": "Davis, Robert",
    "Schedule - Instruction Type": "Lec; Required In-Person Mtgs",
    "Schedule - Building": "CHMB",
    "Schedule - Building Description": "Chino Main Instructional Bldg",
    "Schedule - Room": "101",
    "Schedule - Start Time": "3:00:00 PM",
    "Schedule - End Time": "4:50:00 PM",
    "Schedule - Days": "2",
    "Schedule - Start Date": "Aug 18, 2026",
    "Schedule - End Date": "Dec 18, 2026"
  }
];

const SPREADSHEET_EXACT_DATASET = BUILT_IN_DATASET.slice(0, 14);
const MULTI_CAMPUS_DATASET = BUILT_IN_DATASET;

// --- 2. APPLICATION STATE ---
const appState = {
  rawRows: [],             // All parsed spreadsheet rows
  sections: [],            // Grouped parent course objects
  filteredSections: [],    // Currently matched parent course objects after all active filters
  scheduleBlocks: [],      // Flattened list of individual day meeting occurrences
  filteredBlocks: [],      // Currently displayed meeting blocks after filtering
  conflicts: [],           // List of detected real room conflicts (different instructors)
  stackedPairs: [],        // List of detected stacked classes (same instructor or known cohort program)
  isConflictMode: false,   // "View Room Conflicts" toggle state
  conflictsOnly: true,     // Default to Isolate Conflicts Only when conflict mode is active!
  daySpan: 'active',       // Default to Active Days: only days with scheduled classes in view
  layoutMode: 'room',      // Default to Room View rather than Day View
  geSubLayout: 'room',     // Default GE Coverage grid view to Room View
  eveningSubLayout: 'room',// Default Evening Coverage grid view to Room View
  fasttrackSubLayout: 'room', // Default Fast-Track Coverage grid view to Room View
  fasttrackActiveTab: 'all',  // Default to showing all Fast-Track audit items (pairs and unpaired)
  efficiencySubLayout: 'room', // Default Efficiency grid view to Room View
  efficiencyActiveTab: 'all',  // Default to showing all Efficiency audit items
  currentEfficiencyAnalysis: null,
  geActivePattern: 'CALGETC', // Active pattern for GE Coverage audit
  geShowMissingOnly: false,
  filters: {
    search: '',
    location: '',
    building: '',
    room: '',
    dayEvening: new Set(),
    geArea: new Set(),
    gePattern: new Set(),
    acc: new Set(),
    subject: new Set(),
    course: new Set(),
    modality: new Set(),
    session: new Set(),
    traits: new Set()
  }
};

/**
 * Formats Day or Evening codes (e.g. 'D' -> 'Day (D)', 'E' -> 'Evening (E)')
 */
function formatDayEveningLabel(val) {
  if (!val) return 'Unspecified';
  const upper = String(val).trim().toUpperCase();
  if (upper === 'D' || upper === 'DAY') return 'Day (D)';
  if (upper === 'E' || upper === 'EVENING') return 'Evening (E)';
  return val;
}

// --- 3. TIME PARSING & FORMATTING UTILITIES ---

function parseTimeToMinutes(val) {
  if (val === null || val === undefined || val === '') return null;

  if (typeof val === 'number') {
    if (val >= 0 && val < 1) {
      return Math.round(val * 24 * 60);
    }
    if (val >= 100 && val <= 2400) {
      const h = Math.floor(val / 100);
      const m = val % 100;
      return h * 60 + m;
    }
  }

  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }

  const str = String(val).trim();
  if (!str) return null;

  const ampmMatch = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if (ampmMatch) {
    let hours = parseInt(ampmMatch[1], 10);
    const minutes = parseInt(ampmMatch[2], 10);
    const meridiem = ampmMatch[3].toUpperCase();
    if (meridiem === 'PM' && hours < 12) hours += 12;
    if (meridiem === 'AM' && hours === 12) hours = 0;
    return hours * 60 + minutes;
  }

  const milMatch = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (milMatch) {
    const hours = parseInt(milMatch[1], 10);
    const minutes = parseInt(milMatch[2], 10);
    return hours * 60 + minutes;
  }

  return null;
}

function formatMinutesToAmPm(minutes) {
  if (minutes === null || minutes === undefined || isNaN(minutes)) return '';
  let h = Math.floor(minutes / 60);
  const m = minutes % 60;
  const meridiem = h >= 12 ? 'PM' : 'AM';
  h = h % 12;
  if (h === 0) h = 12;
  return `${h}:${m < 10 ? '0' : ''}${m} ${meridiem}`;
}

function normalizeFacultyName(name) {
  if (!name) return '';
  return String(name).toLowerCase().replace(/[^a-z0-9]/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * Helper to identify Evening and Weekend Class blocks/sections:
 * - Scheduled on Friday (5), Saturday (6), or Sunday (7)
 * - Scheduled on or after 4:30 PM (startMinutes >= 990)
 * - Schedule Type is "E" evening OR Day or Evening column is "E"
 */
function isEveningBlock(block) {
  if (!block) return false;
  if (block.day === 5 || block.day === 6 || block.day === 7) return true;
  if (block.startMinutes !== null && block.startMinutes >= 990) return true;
  const sec = block.parentSection;
  if (sec) {
    if (sec.dayOrEvening && String(sec.dayOrEvening).trim().toUpperCase().startsWith('E')) return true;
    if (block.instructionType && String(block.instructionType).toUpperCase().includes('EVENING')) return true;
    if (sec.sectionTraits && sec.sectionTraits.some(t => t.toUpperCase().includes('EVENING'))) return true;
  }
  return false;
}

function isEveningSection(sec) {
  if (!sec) return false;
  if (sec.dayOrEvening && String(sec.dayOrEvening).trim().toUpperCase().startsWith('E')) return true;
  if (sec.scheduleBlocks && sec.scheduleBlocks.some(b => {
    const days = b.days || [];
    const isFriSatSun = days.some(d => d === 5 || d === 6 || d === 7);
    const isLate = b.startMinutes !== null && b.startMinutes >= 990;
    const isEveType = b.instructionType && String(b.instructionType).toUpperCase().includes('EVENING');
    return isFriSatSun || isLate || isEveType;
  })) return true;
  if (sec.sectionTraits && sec.sectionTraits.some(t => t.toUpperCase().includes('EVENING'))) return true;
  return false;
}

/**
 * Helpers to identify Fast-Track 1 & Fast-Track 2 sections and meeting blocks.
 * Fast-Track 1: First 8-week session in the 18-week semester.
 * Fast-Track 2: Second 8-week session in the 18-week semester.
 */
function getFastTrackSessionType(sec) {
  if (!sec) return null;
  const sess = String(sec.calendarSession || '').toLowerCase();
  if (sess.includes('fast-track 1') || sess.includes('fast track 1') || sess === 'ft1') return 'FT1';
  if (sess.includes('fast-track 2') || sess.includes('fast track 2') || sess === 'ft2') return 'FT2';
  return null;
}

function isFastTrackSection(sec) {
  return Boolean(getFastTrackSessionType(sec));
}

function isFastTrackBlock(block) {
  return block ? isFastTrackSection(block.parentSection) : false;
}

/**
 * Helper to determine if a scheduled meeting block/component has complete physical scheduling info.
 * For Fast-Track Coverage:
 * - IGNORES meeting components/parts that do not list days, times, locations, or hours (e.g., Success Center Hrs in #209439).
 * - IGNORES components/parts listed as Online Lec / Canvas / Via (e.g., Online Lec; Location: Canvas CANVAS in #213742).
 */
function isPhysicalMeetingBlock(blk) {
  if (!blk) return false;
  // Must list meeting days (supports meetingRow with blk.days and calendar block with blk.day / blk.meetingRow.days)
  const days = blk.days || (blk.meetingRow && blk.meetingRow.days) || (blk.day ? [blk.day] : []);
  if (!days || days.length === 0) return false;
  // Must list start and end times
  if (blk.startMinutes === null || blk.endMinutes === null) return false;
  // Must list building & room
  const bldg = String(blk.building || '').trim();
  const room = String(blk.room || '').trim();
  if (!bldg || !room) return false;

  const bldgUpper = bldg.toUpperCase();
  const roomUpper = room.toUpperCase();
  const instUpper = String(blk.instructionType || '').toUpperCase();

  // Ignore components/parts listed as Online Lec / Canvas / Via (Online)
  if (bldgUpper === 'CANVAS' || bldgUpper === 'VIA' || bldgUpper === 'ONLINE') return false;
  if (roomUpper === 'CANVAS' || roomUpper === 'ONLINE') return false;
  if (instUpper.includes('ONLINE') && !instUpper.includes('IN-PERSON')) return false;

  return true;
}

/**
 * ==============================================================================================
 * ENROLLMENT EFFICIENCY & LOW ENROLLED LOGIC:
 * Evaluates whether a section is Low Enrolled according to institutional criteria:
 * - If Capacity >= 30: Low enrolled if Active Students < 20.
 * - If Capacity between 17 and 29: Low enrolled if Active Students < 17.
 * - If Capacity > 0 and < 17: Low enrolled if Fill Rate (Active / Capacity) < 70%.
 * - Otherwise: Not low enrolled.
 * ==============================================================================================
 */
function isLowEnrolledSection(sec) {
  if (!sec) return false;
  const cap = Number(sec.capacity) || 0;
  const active = Number(sec.activeStudents) || 0;
  if (cap >= 30 && active < 20) return true;
  if (cap >= 17 && cap <= 29 && active < 17) return true;
  if (cap > 0 && cap < 17 && (active / cap) < 0.70) return true;
  return false;
}

function getLowEnrolledReason(sec) {
  if (!sec) return '';
  const cap = Number(sec.capacity) || 0;
  const active = Number(sec.activeStudents) || 0;
  const fillPct = cap > 0 ? Math.round((active / cap) * 100) : 0;

  if (cap >= 30 && active < 20) {
    return `Capacity ${cap} (≥30) with only ${active} students enrolled (threshold: 20 students, currently ${fillPct}% fill).`;
  }
  if (cap >= 17 && cap <= 29 && active < 17) {
    return `Capacity ${cap} (17–29) with only ${active} students enrolled (threshold: 17 students, currently ${fillPct}% fill).`;
  }
  if (cap > 0 && cap < 17 && (active / cap) < 0.70) {
    const requiredStudents = Math.ceil(cap * 0.70);
    return `Small-capacity class (${cap} seats) at ${fillPct}% fill rate (threshold: 70%, requires at least ${requiredStudents} students).`;
  }
  return '';
}

// --- 4. DATA NORMALIZATION & GROUPING LOGIC ---

/**
 * INLINE COMMENT EXPLAINING SCHEDULE - DAYS NUMERICAL MAPPING:
 * The spreadsheet represents days of the week as integers or comma-delimited strings
 * (e.g., "1", "2", "2,4").
 * We split by comma, trim whitespace, and map each value:
 * 1 = Monday, 2 = Tuesday, 3 = Wednesday, 4 = Thursday, 5 = Friday, 6 = Saturday, 7 = Sunday.
 * Returns an array of valid day numbers (e.g., [2, 4]).
 */
function parseScheduleDays(daysRaw) {
  if (daysRaw === null || daysRaw === undefined) return [];
  const str = String(daysRaw).trim();
  if (!str) return [];

  return str.split(',')
    .map(d => parseInt(d.trim(), 10))
    .filter(d => !isNaN(d) && d >= 1 && d <= 7);
}

/**
 * INLINE COMMENT EXPLAINING SECTION TRAITS COMMA SPLITTING:
 * The `Section Traits` column contains comma-separated values
 * (e.g., "Not Basic Skills,Not Coop Work Experience,In-Person,Low-Cost Textbook (<$50)").
 * We split on ',', trim each entry, filter out empty strings, and store them as a distinct array.
 * This allows each individual trait to be populated into a dynamic multi-select filter list.
 */
function parseCommaSeparatedList(rawStr) {
  if (!rawStr) return [];
  return String(rawStr)
    .split(',')
    .map(s => s.trim())
    .filter(s => s.length > 0);
}

/**
 * INLINE COMMENT EXPLAINING GENERAL EDUCATION MULTIVALUED ASSOCIATIONS:
 * In scheduling records, `General Ed Pattern` and `General Ed Area` are multivalued
 * position-correlated comma-separated lists. E.g.:
 * General Ed Pattern: "CHAFFEY,IGETC,CSU,CALGETC"
 * General Ed Area:    "2,2,B4,2"
 * Index 0 maps CHAFFEY -> Area 2
 * Index 1 maps IGETC -> Area 2
 * Index 2 maps CSU -> Area B4
 * Index 3 maps CALGETC -> Area 2
 *
 * Requirements:
 * 1. Hide all references to CSU and IGETC patterns.
 * 2. CalGETC valid options: 1A, 1B, 1C, 2, 3A, 3B, 4, 5A, 5B, 6, US1 (hide 5C).
 * 3. Chaffey Local valid options: 1A, 1B, 1C, 2, 3A, 3B, 4, 5, 6 (remove invalid 2A).
 */
function parseGeAssociations(patternRaw, areaRaw) {
  if (!patternRaw || !areaRaw) return [];
  const patterns = String(patternRaw).split(',').map(s => s.trim()).filter(Boolean);
  const areas = String(areaRaw).split(',').map(s => s.trim()).filter(Boolean);

  const minLen = Math.min(patterns.length, areas.length);
  const associations = [];
  const seen = new Set();

  for (let i = 0; i < minLen; i++) {
    const rawP = patterns[i].toUpperCase();
    const rawA = areas[i].toUpperCase();

    let normPattern = null;
    if (rawP === 'CALGETC' || rawP.includes('CALGETC')) {
      normPattern = 'CALGETC';
    } else if (rawP === 'CHAFFEY' || rawP.includes('CHAFFEY')) {
      normPattern = 'CHAFFEY';
    }
    // Omit CSU, CSUGE, IGETC, etc.
    if (!normPattern) continue;

    if (normPattern === 'CALGETC') {
      const validCalGetc = ['1A', '1B', '1C', '2', '3A', '3B', '4', '5A', '5B', '6', 'US1'];
      if (validCalGetc.includes(rawA)) {
        const key = `${normPattern}:${rawA}`;
        if (!seen.has(key)) {
          seen.add(key);
          associations.push({ pattern: normPattern, area: rawA });
        }
      }
    } else if (normPattern === 'CHAFFEY') {
      const validChaffey = ['1A', '1B', '1C', '2', '3A', '3B', '4', '5', '6'];
      if (validChaffey.includes(rawA)) {
        const key = `${normPattern}:${rawA}`;
        if (!seen.has(key)) {
          seen.add(key);
          associations.push({ pattern: normPattern, area: rawA });
        }
      }
    }
  }

  return associations;
}

/**
 * INLINE COMMENT EXPLAINING SECTION UNIQUE ID GROUPING:
 * A single academic course section can have multiple rows in the scheduling spreadsheet
 * (for example, a Lecture meeting and a Lab meeting in FASHD-40, or multiple meeting days).
 * We group rows by composite key: `${Section Unique ID}||${Section Name}`.
 * The primary course attributes (Subject, Course, Modality, Faculty, Traits, etc.) are
 * captured on the parent section object. Each individual row is preserved as a schedule block
 * in the parent's `scheduleBlocks` array.
 */
function processRawDataset(rawData) {
  appState.rawRows = rawData;
  const sectionMap = new Map();
  const allMeetingBlocks = [];

  rawData.forEach((row, rowIndex) => {
    const sectionUniqueId = String(row['Section Unique ID'] || `SEC_${rowIndex}`).trim();
    const sectionName = String(row['Section Name'] || 'Unknown Section').trim();
    const groupKey = `${sectionUniqueId}||${sectionName}`;

    if (!sectionMap.has(groupKey)) {
      const traitsList = parseCommaSeparatedList(row['Section Traits']);
      const associations = parseGeAssociations(row['General Ed Pattern'], row['General Ed Area']);
      const patternSet = new Set();
      const areaSet = new Set();
      const patternAreaMap = { 'CALGETC': new Set(), 'CHAFFEY': new Set() };

      associations.forEach(pair => {
        patternSet.add(pair.pattern);
        areaSet.add(pair.area);
        if (patternAreaMap[pair.pattern]) {
          patternAreaMap[pair.pattern].add(pair.area);
        }
      });

      const rowTitle = String(row['Title'] || row['Course Title'] || '').trim();
      const rowCap = parseFloat(row['Capacity']);
      const rowCapacity = !isNaN(rowCap) ? rowCap : 0;
      const rowAct = parseFloat(row['Active Students']);
      const rowActive = !isNaN(rowAct) ? rowAct : 0;
      const rowHrs = parseFloat(row['Total Hours']);
      const rowHours = !isNaN(rowHrs) ? rowHrs : 0;
      const rowFtesRaw = parseFloat(row['Estimated FTES']);
      const rowFtes = !isNaN(rowFtesRaw) ? rowFtesRaw : (rowHours > 0 ? (rowActive * rowHours) / 525 : 0);

      sectionMap.set(groupKey, {
        sectionUniqueId: sectionUniqueId,
        sectionName: sectionName,
        title: rowTitle,
        capacity: rowCapacity,
        activeStudents: rowActive,
        totalHours: rowHours,
        estimatedFtes: rowFtes,
        unfilledSeats: Math.max(0, rowCapacity - rowActive),
        fillRate: rowCapacity > 0 ? (rowActive / rowCapacity) : 0,
        lostFtes: rowHours > 0
          ? (Math.max(0, rowCapacity - rowActive) * rowHours) / 525
          : (rowActive > 0 ? (Math.max(0, rowCapacity - rowActive) * rowFtes / rowActive) : 0),
        isLowEnrolled: isLowEnrolledSection({ capacity: rowCapacity, activeStudents: rowActive }),
        sectionNumber: String(row['Section Number'] || '').trim(),
        term: String(row['Term'] || '').trim(),
        acc: normalizeAcc(row['ACC']),
        subject: String(row['Subject'] || '').trim(),
        course: String(row['Course'] || '').trim(),
        location: normalizeLocation(row['Location']),
        dayOrEvening: String(row['Day or Evening'] || '').trim(),
        calendarSession: String(row['Calendar Session'] || '').trim(),
        modality: String(row['Modality'] || '').trim(),
        numberOfWeeks: row['Number of Weeks'] || '',
        facultyName: String(row['Faculty Name'] || '').trim(),
        sectionTraits: traitsList,
        sectionTraitsRaw: row['Section Traits'] || '',
        geAssociations: associations,
        gePatternAreaMap: patternAreaMap,
        generalEdPattern: Array.from(patternSet),
        generalEdPatternRaw: row['General Ed Pattern'] || '',
        generalEdArea: Array.from(areaSet),
        generalEdAreaRaw: row['General Ed Area'] || '',
        scheduleBlocks: []
      });
    }

    const parentSection = sectionMap.get(groupKey);

    // Merge or update enrollment fields if present on subsequent rows
    const rowTitle = String(row['Title'] || row['Course Title'] || '').trim();
    if (rowTitle && !parentSection.title) parentSection.title = rowTitle;
    const rowCap = parseFloat(row['Capacity']);
    if (!isNaN(rowCap) && rowCap > 0 && parentSection.capacity === 0) parentSection.capacity = rowCap;
    const rowAct = parseFloat(row['Active Students']);
    if (!isNaN(rowAct) && rowAct > 0 && parentSection.activeStudents === 0) parentSection.activeStudents = rowAct;
    const rowHrs = parseFloat(row['Total Hours']);
    if (!isNaN(rowHrs) && rowHrs > 0 && parentSection.totalHours === 0) parentSection.totalHours = rowHrs;
    const rowFtes = parseFloat(row['Estimated FTES']);
    if (!isNaN(rowFtes) && rowFtes > 0 && parentSection.estimatedFtes === 0) parentSection.estimatedFtes = rowFtes;

    parentSection.unfilledSeats = Math.max(0, parentSection.capacity - parentSection.activeStudents);
    parentSection.fillRate = parentSection.capacity > 0 ? (parentSection.activeStudents / parentSection.capacity) : 0;
    parentSection.lostFtes = parentSection.totalHours > 0
      ? (parentSection.unfilledSeats * parentSection.totalHours) / 525
      : (parentSection.activeStudents > 0 && parentSection.estimatedFtes > 0 ? (parentSection.unfilledSeats * parentSection.estimatedFtes / parentSection.activeStudents) : 0);
    parentSection.isLowEnrolled = isLowEnrolledSection(parentSection);

    // Merge any additional GE associations across multiple rows of the same section if present
    if (row['General Ed Pattern'] && row['General Ed Area']) {
      const rowAssocs = parseGeAssociations(row['General Ed Pattern'], row['General Ed Area']);
      rowAssocs.forEach(pair => {
        if (!parentSection.geAssociations.some(p => p.pattern === pair.pattern && p.area === pair.area)) {
          parentSection.geAssociations.push(pair);
          if (!parentSection.generalEdPattern.includes(pair.pattern)) {
            parentSection.generalEdPattern.push(pair.pattern);
          }
          if (!parentSection.generalEdArea.includes(pair.area)) {
            parentSection.generalEdArea.push(pair.area);
          }
          if (parentSection.gePatternAreaMap[pair.pattern]) {
            parentSection.gePatternAreaMap[pair.pattern].add(pair.area);
          }
        }
      });
    }

    const startMinutes = parseTimeToMinutes(row['Schedule - Start Time']);
    const endMinutes = parseTimeToMinutes(row['Schedule - End Time']);
    const daysArray = parseScheduleDays(row['Schedule - Days']);
    const bldgNorm = normalizeBuilding(row['Schedule - Building'], row['Schedule - Building Description']);

    const meetingRow = {
      rowId: `row_${rowIndex}`,
      instructionType: String(row['Schedule - Instruction Type'] || 'Instruction').trim(),
      building: bldgNorm.building,
      buildingDesc: bldgNorm.buildingDesc,
      room: String(row['Schedule - Room'] || '').trim(),
      startTimeRaw: row['Schedule - Start Time'],
      endTimeRaw: row['Schedule - End Time'],
      startTimeStr: formatMinutesToAmPm(startMinutes) || String(row['Schedule - Start Time'] || ''),
      endTimeStr: formatMinutesToAmPm(endMinutes) || String(row['Schedule - End Time'] || ''),
      startMinutes: startMinutes,
      endMinutes: endMinutes,
      daysRaw: row['Schedule - Days'],
      days: daysArray,
      startDate: row['Schedule - Start Date'] || '',
      endDate: row['Schedule - End Date'] || '',
      parentSection: parentSection
    };

    parentSection.scheduleBlocks.push(meetingRow);

    if (daysArray.length > 0 && startMinutes !== null && endMinutes !== null) {
      daysArray.forEach(dayNumber => {
        allMeetingBlocks.push({
          blockId: `block_${rowIndex}_d${dayNumber}`,
          day: dayNumber,
          days: daysArray,
          startMinutes: startMinutes,
          endMinutes: endMinutes,
          duration: Math.max(endMinutes - startMinutes, 20),
          instructionType: meetingRow.instructionType,
          building: meetingRow.building,
          buildingDesc: meetingRow.buildingDesc,
          room: meetingRow.room,
          startTimeStr: meetingRow.startTimeStr,
          endTimeStr: meetingRow.endTimeStr,
          startDate: meetingRow.startDate,
          endDate: meetingRow.endDate,
          parentSection: parentSection,
          meetingRow: meetingRow,
          isConflict: false,
          conflictsWith: [],
          isStacked: false,
          stackedWith: []
        });
      });
    }
  });

  appState.sections = Array.from(sectionMap.values());
  appState.filteredSections = appState.sections;
  appState.scheduleBlocks = allMeetingBlocks;

  const bldgDescMap = new Map();
  allMeetingBlocks.forEach(b => {
    if (b.building && b.buildingDesc && !bldgDescMap.has(b.building)) {
      bldgDescMap.set(b.building, b.buildingDesc);
    }
  });
  appState.buildingDescMap = bldgDescMap;

  populateFilterControls();
  applyFiltersAndRender();
}

// --- 5. CONFLICT & STACKED CLASSES ENGINE ---

/**
 * Known specialized programs that stack classes across cohorts, even with different instructors:
 * - AMT (Aviation Maintenance Technology)
 * - NURADN (Nursing ADN)
 * - DENTAL (Dental Assisting / Dental Hygiene)
 * - RADTEC (Radiologic Technology)
 * If overlapping classes belong to the SAME program from this list, mark them as Stacked rather than Conflict.
 * If subjects differ (e.g. DENTAL vs RADTEC), it remains a true Room Conflict.
 */
const MULTI_INSTRUCTOR_STACK_PROGRAMS = ['AMT', 'NURADN', 'DENTAL', 'RADTEC'];

function getSectionStackSubject(sec) {
  if (!sec) return '';
  const subj = String(sec.subject || '').trim().toUpperCase();
  const crse = String(sec.course || '').trim().toUpperCase();

  for (const prog of MULTI_INSTRUCTOR_STACK_PROGRAMS) {
    if (subj === prog || subj.startsWith(prog) || subj.includes(prog)) return prog;
    if (crse === prog || crse.startsWith(prog)) return prog;
  }

  // Common descriptive names in schedule data
  if (subj.includes('AVIATION') || subj.includes('AERONAUT') || crse.startsWith('AMT')) return 'AMT';
  if (subj.includes('NURADN') || (subj.includes('NURSING') && subj.includes('ADN')) || crse.startsWith('NURADN')) return 'NURADN';
  if (subj.includes('DENTAL') || crse.startsWith('DENTAL')) return 'DENTAL';
  if (subj.includes('RADTEC') || subj.includes('RADIOLOG') || crse.startsWith('RADTEC')) return 'RADTEC';

  return '';
}

/**
 * Helper to parse academic date strings (e.g., 'Aug 17, 2026') into epoch timestamps.
 */
function parseDateToTimestamp(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) return dateVal.getTime();
  const str = String(dateVal).trim();
  if (!str) return null;
  const ts = Date.parse(str);
  return isNaN(ts) ? null : ts;
}

/**
 * Detects whether two overlapping meeting blocks represent a Fast-track 1 and Fast-track 2 pairing
 * where Fast-track 1 ends before Fast-track 2 begins (or non-overlapping date spans).
 * These are sequential courses that share a room across the term without conflicting.
 */
function isSequentialFastTrackPair(blockA, blockB) {
  const secA = blockA?.parentSection;
  const secB = blockB?.parentSection;
  if (!secA || !secB) return false;

  const sessA = String(secA.calendarSession || '').toLowerCase();
  const sessB = String(secB.calendarSession || '').toLowerCase();

  const isA_FT1 = sessA.includes('fast-track 1') || sessA.includes('fast track 1') || sessA === 'ft1';
  const isA_FT2 = sessA.includes('fast-track 2') || sessA.includes('fast track 2') || sessA === 'ft2';
  const isB_FT1 = sessB.includes('fast-track 1') || sessB.includes('fast track 1') || sessB === 'ft1';
  const isB_FT2 = sessB.includes('fast-track 2') || sessB.includes('fast track 2') || sessB === 'ft2';

  let ft1Block = null;
  let ft2Block = null;

  if (isA_FT1 && isB_FT2) {
    ft1Block = blockA;
    ft2Block = blockB;
  } else if (isA_FT2 && isB_FT1) {
    ft1Block = blockB;
    ft2Block = blockA;
  } else {
    return false;
  }

  // Fast-track 1 must end on or before Fast-track 2 begins
  const end1 = parseDateToTimestamp(ft1Block.endDate || ft1Block.meetingRow?.endDate);
  const start2 = parseDateToTimestamp(ft2Block.startDate || ft2Block.meetingRow?.startDate);

  if (end1 !== null && start2 !== null) {
    return end1 <= start2;
  }

  // If date strings are unavailable, FT1 and FT2 are inherently sequential
  return true;
}

/**
 * Evaluates Room Scheduling Conflicts & Stacked Classes.
 * - STACKED CLASSES: Overlapping classes with the SAME instructor, classes in the same cohort
 *   program (AMT, NURADN, DENTAL, RADTEC), OR sequential Fast-Track 1 & Fast-Track 2 pairings.
 * - REAL ROOM CONFLICTS: Overlapping classes with DIFFERENT instructors / DIFFERENT programs.
 */
function evaluateRoomConflicts(blocks) {
  blocks.forEach(b => {
    b.isConflict = false;
    b.conflictsWith = [];
    b.isStacked = false;
    b.stackedWith = [];
  });

  const conflictPairs = [];
  const stackedPairs = [];

  for (let i = 0; i < blocks.length; i++) {
    const a = blocks[i];
    if (!a.building || !a.room) continue;

    for (let j = i + 1; j < blocks.length; j++) {
      const b = blocks[j];
      if (!b.building || !b.room) continue;

      if (a.parentSection.sectionUniqueId === b.parentSection.sectionUniqueId) continue;
      if (a.day !== b.day) continue;

      if (a.building.toUpperCase() !== b.building.toUpperCase()) continue;
      if (a.room.toUpperCase() !== b.room.toUpperCase()) continue;

      const hasTimeOverlap = a.startMinutes < b.endMinutes && b.startMinutes < a.endMinutes;

      if (hasTimeOverlap) {
        const facA = normalizeFacultyName(a.parentSection.facultyName);
        const facB = normalizeFacultyName(b.parentSection.facultyName);
        const isSameInstructor = facA && facB && facA === facB && facA !== 'staff' && facA !== 'tba';

        const stackProgA = getSectionStackSubject(a.parentSection);
        const stackProgB = getSectionStackSubject(b.parentSection);

        // Check if both classes share the same program known to stack across instructors (AMT, NURADN, DENTAL, RADTEC)
        const isSameProgramMultiStack = stackProgA && stackProgB && (stackProgA === stackProgB);

        // Check if this is a Fast-Track 1 and Fast-Track 2 pairing where FT1 ends before FT2 begins
        const isFastTrackSeq = isSequentialFastTrackPair(a, b);

        const isStacked = isSameInstructor || isSameProgramMultiStack || isFastTrackSeq;

        if (isStacked) {
          // STACKED CLASS: Same room, day, time, and either same faculty, same specialized program, OR sequential FT1/FT2!
          a.isStacked = true;
          b.isStacked = true;
          if (!a.stackedWith.includes(b)) a.stackedWith.push(b);
          if (!b.stackedWith.includes(a)) b.stackedWith.push(a);

          let stackReason = '';
          if (isFastTrackSeq) {
            stackReason = `Fast-Track Sequential Session (FT1 & FT2 - Non-overlapping Dates)`;
          } else if (isSameInstructor) {
            stackReason = `Same Instructor: ${a.parentSection.facultyName}`;
          } else {
            stackReason = `Cohort Program Stack: ${stackProgA} (Multi-Instructor)`;
          }

          stackedPairs.push({
            day: a.day,
            building: a.building,
            room: a.room,
            faculty: a.parentSection.facultyName,
            facultyB: b.parentSection.facultyName,
            program: stackProgA || '',
            reason: stackReason,
            isMultiInstructor: isSameProgramMultiStack && !isSameInstructor,
            isFastTrack: isFastTrackSeq,
            blockA: a,
            blockB: b,
            overlapStart: Math.max(a.startMinutes, b.startMinutes),
            overlapEnd: Math.min(a.endMinutes, b.endMinutes)
          });
        } else {
          // REAL ROOM CONFLICT: Different instructors (or different programs) in the same room at the same time!
          a.isConflict = true;
          b.isConflict = true;
          if (!a.conflictsWith.includes(b)) a.conflictsWith.push(b);
          if (!b.conflictsWith.includes(a)) b.conflictsWith.push(a);

          conflictPairs.push({
            day: a.day,
            building: a.building,
            room: a.room,
            blockA: a,
            blockB: b,
            overlapStart: Math.max(a.startMinutes, b.startMinutes),
            overlapEnd: Math.min(a.endMinutes, b.endMinutes)
          });
        }
      }
    }
  }

  appState.conflicts = conflictPairs;
  appState.stackedPairs = stackedPairs;
  updateConflictUI();
  return conflictPairs;
}

function updateConflictUI() {
  const conflictCount = appState.conflicts.length;
  const stackedCount = appState.stackedPairs.length;

  const statConflicts = document.getElementById('stat-conflicts');
  const statStacked = document.getElementById('stat-stacked');
  const conflictPill = document.getElementById('metric-conflicts-pill');
  const stackedPill = document.getElementById('metric-stacked-pill');
  const controlsConflictBadge = document.getElementById('controls-conflict-badge');
  const banner = document.getElementById('conflict-alert-banner');
  const bannerDesc = document.getElementById('conflict-banner-desc');
  const switcher = document.getElementById('conflict-mode-switcher');

  if (statConflicts) statConflicts.textContent = conflictCount;
  if (controlsConflictBadge) controlsConflictBadge.textContent = conflictCount;
  if (statStacked) statStacked.textContent = stackedCount;

  if (conflictCount > 0) {
    conflictPill?.classList.add('has-conflicts');
    if (appState.isConflictMode) {
      if (banner) banner.style.display = 'flex';
      if (bannerDesc) bannerDesc.textContent = `${conflictCount} room conflict${conflictCount > 1 ? 's' : ''} detected across schedule.`;
      if (switcher) switcher.style.display = 'inline-flex';
    } else {
      if (banner) banner.style.display = 'none';
      if (switcher) switcher.style.display = 'none';
    }
  } else {
    conflictPill?.classList.remove('has-conflicts');
    if (banner) banner.style.display = 'none';
    if (!appState.isConflictMode && switcher) switcher.style.display = 'none';
  }

  if (stackedCount > 0) {
    stackedPill?.classList.add('has-stacked');
  } else {
    stackedPill?.classList.remove('has-stacked');
  }
}

// --- 6. FILTERING LOGIC ---

/**
 * Checks if a section satisfies all active filters, optionally excluding one dimension.
 * This leave-one-out evaluation powers faceted cross-filtering across all filter dimensions.
 */
function matchesOtherFilters(sec, excludeDimension = null) {
  const {
    search, location, building, room,
    dayEvening, geArea, gePattern, acc, subject, course,
    modality, session, traits
  } = appState.filters;

  if (search) {
    const searchLower = search.toLowerCase().trim();
    const combined = `${sec.sectionName} ${sec.course} ${sec.subject} ${sec.facultyName} ${sec.location} ${sec.scheduleBlocks.map(b => `${b.building} ${b.room}`).join(' ')}`.toLowerCase();
    if (!combined.includes(searchLower)) return false;
  }

  if (excludeDimension !== 'location' && location) {
    if (sec.location.toLowerCase() !== location.toLowerCase()) return false;
  }

  if (excludeDimension !== 'dayEvening' && dayEvening.size > 0) {
    if (!dayEvening.has(sec.dayOrEvening)) return false;
  }

  if (excludeDimension !== 'acc' && acc.size > 0) {
    if (!acc.has(sec.acc)) return false;
  }

  if (excludeDimension !== 'subject' && subject.size > 0) {
    if (!subject.has(sec.subject)) return false;
  }

  if (excludeDimension !== 'course' && course.size > 0) {
    if (!course.has(sec.course)) return false;
  }

  if (excludeDimension !== 'modality' && modality.size > 0) {
    if (!modality.has(sec.modality)) return false;
  }

  if (excludeDimension !== 'session' && session.size > 0) {
    if (!session.has(sec.calendarSession)) return false;
  }

  // Respect multivalued associated GE Pattern & GE Area pairing
  if (excludeDimension !== 'gePattern' && excludeDimension !== 'geArea') {
    if (gePattern.size > 0 && geArea.size > 0) {
      const hasMatchingPair = (sec.geAssociations || []).some(pair =>
        gePattern.has(pair.pattern) && geArea.has(pair.area)
      );
      if (!hasMatchingPair) return false;
    } else if (gePattern.size > 0) {
      const hasGePat = (sec.geAssociations || []).some(pair => gePattern.has(pair.pattern));
      if (!hasGePat) return false;
    } else if (geArea.size > 0) {
      const hasGeArea = (sec.geAssociations || []).some(pair => geArea.has(pair.area));
      if (!hasGeArea) return false;
    }
  } else if (excludeDimension === 'gePattern') {
    if (geArea.size > 0) {
      const hasGeArea = (sec.geAssociations || []).some(pair => geArea.has(pair.area));
      if (!hasGeArea) return false;
    }
  } else if (excludeDimension === 'geArea') {
    if (gePattern.size > 0) {
      const hasGePat = (sec.geAssociations || []).some(pair => gePattern.has(pair.pattern));
      if (!hasGePat) return false;
    }
  }

  if (excludeDimension !== 'traits' && traits.size > 0) {
    const hasTrait = Array.from(traits).some(t => sec.sectionTraits.includes(t));
    if (!hasTrait) return false;
  }

  const checkBuilding = (excludeDimension !== 'building' && excludeDimension !== 'location') && Boolean(building);
  const checkRoom = (excludeDimension !== 'room' && excludeDimension !== 'building' && excludeDimension !== 'location') && Boolean(room);

  if (checkBuilding || checkRoom) {
    const hasMatchingBlock = sec.scheduleBlocks.some(b => {
      if (checkBuilding && b.building.toLowerCase() !== building.toLowerCase()) return false;
      if (checkRoom && b.room.toLowerCase() !== room.toLowerCase()) return false;
      return true;
    });
    if (!hasMatchingBlock) return false;
  }

  return true;
}

/**
 * Returns the set of valid options for a specific dimension based on sections that match
 * all other active filters (leave-one-out).
 */
function getValidFacetValues(dimension) {
  const valid = new Set();
  const sections = appState.sections;
  const areaFilter = appState.filters.geArea;
  const patFilter = appState.filters.gePattern;

  for (let i = 0; i < sections.length; i++) {
    const sec = sections[i];
    if (!matchesOtherFilters(sec, dimension)) continue;

    switch (dimension) {
      case 'location':
        if (sec.location) valid.add(sec.location);
        break;
      case 'acc':
        if (sec.acc) valid.add(sec.acc);
        break;
      case 'subject':
        if (sec.subject) valid.add(sec.subject);
        break;
      case 'course':
        if (sec.course) valid.add(sec.course);
        break;
      case 'modality':
        if (sec.modality) valid.add(sec.modality);
        break;
      case 'session':
        if (sec.calendarSession) valid.add(sec.calendarSession);
        break;
      case 'dayEvening':
        if (sec.dayOrEvening) valid.add(sec.dayOrEvening);
        break;
      case 'traits':
        if (sec.sectionTraits) sec.sectionTraits.forEach(t => valid.add(t));
        break;
      case 'gePattern':
        if (sec.geAssociations) {
          sec.geAssociations.forEach(pair => {
            if (areaFilter.size === 0 || areaFilter.has(pair.area)) {
              valid.add(pair.pattern);
            }
          });
        }
        break;
      case 'geArea':
        if (sec.geAssociations) {
          sec.geAssociations.forEach(pair => {
            if (patFilter.size === 0 || patFilter.has(pair.pattern)) {
              valid.add(pair.area);
            }
          });
        }
        break;
    }
  }
  return valid;
}

/**
 * Prunes any values from a filter Set that are no longer valid under active cross-filters.
 */
function pruneInvalidSelections(filterSet, validSet) {
  let pruned = false;
  Array.from(filterSet).forEach(item => {
    if (!validSet.has(item)) {
      filterSet.delete(item);
      pruned = true;
    }
  });
  return pruned;
}

let isPopulatingFilters = false;

function populateFilterControls() {
  if (isPopulatingFilters) return;
  isPopulatingFilters = true;

  try {
    let hadPruning = false;

    const locations = getValidFacetValues('location');
    const accs = getValidFacetValues('acc');
    const subjects = getValidFacetValues('subject');
    const courses = getValidFacetValues('course');
    const modalities = getValidFacetValues('modality');
    const sessions = getValidFacetValues('session');
    const dayEvenings = getValidFacetValues('dayEvening');
    const gePatterns = getValidFacetValues('gePattern');
    const geAreas = getValidFacetValues('geArea');
    const traits = getValidFacetValues('traits');

    // Automatically prune invalid selections in multi-select dimensions
    if (pruneInvalidSelections(appState.filters.acc, accs)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.subject, subjects)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.course, courses)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.modality, modalities)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.session, sessions)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.dayEvening, dayEvenings)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.gePattern, gePatterns)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.geArea, geAreas)) hadPruning = true;
    if (pruneInvalidSelections(appState.filters.traits, traits)) hadPruning = true;

    // Location Dropdown
    const locSelect = document.getElementById('filter-location');
    if (locSelect) {
      if (appState.filters.location && !locations.has(appState.filters.location)) {
        appState.filters.location = '';
        hadPruning = true;
      }
      const currentLoc = appState.filters.location;
      const sortedLocs = Array.from(locations).sort();
      const locKey = sortedLocs.join('|') + '::' + currentLoc;
      if (locSelect._lastKey !== locKey) {
        locSelect._lastKey = locKey;
        locSelect.innerHTML = '<option value="">All Locations</option>';
        sortedLocs.forEach(loc => {
          const opt = document.createElement('option');
          opt.value = loc;
          opt.textContent = loc;
          if (loc === currentLoc) opt.selected = true;
          locSelect.appendChild(opt);
        });
      }
      const countLoc = document.getElementById('count-location');
      if (countLoc) countLoc.textContent = `(${locations.size})`;
    }

    if (updateCascadingBuildingAndRoom()) {
      hadPruning = true;
    }

    renderCheckboxList('list-ge-area', sortGeAreas(geAreas), appState.filters.geArea, 'badge-ge-area', formatGeAreaLabel);
    renderCheckboxList('list-ge-pattern', Array.from(gePatterns).sort(), appState.filters.gePattern, 'badge-ge-pattern', formatGePatternLabel);
    renderCheckboxList('list-acc', sortAccs(accs), appState.filters.acc, 'badge-acc');
    renderCheckboxList('list-subject', Array.from(subjects).sort(), appState.filters.subject, 'badge-subject', null, 'filter-subject-search');
    renderCheckboxList('list-course', Array.from(courses).sort(), appState.filters.course, 'badge-course', null, 'filter-course-search');
    renderCheckboxList('list-modality', Array.from(modalities).sort(), appState.filters.modality, 'badge-modality');
    renderCheckboxList('list-session', Array.from(sessions).sort(), appState.filters.session, 'badge-session');
    renderCheckboxList('list-day-evening', Array.from(dayEvenings).sort(), appState.filters.dayEvening, 'badge-day-evening', formatDayEveningLabel);
    renderCheckboxList('list-traits', Array.from(traits).sort(), appState.filters.traits, 'badge-traits', null, 'filter-traits-search');

    if (hadPruning) {
      applyFiltersAndRender();
    }
  } finally {
    isPopulatingFilters = false;
  }
}

function updateCascadingBuildingAndRoom() {
  let hadPruning = false;
  const selectedBuilding = appState.filters.building;

  const validBuildings = new Set();
  const bldgDescMap = appState.buildingDescMap || new Map();
  const validRooms = new Set();

  appState.sections.forEach(sec => {
    if (matchesOtherFilters(sec, 'building')) {
      sec.scheduleBlocks.forEach(b => {
        if (b.building) {
          validBuildings.add(b.building);
          if (b.buildingDesc && !bldgDescMap.has(b.building)) {
            bldgDescMap.set(b.building, b.buildingDesc);
          }
        }
      });
    }

    if (matchesOtherFilters(sec, 'room')) {
      sec.scheduleBlocks.forEach(b => {
        if (!selectedBuilding || b.building.toLowerCase() === selectedBuilding.toLowerCase()) {
          if (b.room) validRooms.add(b.room);
        }
      });
    }
  });

  appState.buildingDescMap = bldgDescMap;

  const bldgSelect = document.getElementById('filter-building');
  if (bldgSelect) {
    if (appState.filters.building && !validBuildings.has(appState.filters.building)) {
      appState.filters.building = '';
      hadPruning = true;
    }
    const currentBldg = appState.filters.building;
    const sortedBldgs = Array.from(validBuildings).sort();
    const bldgKey = sortedBldgs.join('|') + '::' + currentBldg;
    if (bldgSelect._lastKey !== bldgKey) {
      bldgSelect._lastKey = bldgKey;
      bldgSelect.innerHTML = '<option value="">All Buildings</option>';
      sortedBldgs.forEach(bldg => {
        const opt = document.createElement('option');
        opt.value = bldg;
        const desc = bldgDescMap.get(bldg);
        opt.textContent = formatBuildingDropdownLabel(bldg, desc);
        if (bldg === currentBldg) opt.selected = true;
        bldgSelect.appendChild(opt);
      });
    }
    const countBldg = document.getElementById('count-building');
    if (countBldg) countBldg.textContent = `(${validBuildings.size})`;
  }

  const roomSelect = document.getElementById('filter-room');
  if (roomSelect) {
    if (appState.filters.room && !validRooms.has(appState.filters.room)) {
      appState.filters.room = '';
      hadPruning = true;
    }
    const currentRoom = appState.filters.room;
    const sortedRooms = Array.from(validRooms).sort();
    const roomKey = sortedRooms.join('|') + '::' + currentRoom;
    if (roomSelect._lastKey !== roomKey) {
      roomSelect._lastKey = roomKey;
      roomSelect.innerHTML = '<option value="">All Rooms</option>';
      sortedRooms.forEach(room => {
        const opt = document.createElement('option');
        opt.value = room;
        opt.textContent = room;
        if (room === currentRoom) opt.selected = true;
        roomSelect.appendChild(opt);
      });
    }
    const countRoom = document.getElementById('count-room');
    if (countRoom) countRoom.textContent = `(${validRooms.size})`;
  }

  return hadPruning;
}

function renderCheckboxList(containerId, items, selectedSet, badgeId, labelFormatter = null, searchInputId = null) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const searchInput = searchInputId ? document.getElementById(searchInputId) : null;
  const searchQuery = searchInput ? searchInput.value.toLowerCase().trim() : '';

  const renderKey = items.join('|') + '::' + Array.from(selectedSet).sort().join(',') + '::' + searchQuery;
  if (container._renderKey === renderKey) {
    updateBadge(badgeId, selectedSet.size);
    return;
  }
  container._renderKey = renderKey;

  const prevScrollTop = container.scrollTop;
  container.innerHTML = '';

  items.forEach(item => {
    const label = document.createElement('label');
    label.className = 'checkbox-label';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = item;
    checkbox.checked = selectedSet.has(item);

    checkbox.addEventListener('change', () => {
      if (checkbox.checked) {
        selectedSet.add(item);
      } else {
        selectedSet.delete(item);
      }
      updateBadge(badgeId, selectedSet.size);
      applyFiltersAndRender();
    });

    const span = document.createElement('span');
    const labelText = labelFormatter ? labelFormatter(item) : item;
    span.textContent = labelText;

    label.appendChild(checkbox);
    label.appendChild(span);

    if (searchQuery && !labelText.toLowerCase().includes(searchQuery)) {
      label.style.display = 'none';
    }

    container.appendChild(label);
  });

  container.scrollTop = prevScrollTop;
  updateBadge(badgeId, selectedSet.size);
}

function updateBadge(badgeId, count) {
  const badge = document.getElementById(badgeId);
  if (badge) {
    badge.textContent = count;
    if (count > 0) badge.classList.add('active');
    else badge.classList.remove('active');
  }
}

function applyFiltersAndRender() {
  const { building, room } = appState.filters;

  // Filter Parent Sections using matchesOtherFilters with no excluded dimension
  let matchedSections = appState.sections.filter(sec => matchesOtherFilters(sec, null));

  const matchedSectionIds = new Set(matchedSections.map(s => s.sectionUniqueId));

  // Filter Meeting Blocks
  let filteredBlocks = appState.scheduleBlocks.filter(block => {
    if (!matchedSectionIds.has(block.parentSection.sectionUniqueId)) return false;
    if (building && block.building.toLowerCase() !== building.toLowerCase()) return false;
    if (room && block.room.toLowerCase() !== room.toLowerCase()) return false;
    return true;
  });

  // Evaluate Conflicts and Stacked Classes across filtered set
  evaluateRoomConflicts(filteredBlocks);

  // If in GE Coverage view, show only sections belonging to the active GE audit pattern
  if (appState.layoutMode === 'ge') {
    const activePat = appState.geActivePattern || 'CALGETC';
    filteredBlocks = filteredBlocks.filter(b =>
      b.parentSection.geAssociations &&
      b.parentSection.geAssociations.some(pair => pair.pattern === activePat)
    );
  }

  // If in Evening Coverage view, show only Evening/Weekend classes
  if (appState.layoutMode === 'evening') {
    filteredBlocks = filteredBlocks.filter(isEveningBlock);
  }

  // If in Fast-Track Coverage view, analyze physical Fast-Track 1 & 2 offerings and slice by active tab
  if (appState.layoutMode === 'fasttrack') {
    const ftScopeBlocks = filteredBlocks.filter(b => isFastTrackBlock(b) && isPhysicalMeetingBlock(b));
    const ftAnalysis = analyzeFastTrackOfferings(matchedSections, ftScopeBlocks);
    appState.currentFtAnalysis = ftAnalysis;

    const activeTab = appState.fasttrackActiveTab || 'all';
    if (activeTab === 'optimized') {
      const pairedRowIds = new Set();
      const tabSecIds = new Set();
      ftAnalysis.optPairs.forEach(p => {
        if (p.block1 && p.block1.rowId) pairedRowIds.add(p.block1.rowId);
        if (p.block2 && p.block2.rowId) pairedRowIds.add(p.block2.rowId);
        if (p.sec1 && p.sec1.sectionUniqueId) tabSecIds.add(p.sec1.sectionUniqueId);
        if (p.sec2 && p.sec2.sectionUniqueId) tabSecIds.add(p.sec2.sectionUniqueId);
      });
      filteredBlocks = ftScopeBlocks.filter(b => {
        const rowId = (b.meetingRow && b.meetingRow.rowId) || b.rowId;
        return pairedRowIds.has(rowId);
      });
      matchedSections = matchedSections.filter(s => tabSecIds.has(s.sectionUniqueId));
    } else if (activeTab === 'unoptimized') {
      const pairedRowIds = new Set();
      const tabSecIds = new Set();
      ftAnalysis.unoptPairs.forEach(p => {
        if (p.block1 && p.block1.rowId) pairedRowIds.add(p.block1.rowId);
        if (p.block2 && p.block2.rowId) pairedRowIds.add(p.block2.rowId);
        if (p.sec1 && p.sec1.sectionUniqueId) tabSecIds.add(p.sec1.sectionUniqueId);
        if (p.sec2 && p.sec2.sectionUniqueId) tabSecIds.add(p.sec2.sectionUniqueId);
      });
      filteredBlocks = ftScopeBlocks.filter(b => {
        const rowId = (b.meetingRow && b.meetingRow.rowId) || b.rowId;
        return pairedRowIds.has(rowId);
      });
      matchedSections = matchedSections.filter(s => tabSecIds.has(s.sectionUniqueId));
    } else if (activeTab === 'unpaired') {
      const unpairedRowIds = new Set();
      const tabSecIds = new Set();
      ftAnalysis.unpairedList.forEach(u => {
        if (u.block && u.block.rowId) unpairedRowIds.add(u.block.rowId);
        if (u.section && u.section.sectionUniqueId) tabSecIds.add(u.section.sectionUniqueId);
      });
      filteredBlocks = ftScopeBlocks.filter(b => {
        const rowId = (b.meetingRow && b.meetingRow.rowId) || b.rowId;
        return unpairedRowIds.has(rowId);
      });
      matchedSections = matchedSections.filter(s => tabSecIds.has(s.sectionUniqueId));
    } else {
      // 'all' tab: all physical Fast-Track blocks in current scope
      filteredBlocks = ftScopeBlocks;
      const ftSecIds = new Set();
      ftAnalysis.pairs.forEach(p => {
        if (p.sec1 && p.sec1.sectionUniqueId) ftSecIds.add(p.sec1.sectionUniqueId);
        if (p.sec2 && p.sec2.sectionUniqueId) ftSecIds.add(p.sec2.sectionUniqueId);
      });
      ftAnalysis.unpairedList.forEach(u => {
        if (u.section && u.section.sectionUniqueId) ftSecIds.add(u.section.sectionUniqueId);
      });
      matchedSections = matchedSections.filter(s => ftSecIds.has(s.sectionUniqueId));
    }
  }

  // If in Efficiency Audit view, analyze enrollment efficiency and slice by active tab
  if (appState.layoutMode === 'efficiency') {
    const effAnalysis = analyzeEnrollmentEfficiency(matchedSections, filteredBlocks);
    appState.currentEfficiencyAnalysis = effAnalysis;

    const activeTab = appState.efficiencyActiveTab || 'all';
    let tabSecs = matchedSections;
    if (activeTab === 'low_enrolled') {
      tabSecs = effAnalysis.lowEnrolledSections;
    } else if (activeTab === 'healthy') {
      tabSecs = effAnalysis.healthySections;
    } else if (activeTab === 'critical') {
      tabSecs = effAnalysis.criticalSections;
    } else if (activeTab === 'high_lost_ftes') {
      tabSecs = effAnalysis.highLostFtesSections;
    }

    if (activeTab !== 'all') {
      const tabSecIdSet = new Set(tabSecs.map(s => s.sectionUniqueId));
      filteredBlocks = filteredBlocks.filter(b => tabSecIdSet.has(b.parentSection.sectionUniqueId));
      matchedSections = tabSecs;
    }
  }

  // REQUIREMENT 1: If Conflict Mode is active and "Isolate Conflicts Only" is on, isolate conflicts!
  if (appState.isConflictMode && appState.conflictsOnly) {
    filteredBlocks = filteredBlocks.filter(b => b.isConflict);
  }

  appState.filteredBlocks = filteredBlocks;
  appState.filteredSections = matchedSections;

  document.getElementById('stat-total-sections').textContent = appState.sections.length;
  document.getElementById('stat-total-meetings').textContent = appState.scheduleBlocks.length;
  document.getElementById('status-sections-count').textContent = matchedSections.length;
  document.getElementById('status-meetings-count').textContent = filteredBlocks.length;

  renderActiveFilterPills();
  dispatchViewRender();

  // Dynamically update cascading filter controls
  populateFilterControls();
}

function renderActiveFilterPills() {
  const container = document.getElementById('active-pills-list');
  const bar = document.getElementById('active-filters-bar');
  if (!container || !bar) return;

  container.innerHTML = '';
  let activeCount = 0;

  function addPill(label, value, onRemove) {
    activeCount++;
    const pill = document.createElement('div');
    pill.className = 'filter-pill';
    pill.innerHTML = `<span>${label}: <strong>${value}</strong></span><span class="pill-remove">&times;</span>`;
    pill.querySelector('.pill-remove').addEventListener('click', onRemove);
    container.appendChild(pill);
  }

  if (appState.filters.search) {
    addPill('Search', appState.filters.search, () => {
      appState.filters.search = '';
      document.getElementById('filter-search').value = '';
      applyFiltersAndRender();
    });
  }

  if (appState.filters.location) {
    addPill('Location', appState.filters.location, () => {
      appState.filters.location = '';
      document.getElementById('filter-location').value = '';
      updateCascadingBuildingAndRoom();
      applyFiltersAndRender();
    });
  }

  if (appState.filters.building) {
    const bldgCode = appState.filters.building;
    const bldgDesc = appState.buildingDescMap?.get(bldgCode);
    const bldgLabel = formatBuildingDropdownLabel(bldgCode, bldgDesc);
    addPill('Building', bldgLabel, () => {
      appState.filters.building = '';
      document.getElementById('filter-building').value = '';
      updateCascadingBuildingAndRoom();
      applyFiltersAndRender();
    });
  }

  if (appState.filters.room) {
    addPill('Room', appState.filters.room, () => {
      appState.filters.room = '';
      document.getElementById('filter-room').value = '';
      applyFiltersAndRender();
    });
  }

  appState.filters.dayEvening.forEach(val => {
    addPill('Day/Evening', formatDayEveningLabel(val), () => {
      appState.filters.dayEvening.delete(val);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.geArea.forEach(area => {
    addPill('GE Area', `Area ${area}`, () => {
      appState.filters.geArea.delete(area);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.gePattern.forEach(pat => {
    addPill('GE Pattern', formatGePatternLabel(pat), () => {
      appState.filters.gePattern.delete(pat);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.acc.forEach(val => {
    addPill('ACC', val, () => {
      appState.filters.acc.delete(val);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.subject.forEach(val => {
    addPill('Subject', val, () => {
      appState.filters.subject.delete(val);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.course.forEach(val => {
    addPill('Course', val, () => {
      appState.filters.course.delete(val);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.modality.forEach(val => {
    addPill('Modality', val, () => {
      appState.filters.modality.delete(val);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.session.forEach(val => {
    addPill('Session', val, () => {
      appState.filters.session.delete(val);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.traits.forEach(trait => {
    addPill('Trait', trait, () => {
      appState.filters.traits.delete(trait);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  bar.style.display = activeCount > 0 ? 'flex' : 'none';
}

// --- 7. SCHEDULE GRID & VIEW MODE DISPATCHER ---

/**
 * Dispatches view rendering based on active layout mode:
 * - 'day': Standard Day View calendar with side-by-side tiles
 * - 'room': Stacked Room View (vertically stacked grids by Building/Room)
 * - 'ge': General Education Coverage Audit Dashboard & GE Schedule (swappable to Day or Room View)
 * - 'evening': Evening & Weekend Coverage Audit Dashboard & Schedule (swappable to Day or Room View)
 * - 'fasttrack': Fast-Track Coverage Audit Dashboard & Schedule (swappable to Day or Room View)
 */
function dispatchViewRender() {
  const welcomeHero = document.getElementById('welcome-upload-hero');
  const standardGrid = document.getElementById('schedule-grid');
  const stackedContainer = document.getElementById('room-matrix-stacked');
  const geContainer = document.getElementById('ge-coverage-container');
  const eveningContainer = document.getElementById('evening-coverage-container');
  const fasttrackContainer = document.getElementById('fasttrack-coverage-container');
  const efficiencyContainer = document.getElementById('efficiency-container');

  // If no sections are loaded yet, display prominent Step 1 upload hero
  if (appState.sections.length === 0) {
    if (welcomeHero) welcomeHero.style.display = 'flex';
    if (standardGrid) standardGrid.style.display = 'none';
    if (stackedContainer) stackedContainer.style.display = 'none';
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (fasttrackContainer) fasttrackContainer.style.display = 'none';
    if (efficiencyContainer) efficiencyContainer.style.display = 'none';
    return;
  }

  // Hide welcome hero once data is loaded
  if (welcomeHero) welcomeHero.style.display = 'none';

  if (appState.layoutMode === 'room') {
    if (standardGrid) standardGrid.style.display = 'none';
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (fasttrackContainer) fasttrackContainer.style.display = 'none';
    if (efficiencyContainer) efficiencyContainer.style.display = 'none';
    if (stackedContainer) {
      stackedContainer.style.display = 'flex';
      renderStackedRoomMatrix();
    }
  } else if (appState.layoutMode === 'ge') {
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (fasttrackContainer) fasttrackContainer.style.display = 'none';
    if (efficiencyContainer) efficiencyContainer.style.display = 'none';
    if (geContainer) {
      geContainer.style.display = 'flex';
      renderGeCoverageDashboard();
    }
    // Swap grid below coverage audit to either Day View or Room View
    if (appState.geSubLayout === 'room') {
      if (standardGrid) standardGrid.style.display = 'none';
      if (stackedContainer) {
        stackedContainer.style.display = 'flex';
        renderStackedRoomMatrix();
      }
    } else {
      if (stackedContainer) stackedContainer.style.display = 'none';
      if (standardGrid) {
        standardGrid.style.display = 'grid';
        renderScheduleGrid();
      }
    }
  } else if (appState.layoutMode === 'evening') {
    if (geContainer) geContainer.style.display = 'none';
    if (fasttrackContainer) fasttrackContainer.style.display = 'none';
    if (efficiencyContainer) efficiencyContainer.style.display = 'none';
    if (eveningContainer) {
      eveningContainer.style.display = 'flex';
      renderEveningCoverageDashboard();
    }
    // Swap grid below coverage audit to either Day View or Room View
    if (appState.eveningSubLayout === 'room') {
      if (standardGrid) standardGrid.style.display = 'none';
      if (stackedContainer) {
        stackedContainer.style.display = 'flex';
        renderStackedRoomMatrix();
      }
    } else {
      if (stackedContainer) stackedContainer.style.display = 'none';
      if (standardGrid) {
        standardGrid.style.display = 'grid';
        renderScheduleGrid();
      }
    }
  } else if (appState.layoutMode === 'fasttrack') {
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (efficiencyContainer) efficiencyContainer.style.display = 'none';
    if (fasttrackContainer) {
      fasttrackContainer.style.display = 'flex';
      renderFastTrackCoverageDashboard();
    }
    // Swap grid below coverage audit to either Day View or Room View
    if (appState.fasttrackSubLayout === 'room') {
      if (standardGrid) standardGrid.style.display = 'none';
      if (stackedContainer) {
        stackedContainer.style.display = 'flex';
        renderStackedRoomMatrix();
      }
    } else {
      if (stackedContainer) stackedContainer.style.display = 'none';
      if (standardGrid) {
        standardGrid.style.display = 'grid';
        renderScheduleGrid();
      }
    }
  } else if (appState.layoutMode === 'efficiency') {
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (fasttrackContainer) fasttrackContainer.style.display = 'none';
    if (efficiencyContainer) {
      efficiencyContainer.style.display = 'flex';
      renderEfficiencyDashboard();
    }
    // Swap grid below coverage audit to either Day View or Room View
    if (appState.efficiencySubLayout === 'room') {
      if (standardGrid) standardGrid.style.display = 'none';
      if (stackedContainer) {
        stackedContainer.style.display = 'flex';
        renderStackedRoomMatrix();
      }
    } else {
      if (stackedContainer) stackedContainer.style.display = 'none';
      if (standardGrid) {
        standardGrid.style.display = 'grid';
        renderScheduleGrid();
      }
    }
  } else {
    // Default 'day' view
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (fasttrackContainer) fasttrackContainer.style.display = 'none';
    if (efficiencyContainer) efficiencyContainer.style.display = 'none';
    if (stackedContainer) stackedContainer.style.display = 'none';
    if (standardGrid) {
      standardGrid.style.display = 'grid';
      renderScheduleGrid();
    }
  }
}

/**
 * Google Calendar Style Interval Graph Partitioning Algorithm:
 * Resolves overlapping tiles within a day column and assigns sub-column slots.
 */
function resolveTileOverlaps(dayBlocks) {
  if (dayBlocks.length === 0) return [];

  const sorted = [...dayBlocks].sort((a, b) => {
    if (a.startMinutes !== b.startMinutes) return a.startMinutes - b.startMinutes;
    return b.duration - a.duration;
  });

  const clusters = [];
  let currentCluster = [];
  let clusterEnd = -1;

  sorted.forEach(block => {
    if (currentCluster.length === 0) {
      currentCluster.push(block);
      clusterEnd = block.endMinutes;
    } else {
      if (block.startMinutes < clusterEnd) {
        currentCluster.push(block);
        clusterEnd = Math.max(clusterEnd, block.endMinutes);
      } else {
        clusters.push(currentCluster);
        currentCluster = [block];
        clusterEnd = block.endMinutes;
      }
    }
  });

  if (currentCluster.length > 0) {
    clusters.push(currentCluster);
  }

  clusters.forEach(cluster => {
    const columns = [];

    cluster.forEach(block => {
      let placedCol = -1;
      for (let i = 0; i < columns.length; i++) {
        const lastInCol = columns[i][columns[i].length - 1];
        if (block.startMinutes >= lastInCol.endMinutes) {
          columns[i].push(block);
          placedCol = i;
          break;
        }
      }

      if (placedCol === -1) {
        columns.push([block]);
        placedCol = columns.length - 1;
      }

      block.layoutCol = placedCol;
    });

    const totalCols = columns.length;
    cluster.forEach(block => {
      block.totalCols = totalCols;
    });
  });

  return sorted;
}

/**
 * Standard Day Calendar View (Single Grid with Smart Overlapping Column Placements)
 */
function renderScheduleGrid() {
  const gridContainer = document.getElementById('schedule-grid');
  if (!gridContainer) return;
  gridContainer.innerHTML = '';

  const blocks = appState.filteredBlocks;

  // Limit Day View to <= 200 blocks to prevent overcrowded congestion
  if (blocks.length > 200) {
    gridContainer.style.gridTemplateColumns = '1fr';
    gridContainer.style.gridTemplateRows = 'auto';
    gridContainer.innerHTML = `
      <div class="grid-limit-warning">
        <div class="limit-warning-icon">
          <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
        </div>
        <h3>Too Many Sections to Display in Day View</h3>
        <p>The current selection contains <strong>${blocks.length} scheduled meetings</strong>. Day View is limited to displaying up to 200 blocks to prevent visual clutter and overlapping congestion.</p>
        <p class="limit-warning-hint">Please update your filters (such as Location, Building, or Subject) to narrow down results, or switch to the <strong>Room View</strong> to inspect grids room-by-room.</p>
        <button type="button" class="btn btn-switch-room-matrix" id="btn-switch-to-room-matrix">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="21" x2="9" y2="9"/></svg>
          <span>Switch to Room View</span>
        </button>
      </div>
    `;

    document.getElementById('btn-switch-to-room-matrix')?.addEventListener('click', () => {
      if (appState.layoutMode === 'ge') {
        appState.geSubLayout = 'room';
      } else if (appState.layoutMode === 'evening') {
        appState.eveningSubLayout = 'room';
      } else if (appState.layoutMode === 'fasttrack') {
        appState.fasttrackSubLayout = 'room';
      } else if (appState.layoutMode === 'efficiency') {
        appState.efficiencySubLayout = 'room';
      } else {
        appState.layoutMode = 'room';
        const layoutGroup = document.getElementById('btn-group-layout');
        if (layoutGroup) {
          layoutGroup.querySelectorAll('.btn-toggle').forEach(b => {
            b.classList.toggle('active', b.getAttribute('data-layout') === 'room');
          });
        }
      }
      dispatchViewRender();
    });

    return;
  }

  let displayDays = [1, 2, 3, 4, 5];
  if (appState.daySpan === 6) displayDays = [1, 2, 3, 4, 5, 6];
  if (appState.daySpan === 7) displayDays = [1, 2, 3, 4, 5, 6, 7];
  if (appState.daySpan === 'weekend') displayDays = [5, 6, 7];
  if (appState.daySpan === 'active') {
    const activeDayNums = new Set(blocks.map(b => b.day));
    displayDays = [1, 2, 3, 4, 5, 6, 7].filter(d => activeDayNums.has(d));
    if (displayDays.length === 0) displayDays = [1, 2, 3, 4, 5];
  }

  let minMinutes = 7 * 60;
  let maxMinutes = 22 * 60;

  if (blocks.length > 0) {
    const earliestClass = Math.min(...blocks.map(b => b.startMinutes));
    const latestClass = Math.max(...blocks.map(b => b.endMinutes));
    minMinutes = Math.max(6 * 60, Math.floor((earliestClass - 30) / 60) * 60);
    maxMinutes = Math.min(23 * 60, Math.ceil((latestClass + 30) / 60) * 60);
  }

  const slotDuration = 30;
  const totalSlots = Math.ceil((maxMinutes - minMinutes) / slotDuration);
  const rowHeightPx = 32;

  const dayColCount = displayDays.length;
  gridContainer.style.gridTemplateColumns = `80px repeat(${dayColCount}, minmax(180px, 1fr))`;
  gridContainer.style.gridTemplateRows = `auto repeat(${totalSlots}, ${rowHeightPx}px)`;

  const cornerCell = document.createElement('div');
  cornerCell.className = 'grid-corner-cell';
  cornerCell.textContent = 'TIME';
  gridContainer.appendChild(cornerCell);

  displayDays.forEach((dayNum, dayIndex) => {
    const dayMeta = DAY_MAP[dayNum];
    const dayHeader = document.createElement('div');
    dayHeader.className = 'grid-day-header';
    dayHeader.style.gridColumn = `${dayIndex + 2}`;
    dayHeader.style.gridRow = '1';

    const dayBlocks = blocks.filter(b => b.day === dayNum);

    const nameSpan = document.createElement('span');
    nameSpan.className = 'day-name';
    nameSpan.textContent = dayMeta.name;

    const countSpan = document.createElement('span');
    countSpan.className = 'day-count';
    countSpan.textContent = `${dayBlocks.length} meetings`;

    dayHeader.appendChild(nameSpan);
    dayHeader.appendChild(countSpan);
    gridContainer.appendChild(dayHeader);
  });

  for (let slot = 0; slot < totalSlots; slot++) {
    const slotMinutes = minMinutes + (slot * slotDuration);
    const isHourMark = slotMinutes % 60 === 0;

    const timeCell = document.createElement('div');
    timeCell.className = `grid-time-cell ${isHourMark ? 'hour-mark' : ''}`;
    timeCell.style.gridColumn = '1';
    timeCell.style.gridRow = `${slot + 2}`;
    timeCell.textContent = isHourMark ? formatMinutesToAmPm(slotMinutes) : `:30`;
    gridContainer.appendChild(timeCell);
  }

  displayDays.forEach((dayNum, dayIndex) => {
    const dayColIndex = dayIndex + 2;

    for (let slot = 0; slot < totalSlots; slot++) {
      const slotMinutes = minMinutes + (slot * slotDuration);
      const isHourMark = slotMinutes % 60 === 0;

      const cell = document.createElement('div');
      cell.className = `grid-cell ${isHourMark ? 'hour-mark' : ''}`;
      cell.style.gridColumn = `${dayColIndex}`;
      cell.style.gridRow = `${slot + 2}`;
      gridContainer.appendChild(cell);
    }

    const dayCanvas = document.createElement('div');
    dayCanvas.className = 'grid-day-column';
    dayCanvas.style.gridColumn = `${dayColIndex}`;
    dayCanvas.style.gridRow = `2 / span ${totalSlots}`;
    dayCanvas.style.height = `${totalSlots * rowHeightPx}px`;

    const dayBlocks = blocks.filter(b => b.day === dayNum);
    const resolvedBlocks = resolveTileOverlaps(dayBlocks);

    resolvedBlocks.forEach(block => {
      const topOffset = ((block.startMinutes - minMinutes) / slotDuration) * rowHeightPx;
      const tileHeight = Math.max(((block.endMinutes - block.startMinutes) / slotDuration) * rowHeightPx - 2, 28);
      const colWidthPercent = 100 / block.totalCols;
      const leftPercent = block.layoutCol * colWidthPercent;

      const tileEl = createScheduleTileElement(block, leftPercent, colWidthPercent, topOffset, tileHeight);
      dayCanvas.appendChild(tileEl);
    });

    gridContainer.appendChild(dayCanvas);
  });

  if (blocks.length === 0) {
    const emptyState = document.createElement('div');
    emptyState.className = 'grid-empty-state';
    emptyState.style.gridColumn = `1 / span ${dayColCount + 1}`;
    emptyState.style.gridRow = '2 / span 10';
    emptyState.innerHTML = `
      <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
      <h3>No Schedule Sections Match Active Criteria</h3>
      <p>Adjust your location, building, GE pattern, or conflict filters to see classes.</p>
    `;
    gridContainer.appendChild(emptyState);
  }
}

/**
 * REQUIREMENT 2: STACKED ROOM MATRIX VIEW
 * Vertically stacks dedicated weekly calendar grids for each Building & Room currently in view.
 */
function renderStackedRoomMatrix() {
  const container = document.getElementById('room-matrix-stacked');
  if (!container) return;
  container.innerHTML = '';

  const blocks = appState.filteredBlocks;
  if (blocks.length === 0) {
    container.innerHTML = `
      <div class="grid-empty-state">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
        <h3>No Rooms with Sections in Current Filter</h3>
        <p>Adjust your filters to see room schedules.</p>
      </div>
    `;
    return;
  }

  // Group blocks by unique Building & Room
  const roomGroups = new Map();
  blocks.forEach(b => {
    const key = `${b.building} ${b.room}`.trim() || 'Unassigned Room';
    if (!roomGroups.has(key)) {
      roomGroups.set(key, []);
    }
    roomGroups.get(key).push(b);
  });

  const sortedRooms = Array.from(roomGroups.keys()).sort();

  let displayDays = [1, 2, 3, 4, 5];
  if (appState.daySpan === 6) displayDays = [1, 2, 3, 4, 5, 6];
  if (appState.daySpan === 7) displayDays = [1, 2, 3, 4, 5, 6, 7];
  if (appState.daySpan === 'weekend') displayDays = [5, 6, 7];
  if (appState.daySpan === 'active') {
    const activeDayNums = new Set(blocks.map(b => b.day));
    displayDays = [1, 2, 3, 4, 5, 6, 7].filter(d => activeDayNums.has(d));
    if (displayDays.length === 0) displayDays = [1, 2, 3, 4, 5];
  }

  const slotDuration = 30;
  const rowHeightPx = 32;
  let minMinutes = 7 * 60;
  let maxMinutes = 22 * 60;
  if (blocks.length > 0) {
    const earliestClass = Math.min(...blocks.map(b => b.startMinutes));
    const latestClass = Math.max(...blocks.map(b => b.endMinutes));
    minMinutes = Math.max(6 * 60, Math.floor((earliestClass - 30) / 60) * 60);
    maxMinutes = Math.min(23 * 60, Math.ceil((latestClass + 30) / 60) * 60);
  }
  const totalSlots = Math.ceil((maxMinutes - minMinutes) / slotDuration);

  // Render a clean stacked card for each room
  sortedRooms.forEach(roomKey => {
    const roomBlocks = roomGroups.get(roomKey);
    const firstBlock = roomBlocks[0];
    const bldgDesc = firstBlock.buildingDesc || firstBlock.building;
    const campusLoc = firstBlock.parentSection.location || 'Campus';
    const coursesInRoom = Array.from(new Set(roomBlocks.map(b => b.parentSection.course))).sort();

    const card = document.createElement('div');
    card.className = 'room-matrix-card';

    // Room Card Header
    const cardHeader = document.createElement('div');
    cardHeader.className = 'room-card-header';
    cardHeader.innerHTML = `
      <div class="room-card-title-group">
        <span class="room-code-badge">${roomKey}</span>
        <div>
          <div class="room-bldg-desc">${bldgDesc}</div>
          <div style="font-size:11px; color:var(--text-muted);">${campusLoc.includes('Campus') ? campusLoc : `${campusLoc} Campus`}</div>
        </div>
      </div>
      <div class="room-card-meta">
        <span class="room-meta-pill"><strong>${roomBlocks.length}</strong> Meetings Scheduled</span>
        <span class="room-meta-pill">Courses: <strong>${coursesInRoom.join(', ')}</strong></span>
      </div>
    `;
    card.appendChild(cardHeader);

    // Dedicated Calendar Grid for this Room
    const gridWrapper = document.createElement('div');
    gridWrapper.className = 'room-calendar-grid-wrapper';

    const roomGrid = document.createElement('div');
    roomGrid.className = 'schedule-grid';
    roomGrid.style.gridTemplateColumns = `80px repeat(${displayDays.length}, minmax(160px, 1fr))`;
    roomGrid.style.gridTemplateRows = `auto repeat(${totalSlots}, ${rowHeightPx}px)`;

    const cornerCell = document.createElement('div');
    cornerCell.className = 'grid-corner-cell';
    cornerCell.textContent = 'TIME';
    roomGrid.appendChild(cornerCell);

    displayDays.forEach((dayNum, dayIndex) => {
      const dayMeta = DAY_MAP[dayNum];
      const dayHeader = document.createElement('div');
      dayHeader.className = 'grid-day-header';
      dayHeader.style.gridColumn = `${dayIndex + 2}`;
      dayHeader.style.gridRow = '1';

      const meetingsOnDay = roomBlocks.filter(b => b.day === dayNum);
      dayHeader.innerHTML = `
        <span class="day-name">${dayMeta.name}</span>
        <span class="day-count">${meetingsOnDay.length} in this room</span>
      `;
      roomGrid.appendChild(dayHeader);
    });

    for (let slot = 0; slot < totalSlots; slot++) {
      const slotMinutes = minMinutes + (slot * slotDuration);
      const isHourMark = slotMinutes % 60 === 0;

      const timeCell = document.createElement('div');
      timeCell.className = `grid-time-cell ${isHourMark ? 'hour-mark' : ''}`;
      timeCell.style.gridColumn = '1';
      timeCell.style.gridRow = `${slot + 2}`;
      timeCell.textContent = isHourMark ? formatMinutesToAmPm(slotMinutes) : ':30';
      roomGrid.appendChild(timeCell);
    }

    displayDays.forEach((dayNum, dayIndex) => {
      const dayColIndex = dayIndex + 2;

      for (let slot = 0; slot < totalSlots; slot++) {
        const slotMinutes = minMinutes + (slot * slotDuration);
        const isHourMark = slotMinutes % 60 === 0;

        const cell = document.createElement('div');
        cell.className = `grid-cell ${isHourMark ? 'hour-mark' : ''}`;
        cell.style.gridColumn = `${dayColIndex}`;
        cell.style.gridRow = `${slot + 2}`;
        roomGrid.appendChild(cell);
      }

      const dayCanvas = document.createElement('div');
      dayCanvas.className = 'grid-day-column';
      dayCanvas.style.gridColumn = `${dayColIndex}`;
      dayCanvas.style.gridRow = `2 / span ${totalSlots}`;
      dayCanvas.style.height = `${totalSlots * rowHeightPx}px`;

      const dayRoomBlocks = roomBlocks.filter(b => b.day === dayNum);
      const resolvedBlocks = resolveTileOverlaps(dayRoomBlocks);

      resolvedBlocks.forEach(block => {
        const topOffset = ((block.startMinutes - minMinutes) / slotDuration) * rowHeightPx;
        const tileHeight = Math.max(((block.endMinutes - block.startMinutes) / slotDuration) * rowHeightPx - 2, 28);
        const colWidthPercent = 100 / block.totalCols;
        const leftPercent = block.layoutCol * colWidthPercent;

        const tileEl = createScheduleTileElement(block, leftPercent, colWidthPercent, topOffset, tileHeight);
        dayCanvas.appendChild(tileEl);
      });

      roomGrid.appendChild(dayCanvas);
    });

    gridWrapper.appendChild(roomGrid);
    card.appendChild(gridWrapper);
    container.appendChild(card);
  });
}

/**
 * Helper to determine GE Area Badge color class with unique high-contrast colors per area
 */
function getGeAreaBadgeClass(areaStr) {
  if (!areaStr) return 'ge-color-other';
  const s = String(areaStr).trim().toUpperCase();

  if (s === '1A') return 'ge-color-1a';
  if (s === '1B') return 'ge-color-1b';
  if (s === '1C') return 'ge-color-1c';
  if (s === '1B OR 1C' || s.includes('1B OR 1C') || (s.includes('1B') && s.includes('1C'))) return 'ge-color-1b';
  if (s === '2') return 'ge-color-2';
  if (s === '3A') return 'ge-color-3a';
  if (s === '3B') return 'ge-color-3b';
  if (s === '3A OR 3B' || s.includes('3A OR 3B') || (s.includes('3A') && s.includes('3B'))) return 'ge-color-3a';
  if (s === '4') return 'ge-color-4';
  if (s === '5A') return 'ge-color-5a';
  if (s === '5B') return 'ge-color-5b';
  if (s === '5') return 'ge-color-5a';
  if (s === '6') return 'ge-color-6';
  if (s === 'US1' || s.includes('US1')) return 'ge-color-us1';

  // Fallbacks
  if (s.startsWith('1A')) return 'ge-color-1a';
  if (s.startsWith('1B')) return 'ge-color-1b';
  if (s.startsWith('1C')) return 'ge-color-1c';
  if (s.startsWith('2')) return 'ge-color-2';
  if (s.startsWith('3A')) return 'ge-color-3a';
  if (s.startsWith('3B')) return 'ge-color-3b';
  if (s.startsWith('4')) return 'ge-color-4';
  if (s.startsWith('5A')) return 'ge-color-5a';
  if (s.startsWith('5B')) return 'ge-color-5b';
  if (s.startsWith('5')) return 'ge-color-5a';
  if (s.startsWith('6')) return 'ge-color-6';

  return 'ge-color-other';
}

/**
 * REQUIREMENT 3: GENERAL EDUCATION (GE) COVERAGE AUDIT VIEW
 * Audits GE Pattern & Area offerings for the selected campus,
 * displaying coverage status, missing areas, interactive course chips,
 * and swappable calendar grid (Day View vs Room View).
 * Valid Patterns: CalGETC (11 areas) and Chaffey Local (7 areas).
 */
function renderGeCoverageDashboard() {
  const container = document.getElementById('ge-coverage-container');
  if (!container) return;
  container.innerHTML = '';

  let activePattern = appState.geActivePattern || 'CALGETC';
  if (activePattern !== 'CHAFFEY') activePattern = 'CALGETC';
  appState.geActivePattern = activePattern;

  const standardAreas = GE_STANDARD_AREAS[activePattern] || GE_STANDARD_AREAS['CALGETC'];
  const patternDisplayLabel = activePattern === 'CHAFFEY' ? 'Chaffey Local' : 'CalGETC';

  // Construct comprehensive scope label from active filters
  const scopeParts = [];
  if (appState.filters.location) scopeParts.push(appState.filters.location);
  else scopeParts.push('All Campuses');

  if (appState.filters.building) scopeParts.push(`Bldg ${appState.filters.building}`);
  if (appState.filters.room) scopeParts.push(`Room ${appState.filters.room}`);
  if (appState.filters.subject.size > 0) scopeParts.push(`Subject: ${Array.from(appState.filters.subject).join(', ')}`);
  if (appState.filters.course.size > 0) scopeParts.push(`Course: ${Array.from(appState.filters.course).join(', ')}`);
  if (appState.filters.dayEvening.size > 0) scopeParts.push(`Type: ${Array.from(appState.filters.dayEvening).map(formatDayEveningLabel).join(', ')}`);
  if (appState.filters.acc.size > 0) scopeParts.push(`ACC: ${Array.from(appState.filters.acc).join(', ')}`);
  if (appState.filters.modality.size > 0) scopeParts.push(`Modality: ${Array.from(appState.filters.modality).join(', ')}`);
  if (appState.filters.session.size > 0) scopeParts.push(`Session: ${Array.from(appState.filters.session).join(', ')}`);
  if (appState.filters.search) scopeParts.push(`"${appState.filters.search}"`);

  const scopeLabel = scopeParts.join(' • ');

  // Base sections respect all active user filters!
  const baseSections = appState.filteredSections || appState.sections;
  const currentSections = baseSections.filter(sec => {
    return sec.geAssociations && sec.geAssociations.some(a => a.pattern === activePattern);
  });

  const areaCoverageMap = new Map();
  standardAreas.forEach(a => {
    areaCoverageMap.set(a.code, {
      meta: a,
      offerings: []
    });
  });

  currentSections.forEach(sec => {
    const secAreas = (sec.gePatternAreaMap && sec.gePatternAreaMap[activePattern]) || new Set();
    standardAreas.forEach(stdArea => {
      const satisfies = stdArea.matchAreas.some(targetArea => secAreas.has(targetArea));
      if (satisfies) {
        const item = areaCoverageMap.get(stdArea.code);
        if (item && !item.offerings.some(o => o.sectionUniqueId === sec.sectionUniqueId)) {
          item.offerings.push(sec);
        }
      }
    });
  });

  let coveredCount = 0;
  let missingCount = 0;
  areaCoverageMap.forEach(v => {
    if (v.offerings.length > 0) coveredCount++;
    else missingCount++;
  });

  const coveragePercent = standardAreas.length > 0 ? Math.round((coveredCount / standardAreas.length) * 100) : 0;

  // 1. Header Row (Title, Campus, Pattern Tabs, and Sub-Layout Switcher)
  const headerRow = document.createElement('div');
  headerRow.className = 'ge-coverage-header';
  headerRow.innerHTML = `
    <div class="ge-coverage-title">
      <h3>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#990000" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/></svg>
        General Education Coverage Audit: <span>${scopeLabel}</span>
      </h3>
      <p>Identify active offerings and unscheduled GE requirements for <strong>${patternDisplayLabel}</strong>.</p>
    </div>
    <div class="ge-header-controls">
      <div style="display:flex; align-items:center; gap:6px;">
        <span style="font-size:11px; font-weight:700; color:var(--text-muted);">PATTERN:</span>
        <div class="ge-pattern-tabs" id="ge-pattern-tabs">
          <button type="button" class="ge-pattern-tab ${activePattern === 'CALGETC' ? 'active' : ''}" data-pattern="CALGETC">CalGETC</button>
          <button type="button" class="ge-pattern-tab ${activePattern === 'CHAFFEY' ? 'active' : ''}" data-pattern="CHAFFEY">Chaffey Local</button>
        </div>
      </div>
      <div class="sub-layout-switcher" id="ge-sub-layout-switcher">
        <span class="sub-layout-label">Grid View:</span>
        <div class="btn-sub-toggle-group">
          <button type="button" class="btn-sub-layout-toggle ${appState.geSubLayout === 'day' ? 'active' : ''}" data-sublayout="day" title="View as Standard Day Grid">Day View</button>
          <button type="button" class="btn-sub-layout-toggle ${appState.geSubLayout === 'room' ? 'active' : ''}" data-sublayout="room" title="View as Stacked Room Grids">Room View</button>
        </div>
      </div>
    </div>
  `;
  container.appendChild(headerRow);

  // Pattern Tab switch listener
  headerRow.querySelectorAll('.ge-pattern-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      appState.geActivePattern = tab.getAttribute('data-pattern');
      renderGeCoverageDashboard();
      applyFiltersAndRender();
    });
  });

  // Sub-layout switcher listener (Day View vs Room View)
  headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      appState.geSubLayout = btn.getAttribute('data-sublayout');
      headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      dispatchViewRender();
    });
  });

  // 2. Summary Status Bar
  const summaryBar = document.createElement('div');
  summaryBar.className = 'ge-audit-summary-bar';
  summaryBar.innerHTML = `
    <div class="ge-audit-stat-group">
      <div class="ge-stat-item">
        <span class="ge-stat-label">Coverage Rate:</span>
        <span class="ge-stat-val ${coveragePercent >= 80 ? 'val-covered' : ''}">${coveragePercent}%</span>
      </div>
      <div class="ge-stat-item">
        <span class="ge-stat-label">Scheduled Areas:</span>
        <span class="ge-stat-val val-covered">${coveredCount} of ${standardAreas.length} Areas</span>
      </div>
      <div class="ge-stat-item">
        <span class="ge-stat-label">Missing Requirements:</span>
        <span class="ge-stat-val val-missing">${missingCount} Unscheduled Areas</span>
      </div>
    </div>
    <div>
      <label class="checkbox-label" style="font-size:11px; font-weight:600;">
        <input type="checkbox" id="chk-ge-missing-only" ${appState.geShowMissingOnly ? 'checked' : ''} />
        <span>Highlight Missing Areas Only</span>
      </label>
    </div>
  `;
  container.appendChild(summaryBar);

  summaryBar.querySelector('#chk-ge-missing-only')?.addEventListener('change', (e) => {
    appState.geShowMissingOnly = e.target.checked;
    renderGeCoverageDashboard();
  });

  // 3. Grid of GE Area Cards
  const areasGrid = document.createElement('div');
  areasGrid.className = 'ge-areas-grid';

  let renderedCardsCount = 0;
  areaCoverageMap.forEach((data, areaCode) => {
    const isCovered = data.offerings.length > 0;
    if (appState.geShowMissingOnly && isCovered) return;

    renderedCardsCount++;
    const card = document.createElement('div');
    card.className = `ge-area-card ${isCovered ? 'is-covered' : 'is-missing'}`;

    const colorClass = getGeAreaBadgeClass(areaCode);

    const offeringsChips = data.offerings.map(sec => {
      const times = sec.scheduleBlocks.map(b => `${DAY_MAP[b.days[0]]?.short || ''} ${b.startTimeStr}`).join(', ');
      return `<span class="ge-course-chip" data-sec-id="${sec.sectionUniqueId}" title="${sec.sectionName} (${times})">${sec.course} (${sec.facultyName || 'Staff'})</span>`;
    }).join('');

    card.innerHTML = `
      <div class="ge-area-card-header">
        <span class="ge-area-badge ${colorClass}">Area ${areaCode}</span>
        <span class="ge-status-pill ${isCovered ? 'status-covered' : 'status-missing'}">
          ${isCovered ? `✓ ${data.offerings.length} Scheduled` : `⚠ Missing in ${scopeLabel}`}
        </span>
      </div>
      <div class="ge-area-name">${data.meta.name}</div>
      <div style="font-size:10.5px; color:var(--text-muted);">${data.meta.desc}</div>
      ${isCovered ? `
        <div class="ge-area-courses-list">
          ${offeringsChips}
        </div>
      ` : `
        <div style="font-size:11px; color:#b91c1c; font-weight:600; margin-top:2px;">
          No class section scheduled for Area ${areaCode}
        </div>
      `}
    `;

    areasGrid.appendChild(card);
  });

  if (renderedCardsCount === 0 && appState.geShowMissingOnly) {
    areasGrid.innerHTML = `
      <div style="grid-column: 1 / -1; padding: 28px; text-align: center; background: var(--bg-card); border: 1px dashed var(--border-color); border-radius: 8px; color: var(--text-secondary); font-size: 13px;">
        <span style="font-size: 18px; margin-right: 6px;">🎉</span>
        <strong>All ${standardAreas.length} general education areas are scheduled!</strong> No unscheduled requirements exist for this scope.
      </div>
    `;
  }

  container.appendChild(areasGrid);

  // Attach hover and click listeners to all section chips in GE Coverage
  container.querySelectorAll('.ge-course-chip').forEach(chip => {
    const secId = chip.getAttribute('data-sec-id');
    const sec = appState.sections.find(s => s.sectionUniqueId === secId);
    if (sec) {
      chip.addEventListener('mouseenter', (e) => showPopover(e, sec));
      chip.addEventListener('mousemove', (e) => positionPopover(e));
      chip.addEventListener('mouseleave', hidePopover);
      chip.addEventListener('click', () => showSectionModal(sec));
    }
  });
}

/**
 * REQUIREMENT 6: EVENING & WEEKEND COVERAGE AUDIT VIEW
 * Audits Evening & Weekend course offerings:
 * - Classes scheduled on Friday, Saturday, or Sunday (days 5, 6, 7)
 * - Classes scheduled on or after 4:30 PM (startMinutes >= 990)
 * - Schedule Type is "E" evening
 */
function renderEveningCoverageDashboard() {
  const container = document.getElementById('evening-coverage-container');
  if (!container) return;
  container.innerHTML = '';

  // Construct comprehensive scope label from active filters
  const scopeParts = [];
  if (appState.filters.location) scopeParts.push(appState.filters.location);
  else scopeParts.push('All Campuses');

  if (appState.filters.building) scopeParts.push(`Bldg ${appState.filters.building}`);
  if (appState.filters.room) scopeParts.push(`Room ${appState.filters.room}`);
  if (appState.filters.subject.size > 0) scopeParts.push(`Subject: ${Array.from(appState.filters.subject).join(', ')}`);
  if (appState.filters.course.size > 0) scopeParts.push(`Course: ${Array.from(appState.filters.course).join(', ')}`);
  if (appState.filters.dayEvening.size > 0) scopeParts.push(`Type: ${Array.from(appState.filters.dayEvening).map(formatDayEveningLabel).join(', ')}`);
  if (appState.filters.acc.size > 0) scopeParts.push(`ACC: ${Array.from(appState.filters.acc).join(', ')}`);
  if (appState.filters.modality.size > 0) scopeParts.push(`Modality: ${Array.from(appState.filters.modality).join(', ')}`);
  if (appState.filters.session.size > 0) scopeParts.push(`Session: ${Array.from(appState.filters.session).join(', ')}`);
  if (appState.filters.search) scopeParts.push(`"${appState.filters.search}"`);

  const scopeLabel = scopeParts.join(' • ');

  // Base sections respect ALL active filters (Building, Room, Subject, ACC, Search, etc.)
  const baseSections = appState.filteredSections || appState.sections;
  const eveningSections = baseSections.filter(sec => {
    // If building or room filter is active, only include sections that actually have an evening meeting in that building/room
    const matchingBlocks = sec.scheduleBlocks.filter(b => {
      if (appState.filters.building && b.building.toLowerCase() !== appState.filters.building.toLowerCase()) return false;
      if (appState.filters.room && b.room.toLowerCase() !== appState.filters.room.toLowerCase()) return false;
      return true;
    });

    if (appState.filters.building || appState.filters.room) {
      return matchingBlocks.some(isEveningBlock);
    }
    return isEveningSection(sec);
  });

  const eveningBlocks = appState.filteredBlocks.filter(isEveningBlock);

  // Categorize sections by time band / schedule pattern
  const weekdayEarlyEve = [];  // Mon-Thu, start between 4:30 PM (990) and 6:30 PM (1110)
  const weekdayLateNight = []; // Mon-Thu, start after 6:30 PM (1110)
  const fridayOfferings = [];  // Friday classes
  const weekendOfferings = []; // Saturday & Sunday classes

  eveningSections.forEach(sec => {
    // Only inspect blocks matching the active building & room filters
    const relevantBlocks = sec.scheduleBlocks.filter(b => {
      if (appState.filters.building && b.building.toLowerCase() !== appState.filters.building.toLowerCase()) return false;
      if (appState.filters.room && b.room.toLowerCase() !== appState.filters.room.toLowerCase()) return false;
      return true;
    });

    relevantBlocks.forEach(b => {
      const days = b.days || [];
      if (days.includes(6) || days.includes(7)) {
        if (!weekendOfferings.includes(sec)) weekendOfferings.push(sec);
      }
      if (days.includes(5)) {
        if (!fridayOfferings.includes(sec)) fridayOfferings.push(sec);
      }
      const hasMonThu = days.some(d => d >= 1 && d <= 4);
      if (hasMonThu) {
        if (b.startMinutes !== null && b.startMinutes >= 990 && b.startMinutes < 1110) {
          if (!weekdayEarlyEve.includes(sec)) weekdayEarlyEve.push(sec);
        } else if (b.startMinutes !== null && b.startMinutes >= 1110) {
          if (!weekdayLateNight.includes(sec)) weekdayLateNight.push(sec);
        } else if (String(sec.dayOrEvening).toUpperCase().startsWith('E')) {
          if (!weekdayEarlyEve.includes(sec)) weekdayEarlyEve.push(sec);
        }
      }
    });
  });

  const totalEveningCount = eveningSections.length;
  const activeRoomsSet = new Set(eveningBlocks.map(b => `${b.building} ${b.room}`));

  // 1. Header Row
  const headerRow = document.createElement('div');
  headerRow.className = 'evening-coverage-header';
  headerRow.innerHTML = `
    <div class="evening-coverage-title">
      <h3>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#990000" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>
        Evening & Weekend Coverage Audit: <span>${scopeLabel}</span>
      </h3>
      <p>Auditing classes scheduled Friday–Sunday, starting at/after 4:30 PM, or designated Evening.</p>
    </div>
    <div class="evening-header-controls">
      <div class="sub-layout-switcher" id="evening-sub-layout-switcher">
        <span class="sub-layout-label">Grid View:</span>
        <div class="btn-sub-toggle-group">
          <button type="button" class="btn-sub-layout-toggle ${appState.eveningSubLayout === 'day' ? 'active' : ''}" data-sublayout="day" title="View as Standard Day Grid">Day View</button>
          <button type="button" class="btn-sub-layout-toggle ${appState.eveningSubLayout === 'room' ? 'active' : ''}" data-sublayout="room" title="View as Stacked Room Grids">Room View</button>
        </div>
      </div>
    </div>
  `;
  container.appendChild(headerRow);

  // Sub-layout switcher listener
  headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      appState.eveningSubLayout = btn.getAttribute('data-sublayout');
      headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      dispatchViewRender();
    });
  });

  // 2. Summary KPI Cards
  const summaryBar = document.createElement('div');
  summaryBar.className = 'evening-audit-summary-bar';
  summaryBar.innerHTML = `
    <div class="evening-stat-card highlight-card">
      <span class="evening-stat-card-title">Total Evening/Weekend</span>
      <span class="evening-stat-card-value">${totalEveningCount}</span>
      <span class="evening-stat-card-sub">Active course sections</span>
    </div>
    <div class="evening-stat-card">
      <span class="evening-stat-card-title">Weekday Evenings</span>
      <span class="evening-stat-card-value">${weekdayEarlyEve.length + weekdayLateNight.length}</span>
      <span class="evening-stat-card-sub">Mon–Thu classes (4:30 PM+)</span>
    </div>
    <div class="evening-stat-card friday-card">
      <span class="evening-stat-card-title">Friday Offerings</span>
      <span class="evening-stat-card-value">${fridayOfferings.length}</span>
      <span class="evening-stat-card-sub">Friday schedule sessions</span>
    </div>
    <div class="evening-stat-card weekend-card">
      <span class="evening-stat-card-title">Weekend Offerings</span>
      <span class="evening-stat-card-value">${weekendOfferings.length}</span>
      <span class="evening-stat-card-sub">Saturday & Sunday sessions</span>
    </div>
    <div class="evening-stat-card">
      <span class="evening-stat-card-title">Active Rooms</span>
      <span class="evening-stat-card-value">${activeRoomsSet.size}</span>
      <span class="evening-stat-card-sub">Classrooms utilized</span>
    </div>
  `;
  container.appendChild(summaryBar);

  // 3. Time Bands Breakdown Grid
  const bandsGrid = document.createElement('div');
  bandsGrid.className = 'evening-bands-grid';

  const bandsConfig = [
    {
      title: 'Weekday Twilight & Early Evening',
      timeDesc: 'Mon–Thu starting 4:30 PM – 6:30 PM',
      sections: weekdayEarlyEve,
      tag: '4:30 PM – 6:30 PM',
      tagClass: 'evening-tag-weekday'
    },
    {
      title: 'Weekday Late Evening / Night',
      timeDesc: 'Mon–Thu starting 6:30 PM and later',
      sections: weekdayLateNight,
      tag: '6:30 PM – 10:00 PM',
      tagClass: 'evening-tag-night'
    },
    {
      title: 'Friday Afternoon & Evening',
      timeDesc: 'Friday academic meetings',
      sections: fridayOfferings,
      tag: 'Friday Classes',
      tagClass: 'evening-tag-friday'
    },
    {
      title: 'Weekend College (Sat & Sun)',
      timeDesc: 'Saturday & Sunday intensive sessions',
      sections: weekendOfferings,
      tag: 'Sat / Sun',
      tagClass: 'evening-tag-weekend'
    }
  ];

  bandsConfig.forEach(band => {
    const card = document.createElement('div');
    card.className = 'evening-band-card';

    const chipsHtml = band.sections.length > 0
      ? band.sections.map(sec => {
          const times = sec.scheduleBlocks
            .filter(b => {
              if (appState.filters.building && b.building.toLowerCase() !== appState.filters.building.toLowerCase()) return false;
              if (appState.filters.room && b.room.toLowerCase() !== appState.filters.room.toLowerCase()) return false;
              return true;
            })
            .map(b => `${DAY_MAP[b.days[0]]?.short || ''} ${b.startTimeStr}`).join(', ');
          return `<span class="ge-course-chip" data-sec-id="${sec.sectionUniqueId}" title="${sec.sectionName} (${times})">${sec.course} (${sec.facultyName || 'Staff'})</span>`;
        }).join('')
      : `<span style="font-size:12px; color:var(--text-muted); font-style:italic;">No classes currently scheduled in this band matching active filters (${scopeLabel}).</span>`;

    card.innerHTML = `
      <div class="evening-band-header">
        <div>
          <strong style="font-size:13px; color:var(--text-primary);">${band.title}</strong>
          <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">${band.timeDesc}</div>
        </div>
        <span class="evening-band-tag ${band.tagClass}">${band.tag} (${band.sections.length})</span>
      </div>
      <div class="evening-courses-list">
        ${chipsHtml}
      </div>
    `;

    bandsGrid.appendChild(card);
  });

  container.appendChild(bandsGrid);

  // Attach hover and click listeners to all section chips in Evening Coverage
  container.querySelectorAll('.ge-course-chip').forEach(chip => {
    const secId = chip.getAttribute('data-sec-id');
    const sec = appState.sections.find(s => s.sectionUniqueId === secId);
    if (sec) {
      chip.addEventListener('mouseenter', (e) => showPopover(e, sec));
      chip.addEventListener('mousemove', (e) => positionPopover(e));
      chip.addEventListener('mouseleave', hidePopover);
      chip.addEventListener('click', () => showSectionModal(sec));
    }
  });
}

/**
 * ==============================================================================================
 * REQUIREMENT 10: FAST-TRACK SEQUENTIAL OFFERING & PAIRING AUDIT VIEW
 * ==============================================================================================
 * Analyzes Fast-Track 1 (first 8 weeks) and Fast-Track 2 (second 8 weeks) offerings:
 * 1. Matching Pairs: FT1 and FT2 scheduled in the same room on the same day/timeslot.
 * 2. Optimized Pairings:
 *    - Cross-GE Spanning: FT1 satisfies one GE Area and FT2 satisfies a different GE Area.
 *    - Sequential Department Cohorts: Consecutive courses in same program (e.g. CISGAME-1 -> CISGAME-2).
 *    - Mixed GE + Major: Combines a GE requirement with a degree major prerequisite.
 * 3. Unoptimized Pairings:
 *    - Redundant GE Area: Both FT1 and FT2 satisfy the identical GE Area (already fulfilled in FT1).
 *    - Same Course Repeated: Both FT1 and FT2 offer identical course (e.g. KINLEC-16 -> KINLEC-16).
 * 4. Unpaired Fast-Track Offerings:
 *    - FT1 offered without a complementary FT2 in that room/timeslot (student has empty second 8 weeks).
 *    - FT2 offered without a complementary FT1 in that room/timeslot (student has empty first 8 weeks).
 */
function analyzeFastTrackOfferings(baseSections, activeBlocks) {
  function getSectionGeAreaSet(sec) {
    const areas = new Set();
    if (sec && sec.geAssociations) {
      sec.geAssociations.forEach(pair => {
        if (pair.area) areas.add(pair.area);
      });
    }
    return areas;
  }

  // REQUIREMENT: For Fast-Track Coverage, ignore Scheduled Meeting Components or Parts that do not list days,
  // times, locations, or hours (etc). Also ignore components/parts listed as Online Lec; Canvas CANVAS (Online).
  function getRelevantBlocks(sec) {
    return (sec.scheduleBlocks || []).filter(b => {
      if (!isPhysicalMeetingBlock(b)) return false;
      if (appState.filters.building && b.building.toLowerCase() !== appState.filters.building.toLowerCase()) return false;
      if (appState.filters.room && b.room.toLowerCase() !== appState.filters.room.toLowerCase()) return false;
      return true;
    });
  }

  // Only consider Fast-Track sections that have valid scheduled physical classroom meetings
  const ftSections = baseSections.filter(s => isFastTrackSection(s) && getRelevantBlocks(s).length > 0);
  const ft1Sections = ftSections.filter(s => getFastTrackSessionType(s) === 'FT1');
  const ft2Sections = ftSections.filter(s => getFastTrackSessionType(s) === 'FT2');

  const pairs = [];
  const matchedMeetingKeySet = new Set();

  ft1Sections.forEach(s1 => {
    const blocks1 = getRelevantBlocks(s1);
    blocks1.forEach(b1 => {
      if (!b1.building || !b1.room || b1.startMinutes === null || b1.endMinutes === null) return;
      const b1Days = b1.days || [];
      if (b1Days.length === 0) return;

      ft2Sections.forEach(s2 => {
        const blocks2 = getRelevantBlocks(s2);
        blocks2.forEach(b2 => {
          if (!b2.building || !b2.room || b2.startMinutes === null || b2.endMinutes === null) return;
          if (b1.building.toLowerCase() !== b2.building.toLowerCase()) return;
          if (b1.room.toLowerCase() !== b2.room.toLowerCase()) return;

          const b2Days = b2.days || [];
          const commonDays = b1Days.filter(d => b2Days.includes(d));
          if (commonDays.length === 0) return;

          // Check for overlapping schedule timeslot
          if (b1.startMinutes < b2.endMinutes && b2.startMinutes < b1.endMinutes) {
            const pairKey = `${s1.sectionUniqueId}_${b1.rowId}__${s2.sectionUniqueId}_${b2.rowId}`;
            if (matchedMeetingKeySet.has(pairKey)) return;
            matchedMeetingKeySet.add(pairKey);

            const ge1 = getSectionGeAreaSet(s1);
            const ge2 = getSectionGeAreaSet(s2);
            const isSameCourse = s1.course.toUpperCase() === s2.course.toUpperCase();

            let status = 'OPTIMIZED';
            let reasonType = '';
            let reasonText = '';

            if (ge1.size > 0 && ge2.size > 0) {
              const commonGe = [...ge1].filter(a => ge2.has(a));
              if (commonGe.length > 0 && commonGe.length === ge1.size && commonGe.length === ge2.size) {
                status = 'UNOPTIMIZED';
                reasonType = 'ge_redundant';
                reasonText = `Redundant GE Area: Both courses satisfy Area ${sortGeAreas(commonGe).join(', ')}. Students who satisfied Area ${sortGeAreas(commonGe).join(', ')} in Fast-Track 1 have no incentive to register in Fast-Track 2.`;
              } else {
                status = 'OPTIMIZED';
                reasonType = 'ge_cross';
                reasonText = `Cross-GE Spanning: Satisfies Area ${sortGeAreas(ge1).join(', ')} (FT1) → Area ${sortGeAreas(ge2).join(', ')} (FT2). Enables students to fulfill two distinct GE requirements in one continuous semester timeslot.`;
              }
            } else if (ge1.size === 0 && ge2.size === 0) {
              if (isSameCourse) {
                status = 'UNOPTIMIZED';
                reasonType = 'same_course';
                reasonText = `Same Course Repeated: Both Fast-Track sessions offer identical course ${s1.course}. Students cannot repeat the same course for duplicate degree credit.`;
              } else {
                status = 'OPTIMIZED';
                reasonType = 'sequential_cohort';
                reasonText = `Sequential Program Cohort: Progressive course pair (${s1.course} → ${s2.course}) enables program majors to accelerate through consecutive degree prerequisites in a single term.`;
              }
            } else {
              status = 'OPTIMIZED';
              reasonType = 'mixed_ge_major';
              const geSide = ge1.size > 0 ? `Area ${sortGeAreas(ge1).join(', ')} (${s1.course})` : `Area ${sortGeAreas(ge2).join(', ')} (${s2.course})`;
              const majorSide = ge1.size > 0 ? `Major Course (${s2.course})` : `Major Course (${s1.course})`;
              reasonText = `Balanced GE + Major Pairing: Strategically pairs ${geSide} with ${majorSide} for an accelerated degree-applicable schedule.`;
            }

            pairs.push({
              key: pairKey,
              sec1: s1,
              block1: b1,
              sec2: s2,
              block2: b2,
              building: b1.building,
              buildingDesc: b1.buildingDesc,
              room: b1.room,
              days: commonDays,
              daysStr: commonDays.map(d => DAY_MAP[d]?.name || d).join(', '),
              timeSpanStr: `${b1.startTimeStr} – ${b1.endTimeStr}`,
              status: status,
              reasonType: reasonType,
              reasonText: reasonText
            });
          }
        });
      });
    });
  });

  // Identify Unpaired Offerings
  const unpairedList = [];

  ft1Sections.forEach(s => {
    const blocks = getRelevantBlocks(s);
    blocks.forEach(b => {
      const isMeetingPaired = pairs.some(p => p.sec1.sectionUniqueId === s.sectionUniqueId && p.block1.rowId === b.rowId);
      if (!isMeetingPaired) {
        unpairedList.push({
          type: 'FT1',
          sessionLabel: 'Fast-Track 1 (First 8-Weeks)',
          missingLabel: 'Missing Fast-Track 2 Companion',
          missingDesc: 'Scheduled in the first 8 weeks without a complementary Fast-Track 2 course in this room & timeslot. The classroom sits unutilized for the second 8 weeks.',
          section: s,
          block: b,
          building: b.building,
          buildingDesc: b.buildingDesc,
          room: b.room,
          days: b.days || [],
          daysStr: (b.days || []).map(d => DAY_MAP[d]?.name || d).join(', '),
          timeSpanStr: `${b.startTimeStr} – ${b.endTimeStr}`
        });
      }
    });
  });

  ft2Sections.forEach(s => {
    const blocks = getRelevantBlocks(s);
    blocks.forEach(b => {
      const isMeetingPaired = pairs.some(p => p.sec2.sectionUniqueId === s.sectionUniqueId && p.block2.rowId === b.rowId);
      if (!isMeetingPaired) {
        unpairedList.push({
          type: 'FT2',
          sessionLabel: 'Fast-Track 2 (Second 8-Weeks)',
          missingLabel: 'Missing Fast-Track 1 Companion',
          missingDesc: 'Scheduled in the second 8 weeks without a complementary Fast-Track 1 course in this room & timeslot. The classroom sits unutilized for the first 8 weeks.',
          section: s,
          block: b,
          building: b.building,
          buildingDesc: b.buildingDesc,
          room: b.room,
          days: b.days || [],
          daysStr: (b.days || []).map(d => DAY_MAP[d]?.name || d).join(', '),
          timeSpanStr: `${b.startTimeStr} – ${b.endTimeStr}`
        });
      }
    });
  });

  const optPairs = pairs.filter(p => p.status === 'OPTIMIZED');
  const unoptPairs = pairs.filter(p => p.status === 'UNOPTIMIZED');

  return {
    totalFtSections: ftSections.length,
    totalFt1Sections: ft1Sections.length,
    totalFt2Sections: ft2Sections.length,
    pairs: pairs,
    optPairs: optPairs,
    unoptPairs: unoptPairs,
    unpairedList: unpairedList
  };
}

function renderFastTrackCoverageDashboard() {
  const container = document.getElementById('fasttrack-coverage-container');
  if (!container) return;
  container.innerHTML = '';

  const scopeParts = [];
  if (appState.filters.location) scopeParts.push(appState.filters.location);
  else scopeParts.push('All Campuses');

  if (appState.filters.building) scopeParts.push(`Bldg ${appState.filters.building}`);
  if (appState.filters.room) scopeParts.push(`Room ${appState.filters.room}`);
  if (appState.filters.subject.size > 0) scopeParts.push(`Subject: ${Array.from(appState.filters.subject).join(', ')}`);
  if (appState.filters.course.size > 0) scopeParts.push(`Course: ${Array.from(appState.filters.course).join(', ')}`);
  if (appState.filters.acc.size > 0) scopeParts.push(`ACC: ${Array.from(appState.filters.acc).join(', ')}`);
  if (appState.filters.modality.size > 0) scopeParts.push(`Modality: ${Array.from(appState.filters.modality).join(', ')}`);
  if (appState.filters.search) scopeParts.push(`"${appState.filters.search}"`);

  const scopeLabel = scopeParts.join(' • ');

  const ftAnalysis = appState.currentFtAnalysis || analyzeFastTrackOfferings(appState.filteredSections || appState.sections, appState.filteredBlocks);

  // 1. Header Row
  const headerRow = document.createElement('div');
  headerRow.className = 'fasttrack-coverage-header';
  headerRow.innerHTML = `
    <div class="fasttrack-coverage-title">
      <h3>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#ea580c" stroke-width="2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
        Fast-Track Sequential Pairing Audit: <span>${scopeLabel}</span>
      </h3>
      <p>Auditing Fast-Track 1 & 2 strategic pairings, cross-GE coverage, cohort sequencing, and unpaired offerings.</p>
    </div>
    <div class="fasttrack-header-controls">
      <div class="sub-layout-switcher" id="fasttrack-sub-layout-switcher">
        <span class="sub-layout-label">Grid View:</span>
        <div class="btn-sub-toggle-group">
          <button type="button" class="btn-sub-layout-toggle ${appState.fasttrackSubLayout === 'day' ? 'active' : ''}" data-sublayout="day" title="View as Standard Day Grid">Day View</button>
          <button type="button" class="btn-sub-layout-toggle ${appState.fasttrackSubLayout === 'room' ? 'active' : ''}" data-sublayout="room" title="View as Stacked Room Grids">Room View</button>
        </div>
      </div>
    </div>
  `;
  container.appendChild(headerRow);

  // Sub-layout switcher listener
  headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      appState.fasttrackSubLayout = btn.getAttribute('data-sublayout');
      headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      dispatchViewRender();
    });
  });

  // 2. Summary KPI Cards
  const summaryBar = document.createElement('div');
  summaryBar.className = 'fasttrack-summary-bar';
  summaryBar.innerHTML = `
    <div class="fasttrack-stat-card highlight-card">
      <span class="fasttrack-stat-card-title">Fast-Track Sections</span>
      <span class="fasttrack-stat-card-value">${ftAnalysis.totalFtSections}</span>
      <span class="fasttrack-stat-card-sub">${ftAnalysis.totalFt1Sections} FT1 &bull; ${ftAnalysis.totalFt2Sections} FT2</span>
    </div>
    <div class="fasttrack-stat-card">
      <span class="fasttrack-stat-card-title">Pairings Formed</span>
      <span class="fasttrack-stat-card-value">${ftAnalysis.pairs.length}</span>
      <span class="fasttrack-stat-card-sub">Matching room & timeslot</span>
    </div>
    <div class="fasttrack-stat-card opt-card">
      <span class="fasttrack-stat-card-title">Optimized Pairings</span>
      <span class="fasttrack-stat-card-value">${ftAnalysis.optPairs.length}</span>
      <span class="fasttrack-stat-card-sub">Cross-GE & sequential cohorts</span>
    </div>
    <div class="fasttrack-stat-card unopt-card">
      <span class="fasttrack-stat-card-title">Unoptimized Pairings</span>
      <span class="fasttrack-stat-card-value">${ftAnalysis.unoptPairs.length}</span>
      <span class="fasttrack-stat-card-sub">Redundant GE or duplicate courses</span>
    </div>
    <div class="fasttrack-stat-card unpaired-card">
      <span class="fasttrack-stat-card-title">Unpaired Offerings</span>
      <span class="fasttrack-stat-card-value">${ftAnalysis.unpairedList.length}</span>
      <span class="fasttrack-stat-card-sub">Single 8-week offerings</span>
    </div>
  `;
  container.appendChild(summaryBar);

  // 3. Tab Filter Bar
  const totalItemsCount = ftAnalysis.pairs.length + ftAnalysis.unpairedList.length;
  const activeTab = appState.fasttrackActiveTab || 'all';

  const navBar = document.createElement('div');
  navBar.className = 'fasttrack-audit-nav';
  navBar.innerHTML = `
    <div class="fasttrack-audit-tabs">
      <button type="button" class="ft-tab-btn ${activeTab === 'all' ? 'active' : ''}" data-tab="all">
        <span>All Items</span>
        <span class="ft-tab-badge">${totalItemsCount}</span>
      </button>
      <button type="button" class="ft-tab-btn tab-opt ${activeTab === 'optimized' ? 'active' : ''}" data-tab="optimized">
        <span>✓ Optimized Pairings</span>
        <span class="ft-tab-badge">${ftAnalysis.optPairs.length}</span>
      </button>
      <button type="button" class="ft-tab-btn tab-unopt ${activeTab === 'unoptimized' ? 'active' : ''}" data-tab="unoptimized">
        <span>⚠ Unoptimized Pairings</span>
        <span class="ft-tab-badge">${ftAnalysis.unoptPairs.length}</span>
      </button>
      <button type="button" class="ft-tab-btn tab-unpaired ${activeTab === 'unpaired' ? 'active' : ''}" data-tab="unpaired">
        <span>⚡ Unpaired Offerings</span>
        <span class="ft-tab-badge">${ftAnalysis.unpairedList.length}</span>
      </button>
    </div>
    <div style="font-size:11.5px; color:var(--text-muted);">
      Click any course card to inspect full section details & meetings.
    </div>
  `;
  container.appendChild(navBar);

  // Tab switching listener
  navBar.querySelectorAll('.ft-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      appState.fasttrackActiveTab = btn.getAttribute('data-tab');
      applyFiltersAndRender();
    });
  });

  // 4. Cards Grid
  const cardsGrid = document.createElement('div');
  cardsGrid.className = 'fasttrack-cards-grid';

  let itemsRenderedCount = 0;

  // Render Pairings
  if (activeTab === 'all' || activeTab === 'optimized' || activeTab === 'unoptimized') {
    const pairsToRender = ftAnalysis.pairs.filter(p => {
      if (activeTab === 'optimized') return p.status === 'OPTIMIZED';
      if (activeTab === 'unoptimized') return p.status === 'UNOPTIMIZED';
      return true;
    });

    pairsToRender.forEach(p => {
      itemsRenderedCount++;
      const isOpt = p.status === 'OPTIMIZED';
      const card = document.createElement('div');
      card.className = 'ft-pair-card';

      const ge1Badges = (p.sec1.generalEdArea || []).map(a => `<span class="ge-area-badge ${getGeAreaBadgeClass(a)}">Area ${a}</span>`).join(' ');
      const ge2Badges = (p.sec2.generalEdArea || []).map(a => `<span class="ge-area-badge ${getGeAreaBadgeClass(a)}">Area ${a}</span>`).join(' ');

      card.innerHTML = `
        <div class="ft-card-top-bar">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="ft-room-badge">${p.building} ${p.room}</span>
            <span class="ft-days-time">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              ${p.daysStr} &bull; ${p.timeSpanStr}
            </span>
          </div>
          <span class="ft-status-pill ${isOpt ? 'status-optimized' : 'status-unoptimized'}">
            ${isOpt ? '✓ Optimized Pairing' : '⚠ Unoptimized Pairing'}
          </span>
        </div>
        <div class="ft-reason-banner ${isOpt ? 'banner-optimized' : 'banner-unoptimized'}">
          <span>${isOpt ? '✓' : '⚠'}</span>
          <span>${p.reasonText}</span>
        </div>
        <div class="ft-sections-split">
          <div class="ft-slot-card ft-slot-1" data-sec-id="${p.sec1.sectionUniqueId}" title="Click to view ${p.sec1.course} section details">
            <span class="ft-slot-tag tag-ft1">Fast-Track 1 (1st 8 Wks)</span>
            <div class="ft-slot-course">${p.sec1.course}</div>
            <div class="ft-slot-title">${p.sec1.sectionName}</div>
            <div class="ft-slot-meta">
              <span><strong>Faculty:</strong> ${p.sec1.facultyName || 'Staff'}</span>
              <span><strong>Time:</strong> ${p.block1.startTimeStr} – ${p.block1.endTimeStr}</span>
            </div>
            ${ge1Badges ? `<div class="ft-slot-badges">${ge1Badges}</div>` : ''}
          </div>
          <div class="ft-split-arrow">➔</div>
          <div class="ft-slot-card ft-slot-2" data-sec-id="${p.sec2.sectionUniqueId}" title="Click to view ${p.sec2.course} section details">
            <span class="ft-slot-tag tag-ft2">Fast-Track 2 (2nd 8 Wks)</span>
            <div class="ft-slot-course">${p.sec2.course}</div>
            <div class="ft-slot-title">${p.sec2.sectionName}</div>
            <div class="ft-slot-meta">
              <span><strong>Faculty:</strong> ${p.sec2.facultyName || 'Staff'}</span>
              <span><strong>Time:</strong> ${p.block2.startTimeStr} – ${p.block2.endTimeStr}</span>
            </div>
            ${ge2Badges ? `<div class="ft-slot-badges">${ge2Badges}</div>` : ''}
          </div>
        </div>
      `;

      cardsGrid.appendChild(card);
    });
  }

  // Render Unpaired Offerings
  if (activeTab === 'all' || activeTab === 'unpaired') {
    ftAnalysis.unpairedList.forEach(u => {
      itemsRenderedCount++;
      const card = document.createElement('div');
      card.className = 'ft-unpaired-card';

      const geBadges = (u.section.generalEdArea || []).map(a => `<span class="ge-area-badge ${getGeAreaBadgeClass(a)}">Area ${a}</span>`).join(' ');

      card.innerHTML = `
        <div class="ft-card-top-bar">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="ft-room-badge">${u.building} ${u.room}</span>
            <span class="ft-days-time">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              ${u.daysStr} &bull; ${u.timeSpanStr}
            </span>
          </div>
          <span class="ft-status-pill status-unpaired">⚠ ${u.missingLabel}</span>
        </div>
        <div class="ft-reason-banner banner-unpaired">
          <span>⚠</span>
          <span>${u.missingDesc}</span>
        </div>
        <div class="ft-unpaired-body">
          <div class="ft-slot-card" data-sec-id="${u.section.sectionUniqueId}" title="Click to view ${u.section.course} section details">
            <span class="ft-slot-tag ${u.type === 'FT1' ? 'tag-ft1' : 'tag-ft2'}">${u.sessionLabel}</span>
            <div class="ft-slot-course">${u.section.course}</div>
            <div class="ft-slot-title">${u.section.sectionName}</div>
            <div class="ft-slot-meta">
              <span><strong>Faculty:</strong> ${u.section.facultyName || 'Staff'}</span>
              <span><strong>Scheduled:</strong> ${u.daysStr} &bull; ${u.timeSpanStr}</span>
              <span><strong>Room:</strong> ${u.building} ${u.room} ${u.buildingDesc ? `(${u.buildingDesc})` : ''}</span>
            </div>
            ${geBadges ? `<div class="ft-slot-badges">${geBadges}</div>` : ''}
          </div>
        </div>
      `;

      cardsGrid.appendChild(card);
    });
  }

  if (itemsRenderedCount === 0) {
    cardsGrid.innerHTML = `
      <div style="grid-column: 1 / -1; padding: 36px 20px; text-align: center; background: var(--bg-surface); border: 1px dashed var(--border-color); border-radius: 8px; color: var(--text-secondary); font-size: 13.5px;">
        <span style="font-size: 22px; display: block; margin-bottom: 8px;">🎉</span>
        <strong>No Fast-Track items found for this view filter (${activeTab}) in scope: ${scopeLabel}.</strong>
        <p style="margin-top: 6px; font-size: 12px; color: var(--text-muted);">Adjust your sidebar filters or switch to 'All Items' tab above.</p>
      </div>
    `;
  }

  container.appendChild(cardsGrid);

  // Attach hover and click listeners to all section slots in Fast-Track Coverage
  container.querySelectorAll('.ft-slot-card').forEach(slot => {
    const secId = slot.getAttribute('data-sec-id');
    const sec = appState.sections.find(s => s.sectionUniqueId === secId);
    if (sec) {
      slot.addEventListener('mouseenter', (e) => showPopover(e, sec));
      slot.addEventListener('mousemove', (e) => positionPopover(e));
      slot.addEventListener('mouseleave', hidePopover);
      slot.addEventListener('click', () => showSectionModal(sec));
    }
  });
}

// --- 8B. ENROLLMENT EFFICIENCY & PRODUCTIVITY AUDIT DASHBOARD ---

/**
 * Analyzes enrollment efficiency metrics across the current filtered set of sections:
 * - Institutional Low Enrolled evaluations
 * - Aggregate capacity, active students, and overall fill rate
 * - Generated vs. Lost FTES calculations
 * - Fill Rate Tier distributions
 * - Academic Community (ACC) efficiency breakdowns
 */
function analyzeEnrollmentEfficiency(baseSections, activeBlocks) {
  const sections = baseSections || [];
  const totalSections = sections.length;

  let totalCapacity = 0;
  let totalActive = 0;
  let totalEstimatedFtes = 0;
  let totalLostFtes = 0;

  const lowEnrolledSections = [];
  const healthySections = [];
  const criticalSections = [];
  const highLostFtesSections = [];

  let tierOptCount = 0;
  let tierHighCount = 0;
  let tierMedCount = 0;
  let tierLowCount = 0;

  const accMap = new Map();

  sections.forEach(sec => {
    const cap = Number(sec.capacity) || 0;
    const act = Number(sec.activeStudents) || 0;
    const ftes = Number(sec.estimatedFtes) || 0;
    const lost = Number(sec.lostFtes) || 0;
    const fillRate = cap > 0 ? (act / cap) : 0;

    totalCapacity += cap;
    totalActive += act;
    totalEstimatedFtes += ftes;
    totalLostFtes += lost;

    if (sec.isLowEnrolled) {
      lowEnrolledSections.push(sec);
    }
    if (cap > 0 && fillRate >= 0.80) {
      healthySections.push(sec);
      tierOptCount++;
    } else if (cap > 0 && fillRate >= 0.65) {
      tierHighCount++;
    } else if (cap > 0 && fillRate >= 0.50) {
      tierMedCount++;
    } else {
      tierLowCount++;
    }

    if (cap > 0 && fillRate < 0.50) {
      criticalSections.push(sec);
    }

    if (lost >= 2.0) {
      highLostFtesSections.push(sec);
    }

    const accKey = sec.acc || 'Unassigned';
    if (!accMap.has(accKey)) {
      accMap.set(accKey, {
        acc: accKey,
        sectionsCount: 0,
        capacity: 0,
        active: 0,
        estimatedFtes: 0,
        lostFtes: 0,
        lowEnrolledCount: 0
      });
    }
    const accItem = accMap.get(accKey);
    accItem.sectionsCount++;
    accItem.capacity += cap;
    accItem.active += act;
    accItem.estimatedFtes += ftes;
    accItem.lostFtes += lost;
    if (sec.isLowEnrolled) {
      accItem.lowEnrolledCount++;
    }
  });

  const overallFillRate = totalCapacity > 0 ? (totalActive / totalCapacity) * 100 : 0;
  const totalPotentialFtes = totalEstimatedFtes + totalLostFtes;
  const ftesRealizationRate = totalPotentialFtes > 0 ? (totalEstimatedFtes / totalPotentialFtes) * 100 : 0;

  const accBreakdown = Array.from(accMap.values()).map(item => {
    item.fillRate = item.capacity > 0 ? (item.active / item.capacity) * 100 : 0;
    return item;
  }).sort((a, b) => b.sectionsCount - a.sectionsCount);

  return {
    totalSections,
    totalCapacity,
    totalActive,
    overallFillRate,
    totalEstimatedFtes,
    totalLostFtes,
    totalPotentialFtes,
    ftesRealizationRate,
    lowEnrolledSections,
    healthySections,
    criticalSections,
    highLostFtesSections,
    tiers: {
      optimal: { count: tierOptCount, pct: totalSections > 0 ? (tierOptCount / totalSections) * 100 : 0 },
      high: { count: tierHighCount, pct: totalSections > 0 ? (tierHighCount / totalSections) * 100 : 0 },
      moderate: { count: tierMedCount, pct: totalSections > 0 ? (tierMedCount / totalSections) * 100 : 0 },
      low: { count: tierLowCount, pct: totalSections > 0 ? (tierLowCount / totalSections) * 100 : 0 }
    },
    accBreakdown
  };
}

/**
 * Renders the Enrollment Efficiency & Productivity Audit Dashboard:
 * - Executive KPI bar (Low Enrolled count, Fill Rate %, FTES, Lost FTES, Potential FTES)
 * - Fill Rate Tier progress bars and ACC efficiency breakdown table
 * - Filter tabs (All, Low Enrolled, Healthy, Critical, High Lost FTES)
 * - Section cards with fill gauges, FTES metrics, and low-enrolled reason banners
 * - Synchronizes directly with the underlying Room View and Day View grids
 */
function renderEfficiencyDashboard() {
  const container = document.getElementById('efficiency-container');
  if (!container) return;
  container.innerHTML = '';

  const effAnalysis = appState.currentEfficiencyAnalysis || analyzeEnrollmentEfficiency(appState.sections, appState.scheduleBlocks);

  // Scope label for current filters
  const scopeParts = [];
  if (appState.filters.location) scopeParts.push(appState.filters.location);
  else scopeParts.push('All Campuses');

  if (appState.filters.building) scopeParts.push(`Bldg ${appState.filters.building}`);
  if (appState.filters.room) scopeParts.push(`Room ${appState.filters.room}`);
  if (appState.filters.acc.size > 0) scopeParts.push(`ACC: ${Array.from(appState.filters.acc).join(', ')}`);
  if (appState.filters.subject.size > 0) scopeParts.push(`Subject: ${Array.from(appState.filters.subject).join(', ')}`);
  if (appState.filters.course.size > 0) scopeParts.push(`Course: ${Array.from(appState.filters.course).join(', ')}`);
  if (appState.filters.dayEvening.size > 0) scopeParts.push(`Type: ${Array.from(appState.filters.dayEvening).map(formatDayEveningLabel).join(', ')}`);
  if (appState.filters.modality.size > 0) scopeParts.push(`Modality: ${Array.from(appState.filters.modality).join(', ')}`);
  if (appState.filters.session.size > 0) scopeParts.push(`Session: ${Array.from(appState.filters.session).join(', ')}`);
  if (appState.filters.search) scopeParts.push(`"${appState.filters.search}"`);

  const scopeLabel = scopeParts.join(' • ');

  // 1. Header Row
  const headerRow = document.createElement('div');
  headerRow.className = 'efficiency-header';
  headerRow.innerHTML = `
    <div class="efficiency-title">
      <h3>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
        Enrollment Efficiency & Productivity Audit: <span>${scopeLabel}</span>
      </h3>
      <p>Analyze institutional low-enrolled thresholds, fill rates, generated vs. lost FTES, and academic community productivity.</p>
    </div>
    <div class="efficiency-header-controls">
      <div class="sub-layout-switcher" id="efficiency-sub-layout-switcher">
        <span class="sub-layout-label">Grid View:</span>
        <div class="btn-sub-toggle-group">
          <button type="button" class="btn-sub-layout-toggle ${appState.efficiencySubLayout === 'day' ? 'active' : ''}" data-sublayout="day" title="View as Standard Day Grid">Day View</button>
          <button type="button" class="btn-sub-layout-toggle ${appState.efficiencySubLayout === 'room' ? 'active' : ''}" data-sublayout="room" title="View as Stacked Room Grids">Room View</button>
        </div>
      </div>
    </div>
  `;
  container.appendChild(headerRow);

  // Sub-layout switcher listener
  headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      appState.efficiencySubLayout = btn.getAttribute('data-sublayout');
      headerRow.querySelectorAll('.btn-sub-layout-toggle').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      dispatchViewRender();
    });
  });

  // 2. Summary KPI Cards
  const lowEnrolledPct = effAnalysis.totalSections > 0
    ? ((effAnalysis.lowEnrolledSections.length / effAnalysis.totalSections) * 100).toFixed(1)
    : '0.0';
  const unfilledSeats = Math.max(0, effAnalysis.totalCapacity - effAnalysis.totalActive);

  const summaryBar = document.createElement('div');
  summaryBar.className = 'efficiency-summary-bar';
  summaryBar.innerHTML = `
    <div class="eff-stat-card alert-card">
      <span class="eff-stat-card-title">Low Enrolled Alert</span>
      <span class="eff-stat-card-value">${effAnalysis.lowEnrolledSections.length}</span>
      <span class="eff-stat-card-sub">${lowEnrolledPct}% of offerings &bull; Institutional criteria</span>
    </div>
    <div class="eff-stat-card fill-card">
      <span class="eff-stat-card-title">Average Fill Rate</span>
      <span class="eff-stat-card-value">${effAnalysis.overallFillRate.toFixed(1)}%</span>
      <span class="eff-stat-card-sub">${effAnalysis.totalActive.toLocaleString()} enrolled / ${effAnalysis.totalCapacity.toLocaleString()} seats</span>
    </div>
    <div class="eff-stat-card ftes-card">
      <span class="eff-stat-card-title">Estimated FTES</span>
      <span class="eff-stat-card-value">${effAnalysis.totalEstimatedFtes.toFixed(1)}</span>
      <span class="eff-stat-card-sub">Active census productivity</span>
    </div>
    <div class="eff-stat-card lost-card">
      <span class="eff-stat-card-title">Lost / Missing FTES</span>
      <span class="eff-stat-card-value">${effAnalysis.totalLostFtes.toFixed(1)}</span>
      <span class="eff-stat-card-sub">${unfilledSeats.toLocaleString()} unfilled seats across sections</span>
    </div>
    <div class="eff-stat-card potential-card">
      <span class="eff-stat-card-title">Potential FTES</span>
      <span class="eff-stat-card-value">${effAnalysis.totalPotentialFtes.toFixed(1)}</span>
      <span class="eff-stat-card-sub">${effAnalysis.ftesRealizationRate.toFixed(1)}% realization rate</span>
    </div>
  `;
  container.appendChild(summaryBar);

  // 3. Analytics Row (Tiers & ACC Breakdown)
  const analyticsRow = document.createElement('div');
  analyticsRow.className = 'efficiency-analytics-row';

  // Tier progress panel
  const tierPanel = document.createElement('div');
  tierPanel.className = 'efficiency-panel';
  tierPanel.innerHTML = `
    <div class="efficiency-panel-title">
      <span>Fill Rate Distribution</span>
      <span style="font-size:11px; font-weight:600; color:var(--text-muted);">${effAnalysis.totalSections} Total Sections</span>
    </div>
    <div class="tier-bar-row">
      <span class="tier-label">Full / Optimal (≥80%)</span>
      <div class="tier-bar-track">
        <div class="tier-bar-fill fill-opt" style="width: ${effAnalysis.tiers.optimal.pct.toFixed(1)}%;"></div>
      </div>
      <span class="tier-stats">${effAnalysis.tiers.optimal.count} (${effAnalysis.tiers.optimal.pct.toFixed(0)}%)</span>
    </div>
    <div class="tier-bar-row">
      <span class="tier-label">High Demand (65–79%)</span>
      <div class="tier-bar-track">
        <div class="tier-bar-fill fill-high" style="width: ${effAnalysis.tiers.high.pct.toFixed(1)}%;"></div>
      </div>
      <span class="tier-stats">${effAnalysis.tiers.high.count} (${effAnalysis.tiers.high.pct.toFixed(0)}%)</span>
    </div>
    <div class="tier-bar-row">
      <span class="tier-label">Moderate (50–64%)</span>
      <div class="tier-bar-track">
        <div class="tier-bar-fill fill-med" style="width: ${effAnalysis.tiers.moderate.pct.toFixed(1)}%;"></div>
      </div>
      <span class="tier-stats">${effAnalysis.tiers.moderate.count} (${effAnalysis.tiers.moderate.pct.toFixed(0)}%)</span>
    </div>
    <div class="tier-bar-row">
      <span class="tier-label">Critical / Underfilled (&lt;50%)</span>
      <div class="tier-bar-track">
        <div class="tier-bar-fill fill-low" style="width: ${effAnalysis.tiers.low.pct.toFixed(1)}%;"></div>
      </div>
      <span class="tier-stats">${effAnalysis.tiers.low.count} (${effAnalysis.tiers.low.pct.toFixed(0)}%)</span>
    </div>
  `;

  // ACC Breakdown Panel
  const accPanel = document.createElement('div');
  accPanel.className = 'efficiency-panel';
  const accRowsHtml = effAnalysis.accBreakdown.map(item => `
    <tr>
      <td style="font-weight:700; max-width:160px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${item.acc}">${item.acc}</td>
      <td style="text-align:right;">${item.sectionsCount}</td>
      <td style="text-align:right;"><strong>${item.fillRate.toFixed(1)}%</strong></td>
      <td style="text-align:right;">${item.estimatedFtes.toFixed(1)}</td>
      <td style="text-align:right; color:#d97706; font-weight:600;">${item.lostFtes.toFixed(1)}</td>
      <td style="text-align:right;">
        ${item.lowEnrolledCount > 0 ? `<span style="color:#dc2626; font-weight:700;">⚠ ${item.lowEnrolledCount}</span>` : `<span style="color:#059669;">0</span>`}
      </td>
    </tr>
  `).join('');

  accPanel.innerHTML = `
    <div class="efficiency-panel-title">
      <span>Academic Community (ACC) Breakdown</span>
      <span style="font-size:11px; font-weight:600; color:var(--text-muted);">${effAnalysis.accBreakdown.length} Communities</span>
    </div>
    <div class="efficiency-table-container">
      <table class="efficiency-table">
        <thead>
          <tr>
            <th>ACC</th>
            <th style="text-align:right;">Secs</th>
            <th style="text-align:right;">Fill %</th>
            <th style="text-align:right;">Est FTES</th>
            <th style="text-align:right;">Lost FTES</th>
            <th style="text-align:right;">Low Enr</th>
          </tr>
        </thead>
        <tbody>
          ${accRowsHtml || '<tr><td colspan="6" style="text-align:center; color:var(--text-muted);">No ACC data</td></tr>'}
        </tbody>
      </table>
    </div>
  `;

  analyticsRow.appendChild(tierPanel);
  analyticsRow.appendChild(accPanel);
  container.appendChild(analyticsRow);

  // 4. Tab Navigation Bar
  const activeTab = appState.efficiencyActiveTab || 'all';
  const navBar = document.createElement('div');
  navBar.className = 'efficiency-audit-nav';
  navBar.innerHTML = `
    <div class="efficiency-audit-tabs">
      <button type="button" class="eff-tab-btn ${activeTab === 'all' ? 'active' : ''}" data-tab="all">
        <span>All Sections</span>
        <span class="eff-tab-badge">${effAnalysis.totalSections}</span>
      </button>
      <button type="button" class="eff-tab-btn tab-low ${activeTab === 'low_enrolled' ? 'active' : ''}" data-tab="low_enrolled">
        <span>⚠ Low Enrolled Alert</span>
        <span class="eff-tab-badge">${effAnalysis.lowEnrolledSections.length}</span>
      </button>
      <button type="button" class="eff-tab-btn tab-healthy ${activeTab === 'healthy' ? 'active' : ''}" data-tab="healthy">
        <span>✓ Healthy Fill (≥80%)</span>
        <span class="eff-tab-badge">${effAnalysis.healthySections.length}</span>
      </button>
      <button type="button" class="eff-tab-btn tab-critical ${activeTab === 'critical' ? 'active' : ''}" data-tab="critical">
        <span>⚡ Critical (&lt;50%)</span>
        <span class="eff-tab-badge">${effAnalysis.criticalSections.length}</span>
      </button>
      <button type="button" class="eff-tab-btn tab-lost ${activeTab === 'high_lost_ftes' ? 'active' : ''}" data-tab="high_lost_ftes">
        <span>📉 High Lost FTES (≥2.0)</span>
        <span class="eff-tab-badge">${effAnalysis.highLostFtesSections.length}</span>
      </button>
    </div>
    <div style="font-size:11.5px; color:var(--text-muted);">
      Click any course card to inspect details. Grid below synchronizes to filtered sections.
    </div>
  `;
  container.appendChild(navBar);

  // Tab switching listener
  navBar.querySelectorAll('.eff-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      appState.efficiencyActiveTab = btn.getAttribute('data-tab');
      applyFiltersAndRender();
    });
  });

  // 5. Section Cards Grid
  const cardsGrid = document.createElement('div');
  cardsGrid.className = 'efficiency-cards-grid';

  const sectionsToRender = appState.filteredSections || [];

  if (sectionsToRender.length === 0) {
    const emptyState = document.createElement('div');
    emptyState.style.cssText = 'grid-column: 1 / -1; padding: 36px 20px; text-align: center; color: var(--text-secondary); background: var(--bg-surface); border: 1px dashed var(--border-color); border-radius: 10px;';
    emptyState.innerHTML = `
      <div style="font-size: 24px; margin-bottom: 8px;">✓</div>
      <div style="font-weight: 700; font-size: 14px; color: var(--text-primary); margin-bottom: 4px;">No Sections in Selected Category</div>
      <div style="font-size: 12px; color: var(--text-muted);">There are no sections matching the "${activeTab.replace(/_/g, ' ')}" criteria under current filters.</div>
    `;
    cardsGrid.appendChild(emptyState);
  } else {
    const displayLimit = 120;
    const limitedSections = sectionsToRender.slice(0, displayLimit);

    limitedSections.forEach(sec => {
      const card = document.createElement('div');
      const isLow = sec.isLowEnrolled;
      const isHealthy = sec.fillRate >= 0.80;
      const isCrit = sec.fillRate < 0.50;

      card.className = `eff-section-card ${isLow ? 'is-low-enrolled' : (isHealthy ? 'is-healthy' : '')}`;

      const fillPct = Math.round(sec.fillRate * 100);
      let barColor = '#3b82f6';
      if (isLow || isCrit) barColor = '#ef4444';
      else if (isHealthy) barColor = '#10b981';
      else if (fillPct >= 65) barColor = '#3b82f6';
      else barColor = '#f59e0b';

      let statusBadge = '';
      if (isLow) {
        statusBadge = `<span class="eff-status-pill status-low">⚠ Low Enrolled</span>`;
      } else if (isHealthy) {
        statusBadge = `<span class="eff-status-pill status-healthy">✓ ${fillPct}% Full</span>`;
      } else if (isCrit) {
        statusBadge = `<span class="eff-status-pill status-low">⚡ ${fillPct}% Fill</span>`;
      } else {
        statusBadge = `<span class="eff-status-pill status-normal">● ${fillPct}% Fill</span>`;
      }

      card.innerHTML = `
        <div class="eff-card-top-bar">
          <div>
            <div class="eff-card-course">${sec.course}</div>
            <div style="font-size: 11px; color: var(--text-muted);">${sec.sectionName} (ID: ${sec.sectionUniqueId})</div>
            ${sec.title ? `<div class="eff-card-title">${sec.title}</div>` : ''}
          </div>
          <div>${statusBadge}</div>
        </div>

        ${isLow ? `
          <div class="eff-reason-banner">
            <span>⚠</span>
            <span>${getLowEnrolledReason(sec)}</span>
          </div>
        ` : ''}

        <div class="eff-fill-gauge-container">
          <div class="eff-fill-gauge-meta">
            <span>Enrolled / Capacity</span>
            <span><strong>${sec.activeStudents}</strong> / ${sec.capacity} (${fillPct}%)</span>
          </div>
          <div class="eff-fill-track">
            <div class="eff-fill-bar" style="width: ${Math.min(100, fillPct)}%; background-color: ${barColor};"></div>
          </div>
        </div>

        <div class="eff-metrics-row">
          <span class="eff-metric-tag"><strong>Faculty:</strong> ${sec.facultyName || 'Staff'}</span>
          <span class="eff-metric-tag"><strong>Location:</strong> ${sec.location}</span>
          <span class="eff-metric-tag"><strong>Est FTES:</strong> ${sec.estimatedFtes.toFixed(2)}</span>
          <span class="eff-metric-tag" style="color: ${sec.lostFtes >= 2.0 ? '#d97706' : 'var(--text-secondary)'};"><strong>Lost FTES:</strong> ${sec.lostFtes.toFixed(2)}</span>
          ${sec.acc ? `<span class="eff-metric-tag"><strong>ACC:</strong> ${sec.acc}</span>` : ''}
        </div>
      `;

      card.addEventListener('mouseenter', (e) => showPopover(e, sec));
      card.addEventListener('mousemove', (e) => positionPopover(e));
      card.addEventListener('mouseleave', hidePopover);
      card.addEventListener('click', () => showSectionModal(sec));

      cardsGrid.appendChild(card);
    });

    if (sectionsToRender.length > displayLimit) {
      const moreNote = document.createElement('div');
      moreNote.style.cssText = 'grid-column: 1 / -1; padding: 12px; text-align: center; font-size: 12px; color: var(--text-muted); background: var(--bg-surface-subtle); border-radius: 8px; border: 1px solid var(--border-color);';
      moreNote.textContent = `Showing first ${displayLimit} of ${sectionsToRender.length} sections. Use the sidebar filters to refine the list or view all in Room/Day View below.`;
      cardsGrid.appendChild(moreNote);
    }
  }

  container.appendChild(cardsGrid);
}

/**
 * Creates an individual Schedule Tile DOM Element with dedicated Course Code row
 * preventing text smashing, and support for Stacked Classes & GE Badges
 */
function createScheduleTileElement(block, leftPercent, widthPercent, topPx, heightPx) {
  const tile = document.createElement('div');
  const typeLower = block.instructionType.toLowerCase();

  let typeClass = 'type-lec';
  let shortType = 'LEC';
  if (typeLower.includes('lab')) {
    typeClass = 'type-lab';
    shortType = 'LAB';
  } else if (typeLower.includes('lec') && typeLower.includes('lab')) {
    typeClass = 'type-combo';
    shortType = 'LEC/LAB';
  }

  tile.className = `schedule-tile ${typeClass}`;
  tile.style.top = `${topPx}px`;
  tile.style.height = `${heightPx}px`;
  tile.style.left = `calc(${leftPercent}% + 2px)`;
  tile.style.width = `calc(${widthPercent}% - 4px)`;

  if (block.isStacked) {
    tile.classList.add('is-stacked');
  }

  if (appState.isConflictMode && block.isConflict) {
    tile.classList.add('is-conflict');
  }

  const sec = block.parentSection;

  // GE Area Badge for GE Coverage View (and when GE area is available)
  let geTagHtml = '';
  if (sec.geAssociations && sec.geAssociations.length > 0) {
    let displayArea = '';
    if (appState.layoutMode === 'ge' && appState.geActivePattern) {
      const activeAreas = sec.gePatternAreaMap && sec.gePatternAreaMap[appState.geActivePattern];
      if (activeAreas && activeAreas.size > 0) {
        displayArea = Array.from(activeAreas)[0];
      }
    }
    if (!displayArea && sec.generalEdArea.length > 0) {
      displayArea = sec.generalEdArea[0];
    }
    if (displayArea) {
      const colorClass = getGeAreaBadgeClass(displayArea);
      geTagHtml = `<span class="tile-ge-tag ${colorClass}">Area ${displayArea}</span>`;
    }
  }

  // Evening indicator tag
  let eveningTagHtml = '';
  if (isEveningBlock(block)) {
    let eveLabel = '🌙 Evening';
    if (block.day === 6 || block.day === 7) eveLabel = '📅 Weekend';
    else if (block.day === 5) eveLabel = '⚡ Friday';
    eveningTagHtml = `<span class="tile-ge-tag ge-color-1c" style="background:#4338ca;">${eveLabel}</span>`;
  }

  // Fast-track indicator tag
  let ftTagHtml = '';
  const ftType = getFastTrackSessionType(sec);
  if (ftType === 'FT1') {
    ftTagHtml = `<span class="tile-ft-tag ft1-tag" title="Fast-Track 1 (First 8-Weeks)">⚡ FT1</span>`;
  } else if (ftType === 'FT2') {
    ftTagHtml = `<span class="tile-ft-tag ft2-tag" title="Fast-Track 2 (Second 8-Weeks)">⚡ FT2</span>`;
  }

  const stackedTagHtml = block.isStacked ? `<span class="tile-stacked-tag">🔗 Stacked</span>` : '';
  const conflictTagHtml = (appState.isConflictMode && block.isConflict) ? `<span class="tile-conflict-alert-tag">⚠ CONFLICT</span>` : '';

  tile.innerHTML = `
    <div class="tile-course-row">
      <span class="tile-course-code" title="${sec.course} (${sec.sectionName})">${sec.course}</span>
    </div>
    <div class="tile-badges-row">
      <span class="tile-type-badge">${shortType}</span>
      ${appState.layoutMode === 'ge' ? geTagHtml : ''}
      ${appState.layoutMode === 'evening' ? eveningTagHtml : ''}
      ${ftTagHtml}
      ${stackedTagHtml}
      ${conflictTagHtml}
    </div>
    <div class="tile-body">
      <span class="tile-sec-name" title="${sec.sectionName}">${sec.sectionName}</span>
      <span class="tile-faculty" title="${sec.facultyName || 'Staff'}">${sec.facultyName || 'Staff'}</span>
    </div>
    <div class="tile-footer">
      <span class="tile-time">${block.startTimeStr} - ${block.endTimeStr}</span>
      <span class="tile-room-badge" title="${block.building} ${block.room}">${block.building} ${block.room}</span>
    </div>
  `;

  tile.addEventListener('mouseenter', (e) => showPopover(e, block));
  tile.addEventListener('mousemove', (e) => positionPopover(e));
  tile.addEventListener('mouseleave', hidePopover);
  tile.addEventListener('click', () => showSectionModal(sec, block));

  return tile;
}

// --- 8. HOVER POPOVER TOOLTIP ---

const popoverEl = document.getElementById('schedule-popover');

/**
 * Displays rich interactive popover on hover for either schedule tiles or audit course chips
 */
function showPopover(event, target) {
  if (!popoverEl || !target) return;
  const isBlock = target.startMinutes !== undefined && target.parentSection !== undefined;
  const sec = isBlock ? target.parentSection : target;
  const block = isBlock ? target : (sec.scheduleBlocks && sec.scheduleBlocks.length > 0 ? sec.scheduleBlocks[0] : null);

  const traitsHtml = (sec.sectionTraits || []).map(t => `<span class="trait-tag">${t}</span>`).join('');

  // Stacked class banner
  let stackedAlertHtml = '';
  if (block && block.isStacked && block.stackedWith && block.stackedWith.length > 0) {
    const other = block.stackedWith[0];
    const facA = sec.facultyName || 'Staff';
    const facB = other.parentSection.facultyName || 'Staff';
    const isSameFac = normalizeFacultyName(facA) === normalizeFacultyName(facB);
    const prog = getSectionStackSubject(sec);
    const isFt = isSequentialFastTrackPair(block, other);

    let stackNote = '';
    if (isFt) {
      stackNote = `Fast-Track Sequential Session (${sec.calendarSession} & ${other.parentSection.calendarSession} - Non-overlapping dates)`;
    } else if (isSameFac) {
      stackNote = `Same Instructor: ${facA}`;
    } else {
      stackNote = `Cohort Stack: ${prog || 'Specialized Program'} (Instructors: ${facA} & ${facB})`;
    }

    stackedAlertHtml = `
      <div style="margin-top:8px; padding:6px 8px; background-color:var(--stacked-bg); border:1px solid var(--stacked-border); border-radius:6px; color:var(--stacked-text); font-size:11px; font-weight:600;">
        🔗 <strong>STACKED ${isFt ? 'SESSION' : 'CLASS'}</strong>: Cross-scheduled in ${block.building} ${block.room} with <strong>${other.parentSection.course}</strong> (${stackNote}).
      </div>
    `;
  }

  // Conflict warning banner
  let conflictAlertHtml = '';
  if (block && block.isConflict && block.conflictsWith && block.conflictsWith.length > 0) {
    const other = block.conflictsWith[0];
    conflictAlertHtml = `
      <div class="popover-conflict-warning">
        ⚠ <strong>ROOM CONFLICT</strong>: Overlaps with <strong>${other.parentSection.sectionName}</strong> (${other.startTimeStr} - ${other.endTimeStr}) in ${block.building} ${block.room}.
      </div>
    `;
  }

  const isFtContext = appState.layoutMode === 'fasttrack' || isFastTrackSection(sec);
  const candidateBlocks = (isFtContext && sec.scheduleBlocks)
    ? (sec.scheduleBlocks.filter(isPhysicalMeetingBlock).length > 0 ? sec.scheduleBlocks.filter(isPhysicalMeetingBlock) : sec.scheduleBlocks)
    : (sec.scheduleBlocks || []);

  const meetingTypeStr = block ? block.instructionType : (candidateBlocks.map(b => b.instructionType).join(', ') || 'Class Meeting');
  const timeDayStr = block && block.day ? `${DAY_MAP[block.day]?.name || ''}: ${block.startTimeStr} – ${block.endTimeStr}` : (candidateBlocks.map(b => `${(b.days || []).map(d => DAY_MAP[d]?.short || d).join('/')} ${b.startTimeStr}-${b.endTimeStr}`).join('; ') || 'TBA');
  const locationStr = block && block.building ? `${sec.location} – ${block.building} ${block.room}` : (candidateBlocks.length > 0 && candidateBlocks[0].building ? `${sec.location} – ${candidateBlocks[0].building} ${candidateBlocks[0].room}` : `${sec.location}`);

  const geInfo = formatSectionGeInfo(sec);

  const lowEnrolledAlertHtml = sec.isLowEnrolled ? `
    <div style="margin-top:8px; padding:6px 8px; background-color:#fef2f2; border:1px solid #fee2e2; border-radius:6px; color:#991b1b; font-size:11px; font-weight:600; display:flex; align-items:flex-start; gap:6px;">
      <span>⚠</span>
      <span><strong>LOW ENROLLED:</strong> ${getLowEnrolledReason(sec)}</span>
    </div>
  ` : '';

  const enrollmentRowsHtml = (sec.capacity > 0 || sec.activeStudents > 0) ? `
    <span class="popover-field-label">Enrollment:</span>
    <span class="popover-field-value"><strong>${sec.activeStudents}</strong> / ${sec.capacity} (${Math.round(sec.fillRate * 100)}% Fill)</span>

    <span class="popover-field-label">Est. FTES:</span>
    <span class="popover-field-value">${sec.estimatedFtes.toFixed(2)} FTES ${sec.lostFtes > 0 ? `<span style="color:#d97706; font-size:10.5px;">(Lost: ${sec.lostFtes.toFixed(2)})</span>` : ''}</span>
  ` : '';

  popoverEl.innerHTML = `
    <div class="popover-header">
      <div class="popover-title-row">
        <span class="popover-course">${sec.course}</span>
        <span class="popover-id-pill">ID: ${sec.sectionUniqueId}</span>
      </div>
      ${sec.title ? `<div style="font-size:12px; font-weight:600; color:var(--chaffey-blue); margin-top:2px; line-height:1.25;">${sec.title}</div>` : ''}
      <div class="popover-sec-name">${sec.sectionName}</div>
    </div>
    <div class="popover-grid">
      <span class="popover-field-label">Subject:</span>
      <span class="popover-field-value">${sec.subject}</span>

      ${sec.acc ? `
        <span class="popover-field-label">ACC:</span>
        <span class="popover-field-value">${sec.acc}</span>
      ` : ''}

      <span class="popover-field-label">Faculty:</span>
      <span class="popover-field-value">${sec.facultyName || 'Staff'}</span>

      ${enrollmentRowsHtml}

      <span class="popover-field-label">Meeting:</span>
      <span class="popover-field-value">${meetingTypeStr}</span>

      <span class="popover-field-label">Time & Day:</span>
      <span class="popover-field-value">${timeDayStr}</span>

      <span class="popover-field-label">Location:</span>
      <span class="popover-field-value">${locationStr}</span>

      <span class="popover-field-label">Modality:</span>
      <span class="popover-field-value">${sec.modality || 'In-Person'} (${sec.calendarSession || 'Full-Term'})</span>

      <span class="popover-field-label">Day / Evening:</span>
      <span class="popover-field-value">${formatDayEveningLabel(sec.dayOrEvening)}</span>

      <span class="popover-field-label">Weeks:</span>
      <span class="popover-field-value">${sec.numberOfWeeks ? `${sec.numberOfWeeks} Weeks` : 'N/A'}</span>

      <span class="popover-field-label">GE Pattern:</span>
      <span class="popover-field-value">${geInfo.patternsFormatted}</span>

      <span class="popover-field-label">GE Area:</span>
      <span class="popover-field-value">${geInfo.combinedFormatted !== 'None' ? geInfo.combinedFormatted : 'None'}</span>
    </div>

    ${traitsHtml ? `
      <div class="popover-field-label" style="margin-top:6px;">Section Traits:</div>
      <div class="popover-traits-list">${traitsHtml}</div>
    ` : ''}

    ${lowEnrolledAlertHtml}
    ${stackedAlertHtml}
    ${conflictAlertHtml}
  `;

  popoverEl.style.display = 'block';
  positionPopover(event);
}

function positionPopover(event) {
  if (!popoverEl || popoverEl.style.display === 'none') return;
  const padding = 15;
  const popoverWidth = popoverEl.offsetWidth || 320;
  const popoverHeight = popoverEl.offsetHeight || 300;

  let x = event.clientX + padding;
  let y = event.clientY + padding;

  if (x + popoverWidth > window.innerWidth - padding) {
    x = event.clientX - popoverWidth - padding;
  }

  if (y + popoverHeight > window.innerHeight - padding) {
    y = event.clientY - popoverHeight - padding;
  }

  popoverEl.style.left = `${Math.max(padding, x)}px`;
  popoverEl.style.top = `${Math.max(padding, y)}px`;
}

function hidePopover() {
  if (popoverEl) popoverEl.style.display = 'none';
}

// --- 9. SECTION INSPECTOR MODAL ---

function showSectionModal(section, activeBlock) {
  const modal = document.getElementById('modal-details-backdrop');
  const modalSecId = document.getElementById('modal-sec-id');
  const modalSecName = document.getElementById('modal-sec-name');
  const modalBody = document.getElementById('modal-details-body');
  if (!modal || !modalBody) return;

  modalSecId.textContent = `SECTION UNIQUE ID: #${section.sectionUniqueId}`;
  modalSecName.textContent = section.title
    ? `${section.course} – ${section.title} (${section.sectionName})`
    : `${section.course} – ${section.sectionName}`;

  // For Fast-Track Coverage, ignore Scheduled Meeting Components or Parts that do not list days, times, locations, or hours,
  // and ignore components/parts listed as Online Lec; Canvas CANVAS (Online).
  const isFtMode = appState.layoutMode === 'fasttrack' || isFastTrackSection(section);
  const candidateBlocks = (isFtMode && section.scheduleBlocks)
    ? section.scheduleBlocks.filter(isPhysicalMeetingBlock)
    : (section.scheduleBlocks || []);
  const blocksToDisplay = (isFtMode && candidateBlocks.length > 0) ? candidateBlocks : (section.scheduleBlocks || []);

  const blocksHtml = blocksToDisplay.map((blk, idx) => {
    const daysStr = blk.days.map(d => DAY_MAP[d]?.name || d).join(', ');
    return `
      <div class="block-card">
        <div>
          <strong style="color:var(--text-primary); font-size:13px;">${blk.instructionType}</strong>
          <div style="color:var(--text-secondary); font-size:12px; margin-top:3px;">
            ${daysStr} &bull; ${blk.startTimeStr} – ${blk.endTimeStr}
          </div>
          <div style="color:var(--chaffey-blue); font-size:12px; font-weight:600; margin-top:2px;">
            Location: ${blk.building} ${blk.room} ${blk.buildingDesc ? `(${blk.buildingDesc})` : ''}
          </div>
        </div>
        <span class="badge" style="background:var(--chaffey-red); color:#fff; padding:3px 8px; border-radius:4px; font-size:11px;">
          Part ${idx + 1} of ${blocksToDisplay.length}
        </span>
      </div>
    `;
  }).join('');

  const traitsHtml = section.sectionTraits.map(t => `<span class="trait-tag">${t}</span>`).join('');

  const geInfo = formatSectionGeInfo(section);

  const lowEnrolledModalBanner = section.isLowEnrolled ? `
    <div style="margin-bottom:16px; padding:10px 14px; background-color:#fef2f2; border:1px solid #fecaca; border-left:4px solid #dc2626; border-radius:8px; color:#991b1b; font-size:12.5px; font-weight:600; display:flex; align-items:flex-start; gap:8px;">
      <span style="font-size:16px; line-height:1;">⚠</span>
      <div>
        <div style="font-weight:700; text-transform:uppercase; font-size:11px; letter-spacing:0.5px; color:#b91c1c;">Institutional Low Enrolled Alert</div>
        <div style="margin-top:2px;">${getLowEnrolledReason(section)}</div>
      </div>
    </div>
  ` : '';

  modalBody.innerHTML = `
    ${lowEnrolledModalBanner}
    <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:16px;">
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">SUBJECT & COURSE</div>
        <div style="font-size:13px; font-weight:600;">${section.subject} (${section.course})</div>
      </div>
      ${section.title ? `
        <div>
          <div style="font-size:11px; color:var(--text-muted); font-weight:600;">COURSE TITLE</div>
          <div style="font-size:13px; font-weight:600; color:var(--chaffey-blue);">${section.title}</div>
        </div>
      ` : ''}
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">INSTRUCTOR</div>
        <div style="font-size:13px; font-weight:600;">${section.facultyName || 'Staff'}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">CAMPUS LOCATION</div>
        <div style="font-size:13px; font-weight:600;">${section.location || 'N/A'}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">ACADEMIC COMMUNITY (ACC)</div>
        <div style="font-size:13px; font-weight:600;">${section.acc || 'N/A'}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">CALENDAR SESSION & WEEKS</div>
        <div style="font-size:13px; font-weight:600;">${section.calendarSession} (${section.numberOfWeeks} Weeks)</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">MODALITY</div>
        <div style="font-size:13px; font-weight:600;">${section.modality}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">ENROLLMENT & CAPACITY</div>
        <div style="font-size:13px; font-weight:600;">
          <strong>${section.activeStudents}</strong> / ${section.capacity} Seats (${Math.round(section.fillRate * 100)}% Fill)
        </div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">ESTIMATED FTES</div>
        <div style="font-size:13px; font-weight:600; color:#059669;">
          ${section.estimatedFtes.toFixed(2)} FTES
        </div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">LOST / UNFILLED FTES</div>
        <div style="font-size:13px; font-weight:600; color:${section.lostFtes >= 2.0 ? '#dc2626' : (section.lostFtes > 0 ? '#d97706' : 'var(--text-primary)')};">
          ${section.lostFtes.toFixed(2)} FTES (${section.unfilledSeats} unfilled seats)
        </div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">GENERAL ED PATTERN</div>
        <div style="font-size:13px; font-weight:600;">${geInfo.patternsFormatted}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">GENERAL ED AREA</div>
        <div style="font-size:13px; font-weight:600;">${geInfo.combinedFormatted !== 'None' ? geInfo.combinedFormatted : 'None'}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">DAY / EVENING</div>
        <div style="font-size:13px; font-weight:600;">${formatDayEveningLabel(section.dayOrEvening)}</div>
      </div>
    </div>

    <div style="margin-top:12px; border-top:1px solid var(--border-color); padding-top:12px;">
      <h4 style="font-size:12px; font-weight:700; color:var(--text-secondary); text-transform:uppercase; margin-bottom:8px;">
        Scheduled Meeting Components (${section.scheduleBlocks.length})
      </h4>
      <div class="blocks-list">
        ${blocksHtml}
      </div>
    </div>

    ${traitsHtml ? `
      <div style="margin-top:12px; border-top:1px solid var(--border-color); padding-top:12px;">
        <h4 style="font-size:12px; font-weight:700; color:var(--text-secondary); text-transform:uppercase; margin-bottom:6px;">
          Section Traits
        </h4>
        <div class="popover-traits-list">${traitsHtml}</div>
      </div>
    ` : ''}
  `;

  modal.style.display = 'flex';
}

function closeSectionModal() {
  const modal = document.getElementById('modal-details-backdrop');
  if (modal) modal.style.display = 'none';
}

// --- 10. CONFLICT CENTER MODAL ---

function openConflictCenterModal() {
  const modal = document.getElementById('modal-conflict-backdrop');
  const body = document.getElementById('modal-conflict-body');
  if (!modal || !body) return;

  if (appState.conflicts.length === 0) {
    body.innerHTML = `
      <div style="text-align:center; padding:30px; color:var(--text-muted);">
        <h3>No Room Conflicts Detected</h3>
        <p>There are no overlapping sections with different instructors in the same room on the same day.</p>
      </div>
    `;
  } else {
    body.innerHTML = appState.conflicts.map((c, idx) => {
      const dayName = DAY_MAP[c.day]?.name || c.day;
      const secA = c.blockA.parentSection;
      const secB = c.blockB.parentSection;

      return `
        <div class="conflict-item-card">
          <div class="conflict-item-header">
            <span>Conflict #${idx + 1}: ${c.building} ${c.room} on ${dayName}</span>
            <span>Overlap: ${formatMinutesToAmPm(c.overlapStart)} – ${formatMinutesToAmPm(c.overlapEnd)}</span>
          </div>
          <div class="conflict-parties">
            <div>
              <div style="font-size:11px; color:var(--chaffey-red); font-weight:700;">SECTION A (#${secA.sectionUniqueId})</div>
              <strong style="font-size:13px;">${secA.course} – ${secA.sectionName}</strong>
              <div style="color:var(--text-secondary); font-size:11.5px; margin-top:2px;">
                ${c.blockA.instructionType} &bull; ${c.blockA.startTimeStr} – ${c.blockA.endTimeStr}
              </div>
              <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">Instructor: ${secA.facultyName || 'Staff'}</div>
            </div>
            <div>
              <div style="font-size:11px; color:var(--chaffey-red); font-weight:700;">SECTION B (#${secB.sectionUniqueId})</div>
              <strong style="font-size:13px;">${secB.course} – ${secB.sectionName}</strong>
              <div style="color:var(--text-secondary); font-size:11.5px; margin-top:2px;">
                ${c.blockB.instructionType} &bull; ${c.blockB.startTimeStr} – ${c.blockB.endTimeStr}
              </div>
              <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">Instructor: ${secB.facultyName || 'Staff'}</div>
            </div>
          </div>
        </div>
      `;
    }).join('');
  }

  modal.style.display = 'flex';
}

function closeConflictCenterModal() {
  const modal = document.getElementById('modal-conflict-backdrop');
  if (modal) modal.style.display = 'none';
}

// --- 11. FILE UPLOAD & SHEETJS INTEGRATION ---

function handleFileSelect(file) {
  if (!file) return;

  showToast(`Parsing ${file.name}...`);

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      if (jsonRows.length === 0) {
        alert('The uploaded file contains no data rows.');
        hideToast();
        return;
      }

      processRawDataset(jsonRows);
      const uploadText = document.getElementById('btn-upload-text');
      if (uploadText) uploadText.textContent = 'Upload New File';
      showToast(`Successfully loaded ${jsonRows.length} rows from ${file.name}`);
      setTimeout(hideToast, 2500);
    } catch (err) {
      console.error('Error parsing spreadsheet:', err);
      alert(`Failed to parse spreadsheet file: ${err.message}`);
      hideToast();
    }
  };

  reader.readAsArrayBuffer(file);
}

// --- 12. EXPORT FUNCTIONALITY (PNG, PDF, CSV) ---

async function exportAsPng() {
  if (appState.sections.length === 0) {
    showToast('Please upload a schedule spreadsheet first.');
    setTimeout(hideToast, 2500);
    return;
  }

  const gridViewport = document.getElementById('schedule-grid-viewport');
  if (!gridViewport) return;

  showToast('Generating high-resolution PNG export...');

  try {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    const canvas = await html2canvas(gridViewport, {
      scale: 2,
      useCORS: true,
      logging: false,
      backgroundColor: isDark ? '#090d16' : '#ffffff'
    });

    const link = document.createElement('a');
    link.download = `Chaffey_Schedule_${appState.layoutMode.toUpperCase()}_${new Date().toISOString().slice(0, 10)}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();

    showToast('PNG downloaded successfully!');
    setTimeout(hideToast, 2000);
  } catch (err) {
    console.error('Error exporting PNG:', err);
    alert('Failed to generate PNG export.');
    hideToast();
  }
}

async function exportAsPdf() {
  if (appState.sections.length === 0) {
    showToast('Please upload a schedule spreadsheet first.');
    setTimeout(hideToast, 2500);
    return;
  }

  const gridViewport = document.getElementById('schedule-grid-viewport');
  if (!gridViewport || !window.jspdf) return;

  showToast('Generating PDF document...');

  try {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    const canvas = await html2canvas(gridViewport, {
      scale: 2,
      useCORS: true,
      logging: false,
      backgroundColor: isDark ? '#090d16' : '#ffffff'
    });

    const imgData = canvas.toDataURL('image/png');
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
      orientation: 'landscape',
      unit: 'pt',
      format: 'a4'
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    pdf.setFillColor(153, 0, 0);
    pdf.rect(0, 0, pageWidth, 40, 'F');

    pdf.setTextColor(255, 255, 255);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(14);
    pdf.text('CHAFFEY COLLEGE | Academic Schedule Grid', 20, 25);

    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    const dateStr = `Exported: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()} | View: ${appState.layoutMode.toUpperCase()} | Sections: ${appState.filteredBlocks.length}`;
    pdf.text(dateStr, pageWidth - 20, 25, { align: 'right' });

    const margin = 20;
    const availableWidth = pageWidth - (margin * 2);
    const availableHeight = pageHeight - 55 - margin;

    const imgWidth = availableWidth;
    const imgHeight = (canvas.height * availableWidth) / canvas.width;

    if (imgHeight <= availableHeight) {
      pdf.addImage(imgData, 'PNG', margin, 50, imgWidth, imgHeight);
    } else {
      const scaledWidth = (canvas.width * availableHeight) / canvas.height;
      pdf.addImage(imgData, 'PNG', (pageWidth - scaledWidth) / 2, 50, scaledWidth, availableHeight);
    }

    pdf.save(`Chaffey_Schedule_${appState.layoutMode.toUpperCase()}_${new Date().toISOString().slice(0, 10)}.pdf`);
    showToast('PDF downloaded successfully!');
    setTimeout(hideToast, 2000);
  } catch (err) {
    console.error('Error exporting PDF:', err);
    alert('Failed to generate PDF export.');
    hideToast();
  }
}

function exportAsCsv() {
  if (appState.sections.length === 0 || appState.filteredBlocks.length === 0) {
    showToast('Please upload a schedule spreadsheet first.');
    setTimeout(hideToast, 2500);
    return;
  }

  const headers = [
    'Section Unique ID', 'Course', 'Course Title', 'Section Name', 'Subject', 'ACC', 'Faculty Name',
    'Location', 'Building', 'Room', 'Day', 'Start Time', 'End Time',
    'Capacity', 'Active Students', 'Fill Rate %', 'Estimated FTES', 'Lost FTES', 'Low Enrolled Alert',
    'Instruction Type', 'General Ed Pattern', 'General Ed Area', 'Section Traits',
    'Is Stacked', 'Is Conflict'
  ];

  const rows = appState.filteredBlocks.map(b => {
    const sec = b.parentSection;
    const geInfo = formatSectionGeInfo(sec);
    const fillPct = sec.capacity > 0 ? (sec.fillRate * 100).toFixed(1) + '%' : '0.0%';
    return [
      `"${sec.sectionUniqueId}"`,
      `"${sec.course}"`,
      `"${sec.title || ''}"`,
      `"${sec.sectionName}"`,
      `"${sec.subject}"`,
      `"${sec.acc || ''}"`,
      `"${sec.facultyName}"`,
      `"${sec.location}"`,
      `"${b.building}"`,
      `"${b.room}"`,
      `"${DAY_MAP[b.day].name}"`,
      `"${b.startTimeStr}"`,
      `"${b.endTimeStr}"`,
      sec.capacity,
      sec.activeStudents,
      `"${fillPct}"`,
      sec.estimatedFtes.toFixed(2),
      sec.lostFtes.toFixed(2),
      `"${sec.isLowEnrolled ? 'YES' : 'NO'}"`,
      `"${b.instructionType}"`,
      `"${geInfo.patternsFormatted}"`,
      `"${geInfo.combinedFormatted !== 'None' ? geInfo.combinedFormatted : 'None'}"`,
      `"${sec.sectionTraitsRaw}"`,
      `"${b.isStacked ? 'YES' : 'NO'}"`,
      `"${b.isConflict ? 'YES' : 'NO'}"`
    ].join(',');
  });

  const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows].join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `Chaffey_Schedule_Filtered_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function showToast(text) {
  const toast = document.getElementById('export-toast');
  const toastText = document.getElementById('export-toast-text');
  if (toast && toastText) {
    toastText.textContent = text;
    toast.style.display = 'flex';
  }
}

function hideToast() {
  const toast = document.getElementById('export-toast');
  if (toast) toast.style.display = 'none';
}

// --- 13. UI EVENT LISTENERS & INITIALIZATION ---

function setupEventListeners() {
  const fileInput = document.getElementById('file-upload-input');
  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFileSelect(e.target.files[0]);
      }
    });
  }

  const overlay = document.getElementById('drag-drop-overlay');
  window.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (overlay) overlay.classList.add('active');
  });

  window.addEventListener('dragleave', (e) => {
    if (e.relatedTarget === null && overlay) {
      overlay.classList.remove('active');
    }
  });

  window.addEventListener('drop', (e) => {
    e.preventDefault();
    if (overlay) overlay.classList.remove('active');
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileSelect(e.dataTransfer.files[0]);
    }
  });

  // REQUIREMENT 1: When clicking View Room Conflicts, default to Isolate Conflicts Only!
  const btnConflicts = document.getElementById('btn-toggle-conflicts');
  if (btnConflicts) {
    btnConflicts.addEventListener('click', () => {
      appState.isConflictMode = !appState.isConflictMode;

      if (appState.isConflictMode) {
        // Default to Isolate Conflicts Only
        appState.conflictsOnly = true;

        const switcher = document.getElementById('conflict-mode-switcher');
        if (switcher) switcher.style.display = 'inline-flex';
        document.getElementById('btn-conflict-isolate')?.classList.add('active');
        document.getElementById('btn-conflict-context')?.classList.remove('active');

        if (appState.conflicts.length === 0) {
          showToast('0 Room Conflicts found. (Stacked classes with matching instructors are not counted as conflicts).');
          setTimeout(hideToast, 3500);
        }
      } else {
        appState.conflictsOnly = false;
        const switcher = document.getElementById('conflict-mode-switcher');
        if (switcher) switcher.style.display = 'none';
      }

      btnConflicts.classList.toggle('active', appState.isConflictMode);
      btnConflicts.setAttribute('aria-pressed', String(appState.isConflictMode));
      updateConflictUI();
      applyFiltersAndRender();
    });
  }

  // Conflict sub-switcher options (Isolate Conflicts vs Show in Context)
  document.getElementById('btn-conflict-isolate')?.addEventListener('click', () => {
    appState.conflictsOnly = true;
    document.getElementById('btn-conflict-isolate')?.classList.add('active');
    document.getElementById('btn-conflict-context')?.classList.remove('active');
    applyFiltersAndRender();
  });

  document.getElementById('btn-conflict-context')?.addEventListener('click', () => {
    appState.conflictsOnly = false;
    document.getElementById('btn-conflict-context')?.classList.add('active');
    document.getElementById('btn-conflict-isolate')?.classList.remove('active');
    applyFiltersAndRender();
  });

  // Conflict Center Modal
  document.getElementById('btn-open-conflict-modal')?.addEventListener('click', openConflictCenterModal);
  document.getElementById('btn-close-conflict-modal')?.addEventListener('click', closeConflictCenterModal);

  // Export Dropdown
  const btnExportMenu = document.getElementById('btn-export-menu');
  const exportDropdown = document.getElementById('export-dropdown-menu');
  if (btnExportMenu && exportDropdown) {
    btnExportMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      exportDropdown.classList.toggle('show');
    });

    document.addEventListener('click', () => {
      exportDropdown.classList.remove('show');
    });

    document.getElementById('btn-export-png')?.addEventListener('click', exportAsPng);
    document.getElementById('btn-export-pdf')?.addEventListener('click', exportAsPdf);
    document.getElementById('btn-export-csv')?.addEventListener('click', exportAsCsv);
  }

  // Light / Dark Mode Toggle
  const btnThemeToggle = document.getElementById('btn-theme-toggle');
  const iconSun = document.getElementById('icon-theme-sun');
  const iconMoon = document.getElementById('icon-theme-moon');

  const savedTheme = localStorage.getItem('chaffey_schedule_theme') || 'light';
  setTheme(savedTheme);

  if (btnThemeToggle) {
    btnThemeToggle.addEventListener('click', () => {
      const current = document.body.getAttribute('data-theme') || 'light';
      const next = current === 'light' ? 'dark' : 'light';
      setTheme(next);
      localStorage.setItem('chaffey_schedule_theme', next);
    });
  }

  function setTheme(theme) {
    document.body.setAttribute('data-theme', theme);
    if (theme === 'dark') {
      if (iconSun) iconSun.style.display = 'none';
      if (iconMoon) iconMoon.style.display = 'block';
    } else {
      if (iconSun) iconSun.style.display = 'block';
      if (iconMoon) iconMoon.style.display = 'none';
    }
  }

  // Quick Search Input
  const searchInput = document.getElementById('filter-search');
  const btnClearSearch = document.getElementById('btn-clear-search');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      appState.filters.search = e.target.value;
      if (btnClearSearch) btnClearSearch.style.display = e.target.value ? 'block' : 'none';
      applyFiltersAndRender();
    });
  }
  if (btnClearSearch) {
    btnClearSearch.addEventListener('click', () => {
      searchInput.value = '';
      appState.filters.search = '';
      btnClearSearch.style.display = 'none';
      applyFiltersAndRender();
    });
  }

  // Location Cascading Dropdown
  const locSelect = document.getElementById('filter-location');
  if (locSelect) {
    locSelect.addEventListener('change', (e) => {
      appState.filters.location = e.target.value;
      updateCascadingBuildingAndRoom();
      applyFiltersAndRender();
    });
  }

  // Building Cascading Dropdown
  const bldgSelect = document.getElementById('filter-building');
  if (bldgSelect) {
    bldgSelect.addEventListener('change', (e) => {
      appState.filters.building = e.target.value;
      updateCascadingBuildingAndRoom();
      applyFiltersAndRender();
    });
  }

  // Room Cascading Dropdown
  const roomSelect = document.getElementById('filter-room');
  if (roomSelect) {
    roomSelect.addEventListener('change', (e) => {
      appState.filters.room = e.target.value;
      applyFiltersAndRender();
    });
  }

  // Checkbox Search Inputs inside Accordions (Traits, Subject, Course)
  wireCheckboxSearch('filter-traits-search', 'list-traits');
  wireCheckboxSearch('filter-subject-search', 'list-subject');
  wireCheckboxSearch('filter-course-search', 'list-course');

  // Accordion Header Toggles
  document.querySelectorAll('.accordion-header').forEach(header => {
    header.addEventListener('click', () => {
      const item = header.closest('.accordion-item');
      item.classList.toggle('open');
    });
  });

  // Reset All Filters Button
  document.getElementById('btn-reset-filters')?.addEventListener('click', resetAllFilters);
  document.getElementById('btn-clear-active-pills')?.addEventListener('click', resetAllFilters);

  // Day View Button Group
  const dayGroup = document.getElementById('btn-group-days');
  if (dayGroup) {
    dayGroup.querySelectorAll('.btn-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        dayGroup.querySelectorAll('.btn-toggle').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const raw = btn.getAttribute('data-days');
        if (raw === 'weekend') {
          appState.daySpan = 'weekend';
        } else if (raw === 'active') {
          appState.daySpan = 'active';
        } else {
          appState.daySpan = parseInt(raw, 10);
        }
        dispatchViewRender();
      });
    });
  }

  // Layout View Button Group (Day View vs Room View vs GE Coverage)
  const layoutGroup = document.getElementById('btn-group-layout');
  if (layoutGroup) {
    layoutGroup.querySelectorAll('.btn-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        layoutGroup.querySelectorAll('.btn-toggle').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        appState.layoutMode = btn.getAttribute('data-layout');
        applyFiltersAndRender();
      });
    });
  }

  // Sidebar Toggle
  document.getElementById('btn-toggle-sidebar')?.addEventListener('click', () => {
    document.getElementById('app-sidebar')?.classList.toggle('collapsed');
  });

  // Modal Close Buttons & Backdrop Clicks
  document.getElementById('btn-close-modal')?.addEventListener('click', closeSectionModal);
  document.getElementById('modal-details-backdrop')?.addEventListener('click', (e) => {
    if (e.target.id === 'modal-details-backdrop') closeSectionModal();
  });
  document.getElementById('modal-conflict-backdrop')?.addEventListener('click', (e) => {
    if (e.target.id === 'modal-conflict-backdrop') closeConflictCenterModal();
  });

  // Keyboard Shortcuts (Esc to close modals)
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeSectionModal();
      closeConflictCenterModal();
      hidePopover();
    }
  });
}

function wireCheckboxSearch(inputId, listContainerId) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener('input', (e) => {
    const q = e.target.value.toLowerCase().trim();
    const container = document.getElementById(listContainerId);
    if (container) {
      const labels = container.querySelectorAll('.checkbox-label');
      labels.forEach(label => {
        const text = label.textContent.toLowerCase();
        label.style.display = (!q || text.includes(q)) ? 'flex' : 'none';
      });
      // Invalidate render cache so future populateFilterControls calls reflect the query
      container._renderKey = null;
    }
  });
}

function resetAllFilters() {
  appState.filters.search = '';
  appState.filters.location = '';
  appState.filters.building = '';
  appState.filters.room = '';
  appState.filters.dayEvening.clear();
  appState.filters.geArea.clear();
  appState.filters.gePattern.clear();
  appState.filters.acc.clear();
  appState.filters.subject.clear();
  appState.filters.course.clear();
  appState.filters.modality.clear();
  appState.filters.session.clear();
  appState.filters.traits.clear();
  appState.fasttrackActiveTab = 'all';
  appState.efficiencyActiveTab = 'all';

  const searchInput = document.getElementById('filter-search');
  if (searchInput) searchInput.value = '';

  const locSelect = document.getElementById('filter-location');
  if (locSelect) locSelect.value = '';

  const bldgSelect = document.getElementById('filter-building');
  if (bldgSelect) bldgSelect.value = '';

  const roomSelect = document.getElementById('filter-room');
  if (roomSelect) roomSelect.value = '';

  const subjectSearch = document.getElementById('filter-subject-search');
  if (subjectSearch) subjectSearch.value = '';

  const courseSearch = document.getElementById('filter-course-search');
  if (courseSearch) courseSearch.value = '';

  const traitsSearch = document.getElementById('filter-traits-search');
  if (traitsSearch) traitsSearch.value = '';

  // Invalidate render keys so all lists re-render freshly
  [
    'list-ge-area', 'list-ge-pattern', 'list-acc', 'list-subject',
    'list-course', 'list-modality', 'list-session', 'list-day-evening', 'list-traits'
  ].forEach(id => {
    const el = document.getElementById(id);
    if (el) el._renderKey = null;
  });

  populateFilterControls();
  applyFiltersAndRender();
}

// --- 14. BOOTSTRAP APPLICATION ---
window.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  processRawDataset([]);
});
