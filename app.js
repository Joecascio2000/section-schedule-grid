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
const GE_STANDARD_AREAS = {
  'CALGETC': [
    { code: '1A', name: 'English Composition', desc: 'Written Communication' },
    { code: '1B', name: 'Critical Thinking and Composition', desc: 'Critical Thinking' },
    { code: '1C', name: 'Oral Communication', desc: 'Public Speaking' },
    { code: '2', name: 'Mathematical Concepts & Quantitative Reasoning', desc: 'Mathematics / Statistics' },
    { code: '3A', name: 'Arts', desc: 'Arts, Cinema, Dance, Music, Theater' },
    { code: '3B', name: 'Humanities', desc: 'Literature, Philosophy, World Languages' },
    { code: '4', name: 'Social and Behavioral Sciences', desc: 'History, Psychology, Sociology, Poli Sci' },
    { code: '5A', name: 'Physical Sciences', desc: 'Chemistry, Physics, Astronomy, Geology' },
    { code: '5B', name: 'Biological Sciences', desc: 'Biology, Anatomy, Microbiology, Ecology' },
    { code: '5C', name: 'Laboratory Activity', desc: 'Physical/Bio Science Laboratory Component' }
  ],
  'CSUGE': [
    { code: 'A1', name: 'Oral Communication', desc: 'Public Speaking' },
    { code: 'A2', name: 'Written Communication', desc: 'English Composition' },
    { code: 'A3', name: 'Critical Thinking', desc: 'Logic & Reasoning' },
    { code: 'B1', name: 'Physical Science', desc: 'Physical Sciences' },
    { code: 'B2', name: 'Life Science', desc: 'Biological Sciences' },
    { code: 'B3', name: 'Laboratory Activity', desc: 'Lab' },
    { code: 'B4', name: 'Mathematics/Quantitative Reasoning', desc: 'Math & Quantitative' },
    { code: 'C1', name: 'Arts', desc: 'Arts & Cinema' },
    { code: 'C2', name: 'Humanities', desc: 'Literature, Philosophy, Languages' },
    { code: 'D', name: 'Social Sciences', desc: 'Social & Behavioral Sciences' },
    { code: 'E', name: 'Lifelong Learning', desc: 'Self-Development' },
    { code: 'F', name: 'Ethnic Studies', desc: 'Ethnic Studies' }
  ],
  'CHAFFEY': [
    { code: '1A', name: 'English Composition', desc: 'Language & Rationality' },
    { code: '1C', name: 'Oral Communication', desc: 'Oral Communication' },
    { code: '2A', name: 'Mathematics', desc: 'Mathematics Competency' },
    { code: '3A', name: 'Arts', desc: 'Arts & Humanities' },
    { code: '3B', name: 'Humanities', desc: 'Humanities' },
    { code: '4', name: 'Social and Behavioral Sciences', desc: 'Social Sciences' },
    { code: '5A', name: 'Natural Sciences', desc: 'Physical & Life Sciences' }
  ],
  'IGETC': [
    { code: '1A', name: 'English Composition', desc: 'English Composition' },
    { code: '1B', name: 'Critical Thinking', desc: 'Critical Thinking' },
    { code: '1C', name: 'Oral Communication', desc: 'Oral Communication' },
    { code: '2A', name: 'Mathematical Concepts', desc: 'Math Concepts' },
    { code: '3A', name: 'Arts', desc: 'Arts' },
    { code: '3B', name: 'Humanities', desc: 'Humanities' },
    { code: '4', name: 'Social and Behavioral Sciences', desc: 'Social Sciences' },
    { code: '5A', name: 'Physical Sciences', desc: 'Physical Sciences' },
    { code: '5B', name: 'Biological Sciences', desc: 'Biological Sciences' }
  ]
};

// Built-in Dataset matching IR - Schedule Sections Building-Room Grids (JC) - Copy.xlsx
// Enriched with stacked classes and multi-campus GE sections for full feature demonstration
const BUILT_IN_DATASET = [
  {
    "Section Unique ID": "210928",
    "Term": "2026/FA",
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48898",
    "Section Number": "48898",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48898",
    "Section Number": "48898",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-45",
    "Section Name": "FASHD-45-48900",
    "Section Number": "48900",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48924",
    "Section Number": "48924",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-40",
    "Section Name": "FASHD-40-48924",
    "Section Number": "48924",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-42",
    "Section Name": "FASHD-42-48933",
    "Section Number": "48933",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-42",
    "Section Name": "FASHD-42-48933",
    "Section Number": "48933",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-470",
    "Section Name": "FASHD-470-48964",
    "Section Number": "48964",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-61",
    "Section Name": "FASHD-61-48977",
    "Section Number": "48977",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-61",
    "Section Name": "FASHD-61-48977",
    "Section Number": "48977",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Merchandising",
    "Course": "FASHM-60",
    "Section Name": "FASHM-60-49027",
    "Section Number": "49027",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-72",
    "Section Name": "FASHD-72-51708",
    "Section Number": "51708",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-72",
    "Section Name": "FASHD-72-51708",
    "Section Number": "51708",
    "Location": "Chino",
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
    "ACC": "Arts, Communication & Design-Chino",
    "Subject": "Fashion Design",
    "Course": "FASHD-72",
    "Section Name": "FASHD-72-51708",
    "Section Number": "51708",
    "Location": "Chino",
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
    "ACC": "Arts, Communication and Design",
    "Subject": "Communication Studies",
    "Course": "COMM-C1000",
    "Section Name": "COMM-C1000-47109",
    "Section Number": "47109",
    "Location": "Chino",
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
    "ACC": "Arts, Communication and Design",
    "Subject": "Art",
    "Course": "ART-10",
    "Section Name": "ART-10-47200",
    "Section Number": "47200",
    "Location": "Chino",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Zero Textbook Cost ($0)",
    "General Ed Pattern": "CHAFFEY,CALGETC,CSUGE",
    "General Ed Area": "CalGETC,3A,C1",
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
    "ACC": "Arts, Humanities & Social Sciences",
    "Subject": "English",
    "Course": "ENGL-1A",
    "Section Name": "ENGL-1A-47300",
    "Section Number": "47300",
    "Location": "Chino",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "CHAFFEY,CALGETC",
    "General Ed Area": "CalGETC,1A",
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
    "ACC": "Mathematics and Science",
    "Subject": "Mathematics",
    "Course": "MATH-25",
    "Section Name": "MATH-25-50111",
    "Section Number": "50111",
    "Location": "Rancho",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "Not Basic Skills,In-Person,Transferable to UC/CSU,Zero Textbook Cost ($0)",
    "General Ed Pattern": "CHAFFEY,CALGETC,CSUGE",
    "General Ed Area": "CalGETC,2,B4",
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
    "ACC": "Arts, Humanities & Social Sciences",
    "Subject": "English",
    "Course": "ENGL-1A",
    "Section Name": "ENGL-1A-50222",
    "Section Number": "50222",
    "Location": "Rancho",
    "Day or Evening": "D",
    "Calendar Session": "Full-Term",
    "Modality": "In-Person",
    "Number of Weeks": 18,
    "Section Traits": "In-Person,Low-Cost Textbook (<$50)",
    "General Ed Pattern": "CALGETC",
    "General Ed Area": "CalGETC,1A",
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
    "ACC": "Social and Behavioral Sciences",
    "Subject": "Psychology",
    "Course": "PSYCH-1",
    "Section Name": "PSYCH-1-51204",
    "Section Number": "51204",
    "Location": "Chino",
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
    "ACC": "Arts, Communication and Design",
    "Subject": "Art",
    "Course": "ART-10",
    "Section Name": "ART-10-51402",
    "Section Number": "51402",
    "Location": "Chino",
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
    "ACC": "STEM & Health Sciences",
    "Subject": "Biology",
    "Course": "BIOL-10",
    "Section Name": "BIOL-10-51610",
    "Section Number": "51610",
    "Location": "Rancho",
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
    "ACC": "STEM & Health Sciences",
    "Subject": "Dental Assisting",
    "Course": "DENTAL-100",
    "Section Name": "DENTAL-100-52010",
    "Section Number": "52010",
    "Location": "Chino",
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
    "ACC": "STEM & Health Sciences",
    "Subject": "Dental Assisting",
    "Course": "DENTAL-102",
    "Section Name": "DENTAL-102-52012",
    "Section Number": "52012",
    "Location": "Chino",
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
    "ACC": "STEM & Health Sciences",
    "Subject": "Radiologic Technology",
    "Course": "RADTEC-10",
    "Section Name": "RADTEC-10-52020",
    "Section Number": "52020",
    "Location": "Chino",
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
  scheduleBlocks: [],      // Flattened list of individual day meeting occurrences
  filteredBlocks: [],      // Currently displayed meeting blocks after filtering
  conflicts: [],           // List of detected real room conflicts (different instructors)
  stackedPairs: [],        // List of detected stacked classes (same instructor or known cohort program)
  isConflictMode: false,   // "View Room Conflicts" toggle state
  conflictsOnly: true,     // Default to Isolate Conflicts Only when conflict mode is active!
  daySpan: 'active',       // Default to Active Days: only days with scheduled classes in view
  layoutMode: 'room',      // Default to Room Matrix rather than Day View
  geSubLayout: 'room',     // Default GE Coverage grid view to Room Matrix
  eveningSubLayout: 'room',// Default Evening Coverage grid view to Room Matrix
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
      const gePatternList = parseCommaSeparatedList(row['General Ed Pattern']);
      const geAreaList = parseCommaSeparatedList(row['General Ed Area']);

      sectionMap.set(groupKey, {
        sectionUniqueId: sectionUniqueId,
        sectionName: sectionName,
        sectionNumber: String(row['Section Number'] || '').trim(),
        term: String(row['Term'] || '').trim(),
        acc: String(row['ACC'] || '').trim(),
        subject: String(row['Subject'] || '').trim(),
        course: String(row['Course'] || '').trim(),
        location: String(row['Location'] || '').trim(),
        dayOrEvening: String(row['Day or Evening'] || '').trim(),
        calendarSession: String(row['Calendar Session'] || '').trim(),
        modality: String(row['Modality'] || '').trim(),
        numberOfWeeks: row['Number of Weeks'] || '',
        facultyName: String(row['Faculty Name'] || '').trim(),
        sectionTraits: traitsList,
        sectionTraitsRaw: row['Section Traits'] || '',
        generalEdPattern: gePatternList,
        generalEdPatternRaw: row['General Ed Pattern'] || '',
        generalEdArea: geAreaList,
        generalEdAreaRaw: row['General Ed Area'] || '',
        scheduleBlocks: []
      });
    }

    const parentSection = sectionMap.get(groupKey);

    const startMinutes = parseTimeToMinutes(row['Schedule - Start Time']);
    const endMinutes = parseTimeToMinutes(row['Schedule - End Time']);
    const daysArray = parseScheduleDays(row['Schedule - Days']);

    const meetingRow = {
      rowId: `row_${rowIndex}`,
      instructionType: String(row['Schedule - Instruction Type'] || 'Instruction').trim(),
      building: String(row['Schedule - Building'] || '').trim(),
      buildingDesc: String(row['Schedule - Building Description'] || '').trim(),
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
          startMinutes: startMinutes,
          endMinutes: endMinutes,
          duration: Math.max(endMinutes - startMinutes, 20),
          instructionType: meetingRow.instructionType,
          building: meetingRow.building,
          buildingDesc: meetingRow.buildingDesc,
          room: meetingRow.room,
          startTimeStr: meetingRow.startTimeStr,
          endTimeStr: meetingRow.endTimeStr,
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
  appState.scheduleBlocks = allMeetingBlocks;

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
 * Evaluates Room Scheduling Conflicts & Stacked Classes.
 * - STACKED CLASSES: Overlapping classes with the SAME instructor OR classes in the same cohort
 *   program (AMT, NURADN, DENTAL, RADTEC) with matching subjects.
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

        const isStacked = isSameInstructor || isSameProgramMultiStack;

        if (isStacked) {
          // STACKED CLASS: Same room, day, time, and either same faculty OR same program in (AMT, NURADN, DENTAL, RADTEC)!
          a.isStacked = true;
          b.isStacked = true;
          if (!a.stackedWith.includes(b)) a.stackedWith.push(b);
          if (!b.stackedWith.includes(a)) b.stackedWith.push(a);

          const stackReason = isSameInstructor
            ? `Same Instructor: ${a.parentSection.facultyName}`
            : `Cohort Program Stack: ${stackProgA} (Multi-Instructor)`;

          stackedPairs.push({
            day: a.day,
            building: a.building,
            room: a.room,
            faculty: a.parentSection.facultyName,
            facultyB: b.parentSection.facultyName,
            program: stackProgA || '',
            reason: stackReason,
            isMultiInstructor: isSameProgramMultiStack && !isSameInstructor,
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

function populateFilterControls() {
  const locations = new Set();
  const buildings = new Set();
  const rooms = new Set();
  const accs = new Set();
  const subjects = new Set();
  const courses = new Set();
  const modalities = new Set();
  const sessions = new Set();
  const dayEvenings = new Set();
  const gePatterns = new Set();
  const geAreas = new Set();
  const traits = new Set();

  appState.sections.forEach(sec => {
    if (sec.location) locations.add(sec.location);
    if (sec.acc) accs.add(sec.acc);
    if (sec.subject) subjects.add(sec.subject);
    if (sec.course) courses.add(sec.course);
    if (sec.modality) modalities.add(sec.modality);
    if (sec.calendarSession) sessions.add(sec.calendarSession);
    if (sec.dayOrEvening) dayEvenings.add(sec.dayOrEvening);

    sec.generalEdPattern.forEach(p => gePatterns.add(p));
    sec.generalEdArea.forEach(a => geAreas.add(a));
    sec.sectionTraits.forEach(t => traits.add(t));

    sec.scheduleBlocks.forEach(b => {
      if (b.building) buildings.add(b.building);
      if (b.room) rooms.add(b.room);
    });
  });

  const locSelect = document.getElementById('filter-location');
  if (locSelect) {
    const currentLoc = locSelect.value;
    locSelect.innerHTML = '<option value="">All Locations</option>';
    Array.from(locations).sort().forEach(loc => {
      const opt = document.createElement('option');
      opt.value = loc;
      opt.textContent = loc;
      if (loc === currentLoc) opt.selected = true;
      locSelect.appendChild(opt);
    });
    document.getElementById('count-location').textContent = `(${locations.size})`;
  }

  updateCascadingBuildingAndRoom();

  renderCheckboxList('list-ge-area', Array.from(geAreas).sort(), appState.filters.geArea, 'badge-ge-area');
  renderCheckboxList('list-ge-pattern', Array.from(gePatterns).sort(), appState.filters.gePattern, 'badge-ge-pattern');
  renderCheckboxList('list-acc', Array.from(accs).sort(), appState.filters.acc, 'badge-acc');
  renderCheckboxList('list-subject', Array.from(subjects).sort(), appState.filters.subject, 'badge-subject');
  renderCheckboxList('list-course', Array.from(courses).sort(), appState.filters.course, 'badge-course');
  renderCheckboxList('list-modality', Array.from(modalities).sort(), appState.filters.modality, 'badge-modality');
  renderCheckboxList('list-session', Array.from(sessions).sort(), appState.filters.session, 'badge-session');
  renderCheckboxList('list-day-evening', Array.from(dayEvenings).sort(), appState.filters.dayEvening, 'badge-day-evening', formatDayEveningLabel);
  renderCheckboxList('list-traits', Array.from(traits).sort(), appState.filters.traits, 'badge-traits');
}

function updateCascadingBuildingAndRoom() {
  const selectedLocation = appState.filters.location;
  const selectedBuilding = appState.filters.building;

  const validBuildings = new Set();
  const validRooms = new Set();

  appState.sections.forEach(sec => {
    if (!selectedLocation || sec.location.toLowerCase() === selectedLocation.toLowerCase()) {
      sec.scheduleBlocks.forEach(b => {
        if (b.building) validBuildings.add(b.building);
        if (!selectedBuilding || b.building.toLowerCase() === selectedBuilding.toLowerCase()) {
          if (b.room) validRooms.add(b.room);
        }
      });
    }
  });

  const bldgSelect = document.getElementById('filter-building');
  if (bldgSelect) {
    const prevBldg = bldgSelect.value;
    bldgSelect.innerHTML = '<option value="">All Buildings</option>';
    Array.from(validBuildings).sort().forEach(bldg => {
      const opt = document.createElement('option');
      opt.value = bldg;
      opt.textContent = bldg;
      if (bldg === prevBldg && validBuildings.has(prevBldg)) opt.selected = true;
      bldgSelect.appendChild(opt);
    });
    document.getElementById('count-building').textContent = `(${validBuildings.size})`;
    if (!validBuildings.has(appState.filters.building)) {
      appState.filters.building = '';
    }
  }

  const roomSelect = document.getElementById('filter-room');
  if (roomSelect) {
    const prevRoom = roomSelect.value;
    roomSelect.innerHTML = '<option value="">All Rooms</option>';
    Array.from(validRooms).sort().forEach(room => {
      const opt = document.createElement('option');
      opt.value = room;
      opt.textContent = room;
      if (room === prevRoom && validRooms.has(prevRoom)) opt.selected = true;
      roomSelect.appendChild(opt);
    });
    document.getElementById('count-room').textContent = `(${validRooms.size})`;
    if (!validRooms.has(appState.filters.room)) {
      appState.filters.room = '';
    }
  }
}

function renderCheckboxList(containerId, items, selectedSet, badgeId, labelFormatter = null) {
  const container = document.getElementById(containerId);
  if (!container) return;
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
    span.textContent = labelFormatter ? labelFormatter(item) : item;

    label.appendChild(checkbox);
    label.appendChild(span);
    container.appendChild(label);
  });

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
  const {
    search, location, building, room,
    dayEvening, geArea, gePattern, acc, subject, course,
    modality, session, traits
  } = appState.filters;

  const searchLower = search.toLowerCase().trim();

  // Filter Parent Sections
  const matchedSections = appState.sections.filter(sec => {
    if (searchLower) {
      const combined = `${sec.sectionName} ${sec.course} ${sec.subject} ${sec.facultyName} ${sec.location} ${sec.scheduleBlocks.map(b => `${b.building} ${b.room}`).join(' ')}`.toLowerCase();
      if (!combined.includes(searchLower)) return false;
    }

    if (location && sec.location.toLowerCase() !== location.toLowerCase()) return false;

    if (dayEvening.size > 0 && !dayEvening.has(sec.dayOrEvening)) return false;
    if (acc.size > 0 && !acc.has(sec.acc)) return false;
    if (subject.size > 0 && !subject.has(sec.subject)) return false;
    if (course.size > 0 && !course.has(sec.course)) return false;
    if (modality.size > 0 && !modality.has(sec.modality)) return false;
    if (session.size > 0 && !session.has(sec.calendarSession)) return false;

    if (geArea.size > 0) {
      const hasGeArea = Array.from(geArea).some(a => sec.generalEdArea.some(sa => sa.toLowerCase() === a.toLowerCase()));
      if (!hasGeArea) return false;
    }

    if (gePattern.size > 0) {
      const hasGePat = Array.from(gePattern).some(p => sec.generalEdPattern.some(sp => sp.toLowerCase() === p.toLowerCase()));
      if (!hasGePat) return false;
    }

    if (traits.size > 0) {
      const hasTrait = Array.from(traits).some(t => sec.sectionTraits.includes(t));
      if (!hasTrait) return false;
    }

    if (building || room) {
      const hasMatchingBlock = sec.scheduleBlocks.some(b => {
        if (building && b.building.toLowerCase() !== building.toLowerCase()) return false;
        if (room && b.room.toLowerCase() !== room.toLowerCase()) return false;
        return true;
      });
      if (!hasMatchingBlock) return false;
    }

    return true;
  });

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

  // If in GE Coverage view, show only sections with General Ed attributes
  if (appState.layoutMode === 'ge') {
    filteredBlocks = filteredBlocks.filter(b => b.parentSection.generalEdPattern.length > 0 || b.parentSection.generalEdArea.length > 0);
  }

  // If in Evening Coverage view, show only Evening/Weekend classes
  if (appState.layoutMode === 'evening') {
    filteredBlocks = filteredBlocks.filter(isEveningBlock);
  }

  // REQUIREMENT 1: If Conflict Mode is active and "Isolate Conflicts Only" is on, isolate conflicts!
  if (appState.isConflictMode && appState.conflictsOnly) {
    filteredBlocks = filteredBlocks.filter(b => b.isConflict);
  }

  appState.filteredBlocks = filteredBlocks;

  document.getElementById('stat-total-sections').textContent = appState.sections.length;
  document.getElementById('stat-total-meetings').textContent = appState.scheduleBlocks.length;
  document.getElementById('status-sections-count').textContent = matchedSections.length;
  document.getElementById('status-meetings-count').textContent = filteredBlocks.length;

  renderActiveFilterPills();
  dispatchViewRender();
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
    addPill('Building', appState.filters.building, () => {
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
    addPill('GE Area', area, () => {
      appState.filters.geArea.delete(area);
      populateFilterControls();
      applyFiltersAndRender();
    });
  });

  appState.filters.gePattern.forEach(pat => {
    addPill('GE Pattern', pat, () => {
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
 * - 'room': Stacked Room Matrix View (vertically stacked grids by Building/Room)
 * - 'ge': General Education Coverage Audit Dashboard & GE Schedule (swappable to Day or Room Matrix)
 * - 'evening': Evening & Weekend Coverage Audit Dashboard & Schedule (swappable to Day or Room Matrix)
 */
function dispatchViewRender() {
  const welcomeHero = document.getElementById('welcome-upload-hero');
  const standardGrid = document.getElementById('schedule-grid');
  const stackedContainer = document.getElementById('room-matrix-stacked');
  const geContainer = document.getElementById('ge-coverage-container');
  const eveningContainer = document.getElementById('evening-coverage-container');

  // If no sections are loaded yet, display prominent Step 1 upload hero
  if (appState.sections.length === 0) {
    if (welcomeHero) welcomeHero.style.display = 'flex';
    if (standardGrid) standardGrid.style.display = 'none';
    if (stackedContainer) stackedContainer.style.display = 'none';
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    return;
  }

  // Hide welcome hero once data is loaded
  if (welcomeHero) welcomeHero.style.display = 'none';

  if (appState.layoutMode === 'room') {
    if (standardGrid) standardGrid.style.display = 'none';
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (stackedContainer) {
      stackedContainer.style.display = 'flex';
      renderStackedRoomMatrix();
    }
  } else if (appState.layoutMode === 'ge') {
    if (eveningContainer) eveningContainer.style.display = 'none';
    if (geContainer) {
      geContainer.style.display = 'flex';
      renderGeCoverageDashboard();
    }
    // Swap grid below coverage audit to either Day View or Room Matrix
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
    if (eveningContainer) {
      eveningContainer.style.display = 'flex';
      renderEveningCoverageDashboard();
    }
    // Swap grid below coverage audit to either Day View or Room Matrix
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
  } else {
    // Default 'day' view
    if (geContainer) geContainer.style.display = 'none';
    if (eveningContainer) eveningContainer.style.display = 'none';
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
          <div style="font-size:11px; color:var(--text-muted);">${campusLoc} Campus</div>
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

  if (s.includes('1A') || s.includes('A2')) return 'ge-color-1a';
  if (s.includes('1B') || s.includes('A3')) return 'ge-color-1b';
  if (s.includes('1C') || s.includes('A1')) return 'ge-color-1c';
  if (s.includes('2A') || s === '2' || s.includes('B4') || s.includes('AREA 2')) return 'ge-color-2';
  if (s.includes('3A') || s.includes('C1')) return 'ge-color-3a';
  if (s.includes('3B') || s.includes('C2')) return 'ge-color-3b';
  if (s.includes('4') || s.includes('AREA 4') || s === 'D' || s.startsWith('D')) return 'ge-color-4';
  if (s.includes('5A') || s.includes('B1')) return 'ge-color-5a';
  if (s.includes('5B') || s.includes('B2')) return 'ge-color-5b';
  if (s.includes('5C') || s.includes('B3')) return 'ge-color-5c';
  if (s.includes('6') || s.includes('AREA 6')) return 'ge-color-6';
  if (s === 'E' || s.includes('AREA E')) return 'ge-color-e';
  if (s === 'F' || s.includes('AREA F')) return 'ge-color-f';
  if (s.includes('1')) return 'ge-color-1a';
  if (s.includes('2')) return 'ge-color-2';
  if (s.includes('3')) return 'ge-color-3a';
  if (s.includes('5')) return 'ge-color-5a';
  return 'ge-color-other';
}

/**
 * REQUIREMENT 3: GENERAL EDUCATION (GE) COVERAGE AUDIT VIEW
 * Audits GE Pattern & Area offerings for the selected campus,
 * displaying coverage status, missing areas, interactive course chips,
 * and swappable calendar grid (Day View vs Room Matrix).
 */
function renderGeCoverageDashboard() {
  const container = document.getElementById('ge-coverage-container');
  if (!container) return;
  container.innerHTML = '';

  const activePattern = appState.geActivePattern || 'CALGETC';
  const campusLocation = appState.filters.location || 'All Campuses';
  const standardAreas = GE_STANDARD_AREAS[activePattern] || GE_STANDARD_AREAS['CALGETC'];

  // Check which sections satisfy each area for this pattern
  const currentSections = appState.sections.filter(sec => {
    if (appState.filters.location && sec.location.toLowerCase() !== appState.filters.location.toLowerCase()) {
      return false;
    }
    return sec.generalEdPattern.some(p => p.toUpperCase().includes(activePattern.toUpperCase()));
  });

  const areaCoverageMap = new Map();
  standardAreas.forEach(a => {
    areaCoverageMap.set(a.code, {
      meta: a,
      offerings: []
    });
  });

  currentSections.forEach(sec => {
    sec.generalEdArea.forEach(areaCode => {
      // Find matching standard area
      const matched = standardAreas.find(sa => sa.code.toUpperCase() === areaCode.toUpperCase() || areaCode.toUpperCase().includes(sa.code.toUpperCase()));
      if (matched && areaCoverageMap.has(matched.code)) {
        const item = areaCoverageMap.get(matched.code);
        if (!item.offerings.some(o => o.sectionUniqueId === sec.sectionUniqueId)) {
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

  const coveragePercent = Math.round((coveredCount / standardAreas.length) * 100);

  // 1. Header Row (Title, Campus, Pattern Tabs, and Sub-Layout Switcher)
  const headerRow = document.createElement('div');
  headerRow.className = 'ge-coverage-header';
  headerRow.innerHTML = `
    <div class="ge-coverage-title">
      <h3>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#990000" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/></svg>
        General Education Coverage Audit: <span>${campusLocation}</span>
      </h3>
      <p>Identify active offerings and unscheduled GE requirements for <strong>${activePattern}</strong>.</p>
    </div>
    <div class="ge-header-controls">
      <div style="display:flex; align-items:center; gap:6px;">
        <span style="font-size:11px; font-weight:700; color:var(--text-muted);">PATTERN:</span>
        <div class="ge-pattern-tabs" id="ge-pattern-tabs">
          <button type="button" class="ge-pattern-tab ${activePattern === 'CALGETC' ? 'active' : ''}" data-pattern="CALGETC">CalGETC</button>
          <button type="button" class="ge-pattern-tab ${activePattern === 'CHAFFEY' ? 'active' : ''}" data-pattern="CHAFFEY">CHAFFEY Local</button>
          <button type="button" class="ge-pattern-tab ${activePattern === 'CSUGE' ? 'active' : ''}" data-pattern="CSUGE">CSUGE</button>
          <button type="button" class="ge-pattern-tab ${activePattern === 'IGETC' ? 'active' : ''}" data-pattern="IGETC">IGETC</button>
        </div>
      </div>
      <div class="sub-layout-switcher" id="ge-sub-layout-switcher">
        <span class="sub-layout-label">Grid View:</span>
        <div class="btn-sub-toggle-group">
          <button type="button" class="btn-sub-layout-toggle ${appState.geSubLayout === 'day' ? 'active' : ''}" data-sublayout="day" title="View as Standard Day Grid">Day View</button>
          <button type="button" class="btn-sub-layout-toggle ${appState.geSubLayout === 'room' ? 'active' : ''}" data-sublayout="room" title="View as Stacked Room Grids">Room Matrix</button>
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

  // Sub-layout switcher listener (Day View vs Room Matrix)
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

  areaCoverageMap.forEach((data, areaCode) => {
    const isCovered = data.offerings.length > 0;
    if (appState.geShowMissingOnly && isCovered) return;

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
          ${isCovered ? `✓ ${data.offerings.length} Scheduled` : `⚠ Missing at ${campusLocation}`}
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

  const campusLocation = appState.filters.location || 'All Campuses';
  const eveningSections = appState.sections.filter(sec => {
    if (appState.filters.location && sec.location.toLowerCase() !== appState.filters.location.toLowerCase()) {
      return false;
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
    sec.scheduleBlocks.forEach(b => {
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
        Evening & Weekend Coverage Audit: <span>${campusLocation}</span>
      </h3>
      <p>Auditing classes scheduled Friday–Sunday, starting at/after 4:30 PM, or designated Evening.</p>
    </div>
    <div class="evening-header-controls">
      <div class="sub-layout-switcher" id="evening-sub-layout-switcher">
        <span class="sub-layout-label">Grid View:</span>
        <div class="btn-sub-toggle-group">
          <button type="button" class="btn-sub-layout-toggle ${appState.eveningSubLayout === 'day' ? 'active' : ''}" data-sublayout="day" title="View as Standard Day Grid">Day View</button>
          <button type="button" class="btn-sub-layout-toggle ${appState.eveningSubLayout === 'room' ? 'active' : ''}" data-sublayout="room" title="View as Stacked Room Grids">Room Matrix</button>
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
          const times = sec.scheduleBlocks.map(b => `${DAY_MAP[b.days[0]]?.short || ''} ${b.startTimeStr}`).join(', ');
          return `<span class="ge-course-chip" data-sec-id="${sec.sectionUniqueId}" title="${sec.sectionName} (${times})">${sec.course} (${sec.facultyName || 'Staff'})</span>`;
        }).join('')
      : `<span style="font-size:12px; color:var(--text-muted); font-style:italic;">No classes currently scheduled in this band at ${campusLocation}.</span>`;

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
  if (sec.generalEdArea.length > 0) {
    const areaFirst = sec.generalEdArea[0];
    const colorClass = getGeAreaBadgeClass(areaFirst);
    geTagHtml = `<span class="tile-ge-tag ${colorClass}">Area ${areaFirst}</span>`;
  }

  // Evening indicator tag
  let eveningTagHtml = '';
  if (isEveningBlock(block)) {
    let eveLabel = '🌙 Evening';
    if (block.day === 6 || block.day === 7) eveLabel = '📅 Weekend';
    else if (block.day === 5) eveLabel = '⚡ Friday';
    eveningTagHtml = `<span class="tile-ge-tag ge-color-1c" style="background:#4338ca;">${eveLabel}</span>`;
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
    const stackNote = isSameFac
      ? `Same Instructor: ${facA}`
      : `Cohort Stack: ${prog || 'Specialized Program'} (Instructors: ${facA} & ${facB})`;

    stackedAlertHtml = `
      <div style="margin-top:8px; padding:6px 8px; background-color:var(--stacked-bg); border:1px solid var(--stacked-border); border-radius:6px; color:var(--stacked-text); font-size:11px; font-weight:600;">
        🔗 <strong>STACKED CLASS</strong>: Cross-scheduled simultaneously with <strong>${other.parentSection.course}</strong> (${stackNote}) in ${block.building} ${block.room}.
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

  const meetingTypeStr = block ? block.instructionType : (sec.scheduleBlocks?.map(b => b.instructionType).join(', ') || 'Class Meeting');
  const timeDayStr = block && block.day ? `${DAY_MAP[block.day]?.name || ''}: ${block.startTimeStr} – ${block.endTimeStr}` : (sec.scheduleBlocks?.map(b => `${(b.days || []).map(d => DAY_MAP[d]?.short || d).join('/')} ${b.startTimeStr}-${b.endTimeStr}`).join('; ') || 'TBA');
  const locationStr = block && block.building ? `${sec.location} – ${block.building} ${block.room}` : `${sec.location}`;

  popoverEl.innerHTML = `
    <div class="popover-header">
      <div class="popover-title-row">
        <span class="popover-course">${sec.course}</span>
        <span class="popover-id-pill">ID: ${sec.sectionUniqueId}</span>
      </div>
      <div class="popover-sec-name">${sec.sectionName}</div>
    </div>
    <div class="popover-grid">
      <span class="popover-field-label">Subject:</span>
      <span class="popover-field-value">${sec.subject}</span>

      <span class="popover-field-label">Faculty:</span>
      <span class="popover-field-value">${sec.facultyName || 'Staff'}</span>

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
      <span class="popover-field-value">${sec.generalEdPatternRaw || 'None'}</span>

      <span class="popover-field-label">GE Area:</span>
      <span class="popover-field-value">${sec.generalEdAreaRaw || 'None'}</span>
    </div>

    ${traitsHtml ? `
      <div class="popover-field-label" style="margin-top:6px;">Section Traits:</div>
      <div class="popover-traits-list">${traitsHtml}</div>
    ` : ''}

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
  modalSecName.textContent = `${section.course} – ${section.sectionName}`;

  const blocksHtml = section.scheduleBlocks.map((blk, idx) => {
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
          Part ${idx + 1} of ${section.scheduleBlocks.length}
        </span>
      </div>
    `;
  }).join('');

  const traitsHtml = section.sectionTraits.map(t => `<span class="trait-tag">${t}</span>`).join('');

  modalBody.innerHTML = `
    <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:16px;">
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">SUBJECT & COURSE</div>
        <div style="font-size:13px; font-weight:600;">${section.subject} (${section.course})</div>
      </div>
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
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">GENERAL ED PATTERN</div>
        <div style="font-size:13px; font-weight:600;">${section.generalEdPatternRaw || 'None'}</div>
      </div>
      <div>
        <div style="font-size:11px; color:var(--text-muted); font-weight:600;">GENERAL ED AREA</div>
        <div style="font-size:13px; font-weight:600;">${section.generalEdAreaRaw || 'None'}</div>
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
    'Section Unique ID', 'Course', 'Section Name', 'Subject', 'Faculty Name',
    'Location', 'Building', 'Room', 'Day', 'Start Time', 'End Time',
    'Instruction Type', 'General Ed Pattern', 'General Ed Area', 'Section Traits',
    'Is Stacked', 'Is Conflict'
  ];

  const rows = appState.filteredBlocks.map(b => {
    const sec = b.parentSection;
    return [
      `"${sec.sectionUniqueId}"`,
      `"${sec.course}"`,
      `"${sec.sectionName}"`,
      `"${sec.subject}"`,
      `"${sec.facultyName}"`,
      `"${sec.location}"`,
      `"${b.building}"`,
      `"${b.room}"`,
      `"${DAY_MAP[b.day].name}"`,
      `"${b.startTimeStr}"`,
      `"${b.endTimeStr}"`,
      `"${b.instructionType}"`,
      `"${sec.generalEdPatternRaw}"`,
      `"${sec.generalEdAreaRaw}"`,
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

  // Traits Search inside Accordion
  const traitsSearch = document.getElementById('filter-traits-search');
  if (traitsSearch) {
    traitsSearch.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase();
      const list = document.querySelectorAll('#list-traits .checkbox-label');
      list.forEach(label => {
        const text = label.textContent.toLowerCase();
        label.style.display = text.includes(q) ? 'flex' : 'none';
      });
    });
  }

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

  // Layout View Button Group (Day View vs Room Matrix vs GE Coverage)
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

  const searchInput = document.getElementById('filter-search');
  if (searchInput) searchInput.value = '';

  const locSelect = document.getElementById('filter-location');
  if (locSelect) locSelect.value = '';

  populateFilterControls();
  applyFiltersAndRender();
}

// --- 14. BOOTSTRAP APPLICATION ---
window.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  processRawDataset([]);
});
