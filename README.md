# Chaffey College | Academic Schedule Grid & Room Conflict Analyzer

An interactive, client-side web application designed for academic deans, department chairs, and scheduling specialists at **Chaffey College** to visualize course offerings, detect room scheduling conflicts, analyze program completability, and evaluate general education coverage across campuses.

🌐 **Live Application**: [https://joecascio2000.github.io/section-schedule-grid/](https://joecascio2000.github.io/section-schedule-grid/)

---

## 🎯 What the Tool Does

The analyzer transforms raw Informer schedule exports (`.xlsx` or `.csv`) into interactive, visual weekly calendar matrices and executive audit dashboards. It runs **100% in the user's browser** (no server upload required, ensuring complete data privacy).

---

## 🧭 Key Views & Dashboards

The application provides seven specialized view layouts accessible from the top toolbar:

### 1. 🏢 Room View (Default)
- **Vertically Stacked Room Grids**: Generates a dedicated weekly calendar grid for each campus building and room currently in view.
- Eliminates visual clutter by letting schedulers inspect room utilization, open timeslots, and schedule density room-by-room.
- Adapts time columns automatically with configurable day spans (Active Days, Weekdays, Weekends, or All 7 Days).

### 2. 📅 Day View
- **Side-by-Side Weekly Calendar**: Visualizes scheduled sections on a single unified calendar grid (Monday–Sunday).
- Uses interval graph partitioning to place overlapping sections into clean, side-by-side sub-columns without overlap collisions.
- Best for department-level or subject-level schedule overviews.

### 3. 🎓 Program Coverage & Completability Audit
- **Degree & Certificate Completability**: Analyzes whether students can complete an academic program of study (e.g., *Business Administration 2.0 AS-T*, *Automotive Technology*) entirely at a specific campus (Chino, Fontana, Rancho) or district-wide.
- **Scheduled vs. Missing Major Courses**: Categorizes all required courses into scheduled offerings vs. unscheduled requirements with clear action items.
- **Cross-Campus Offerings**: When a required course is missing from a campus, the tool indicates where else it is offered (e.g., Rancho Campus or Online).
- **All-Programs Catalog Matrix**: Ranks all 200+ active degrees and certificates by completion percentage to identify under-served programs.
- **GE Breadth for Associate Degrees**: Automatically evaluates Cal-GETC general education breadth requirements for associate degrees (see below).

### 4. 📚 General Education (GE) Coverage Audit
- **Cal-GETC & Chaffey Local Audits**: Audits offerings against California General Education requirements across 12 audited Cal-GETC areas and 7 Chaffey Local areas.
- **Missing Area Detection**: Highlights unscheduled GE areas at a glance so schedulers can add missing breadth requirements.
- **Cal-GETC Area 3A & 3B Mutual Exclusivity**: Enforces the requirement that students complete two distinct courses (one in Arts 3A and one in Humanities 3B). If only one distinct course is scheduled across both areas (e.g., `PHOTO-1`), Area 3B is flagged as deficient with guidance to add a distinct course.
- **American Institutions Tagging**: Incorporates overrides for CSU American Institutions graduation requirements (US-1: `HIST-C1001`, `HIST-C1002`, `HIST-50`, `HIST-51`; US-2/US-3: `POLS-C1000`).

### 5. ⚡ Fast-Track 1 & 2 Pairing Audit
- **8-Week Session Optimization**: Evaluates Fast-Track 1 (first 8 weeks) and Fast-Track 2 (second 8 weeks) offerings.
- Identifies optimized room pairings (sequential FT1 and FT2 sharing the same room and time slot without conflict) and flags unpaired Fast-Track sections that leave rooms vacant for half the semester.

### 6. 🌙 Evening & Weekend Coverage Audit
- **Non-Traditional Schedule Analysis**: Filters and audits evening classes (starting $\ge 4:30\text{ PM}$) and Friday/weekend courses.
- Segregates offerings into early evening (4:30–6:30 PM), late evening (after 6:30 PM), Friday, and weekend cohorts.

### 7. 📊 Enrollment Efficiency & FTES Loss Audit
- **Fill Rates & Capacity Utilization**: Audits student enrollment against seat capacity for every section.
- **Institutional Low-Enrolled Alerts**: Flags low-enrolled sections based on college guidelines:
  - Capacity $\ge 30$: Under-enrolled if $<20$ active students.
  - Capacity $17$–$29$: Under-enrolled if $<17$ active students.
  - Small classes ($<17$ seats): Under-enrolled if fill rate $<70\%$.
- **Lost FTES Tracking**: Quantifies unrealized Full-Time Equivalent Students (FTES) lost from unfilled seats to support enrollment management decisions.

---

## ⚠ Room Conflict & Stacked Classes Engine

The application includes an automated conflict detection engine that evaluates every physical meeting occurrence:

- **Real Room Conflicts**: Detects overlapping meetings in the same room on the same day and time taught by **different instructors** or across unrelated programs.
- **Stacked Classes Recognition**: Distinguishes intentional concurrent courses (same instructor, cross-listed sections, or known cohort programs such as `AMT`, `NURADN`, `DENTAL`, `RADTEC`, `AUTOTEC`, `NURVN`) and marks them as **Stacked** rather than conflicting.
- **Fast-Track Sequential Recognition**: Recognizes sequential FT1 and FT2 sessions sharing a room across non-overlapping date spans as stacked.
- **Virtual & Online Suppression**: Classes scheduled with `Location: Online Class` or `Building: Canvas - Online` (as well as `Canvas`, `Via`, and virtual rooms) are recognized as online spaces and **never** produce room conflicts.
- **One-Click Conflict Isolation**: The "View Room Conflicts" toggle instantly filters the grid to display only conflicting sections and opens a detailed conflict resolver modal.

---

## 🔍 Multi-Dimensional Cascading Filters

Schedulers can isolate exact subsets of the schedule using real-time cascading filters in the collapsible sidebar:

- **Academic Term**: Supports files containing multiple terms (e.g., `2026/SU`, `2026/FA`, `2027/SP`) with chronological sorting. Defaults to the earliest term.
- **Location (Campus)**: Filter by Chino Campus, Fontana Campus, Rancho Campus, or Online Class.
- **Building & Room**: Cascades dynamically based on the selected campus.
- **Program of Study**: Filter by active degree or certificate curriculum with integrated search.
- **General Education**: Filter by GE Pattern (Cal-GETC, Chaffey Local) and specific GE Areas.
- **Academic & Career Community (ACC)**: Filter by community (Arts, Business, Health, STEM, etc.).
- **Subject & Course**: Multi-select filters with live search.
- **Modality & Calendar Session**: In-Person, Online, Hybrid; Full-Term, Fast-Track 1, Fast-Track 2.
- **Day / Evening & Section Traits**: Filter by day/evening classification or specific traits (e.g., Zero Textbook Cost, Friday Class).
- **Quick Global Search**: Instant free-text search across courses, sections, faculty names, rooms, and traits.
- **Active Filter Pills**: Displays all applied filters in a top badge bar with one-click individual or bulk removal.

---

## 📥 Data Ingestion & Compatibility

- **Informer Schedule Exports**: Compatible with `.xlsx`, `.xls`, and `.csv` exports from Chaffey College Informer (`Schedule Sections Building-Room Grids`).
- **Drag & Drop**: Drop any schedule spreadsheet anywhere on the browser window to parse instantly.
- **Zero-Maintenance Program Curriculum Auto-Discovery**: Automatically discovers and loads the latest `Program Summary*.csv` curriculum file on page load, requiring zero configuration when program requirements are updated.
- **Schedule Deduplication**: Automatically detects and consolidates duplicate schedule rows (e.g., duplicated lab components) to preserve clean calendar displays.

---

## 📤 Export & Usability Features

- **High-Resolution PNG Export**: Download clean calendar graphics for schedule review meetings.
- **Vector PDF Export**: Generate multi-page printable PDF reports of any active grid view.
- **Filtered CSV Export**: Export the currently filtered course sections for spreadsheet analysis.
- **Light / Dark Mode**: Toggle between light and high-contrast dark themes.
- **Section Details Modal & Popovers**: Hover over any calendar block for a summary tooltip, or click for the full modal detailing faculty, enrollments, FTES, dates, traits, and applicable programs.

---

## 🚀 How to Use

1. **Open the Application**: Navigate to the [live site](https://joecascio2000.github.io/section-schedule-grid/).
2. **Load Data**: Click **Upload Spreadsheet** (or drag and drop your schedule `.xlsx`/`.csv` export). If program files exist in the directory, active programs are loaded automatically.
3. **Select Term & Campus**: Choose the desired Academic Term and Campus Location from the left sidebar.
4. **Choose a View**:
   - Use **Room View** or **Day View** to inspect physical room scheduling.
   - Use **Program Coverage** to audit degree completability.
   - Use **GE Coverage** to check general education breadth offerings.
   - Use **Efficiency Audit** to analyze fill rates and identify under-enrolled sections.
5. **Resolve Conflicts**: Click **View Room Conflicts** to isolate any scheduling errors.
6. **Export**: Use the **Export** menu to save the view as a PNG, PDF, or CSV.
