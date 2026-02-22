
# Product Specification: AutoShift - Automated Shift Scheduler

**Project Name**: AutoShift
**Owner**: To be assigned
**Date**: 2026-02-18
**Status**: DRAFT

---

## 1. Introduction
### 1.1 Purpose
AutoShift is a software solution designed to automate the complex task of creating optimal work arrangements (rosters) for a multi-dimensional work environment. It leverages mathematical optimization (Google OR-Tools) to satisfy rigid constraints and minimize manpower usage through smart scheduling (e.g., "Double Shifts").

### 1.2 Target Audience
- **Managers**: Who create schedules and manage workforce.
- **Employees**: Who view their schedules.

### 1.3 Scope
*   **In Scope**:
    *   Position/Station management (Work hours, Shifts, Manpower needs).
    *   Constraint configuration (Global & Per-Employee).
    *   Employee Management (Availability, Positions, Double Shift capability).
    *   Automatic Roster Generation using OR-Tools (CP-SAT).
    *   Streamlit-based UI for all interactions.
*   **Out of Scope**:
    *   Payroll integration (initially).
    *   Real-time attendance tracking.

---

## 2. Business Requirements
### 2.1 Goals & Success Metrics
*   **Goal**: Reduce scheduling time from hours to minutes.
*   **Goal**: Ensure 100% compliance with "Iron Rules" (constraints).
*   **Goal**: Optimize manpower utilization (minimize unstaffed shifts, maximize efficient use of available staff).

### 2.2 Assumptions & Constraints
*   **Solver**: Google OR-Tools (CP-SAT).
*   **Language**: Python.
*   **UI**: Streamlit.
*   **Input Data**: Excel (`availability.xlsx`).

---

## 3. Functional Requirements

### 3.1 Feature: Position Management
**Description**: Define stations and their requirements.
*   **Attributes**: Name, Days of Operation, Operating Hours, Shift Types (Morning/Afternoon/Night), Manpower per shift (Security Guards).
*   **Example**: Station A (24/7, 3 shifts, 2 guards/shift).

### 3.2 Feature: Employee Management
**Description**: Manage employee data and constraints.
*   **Attributes**: Name, Qualified Positions, Availability (Days/Shifts), "Double Shift" capability.
*   **Data Source**: Upload via Excel or manual entry.

### 3.3 Feature: Global Constraints ("Iron Rules")
**Description**: Configurable rules that the solver must respect.
1.  **No Overlapping Shifts**: Employee cannot be in 2 places at once.
2.  **No Back-to-Back Shifts**: e.g., working Morning then Afternoon is forbidden (unless specifically allowed).
    *   *Constraint*: No Night -> Morning (next day).
    *   *Allowed*: Afternoon -> Morning (next day).
3.  **Minimum Rest**: Specific hours between shifts.

### 3.4 Feature: "Double Shift" Optimization
**Description**: The system identifies opportunities to use "Double Shifts" (12h blocks) to cover manpower gaps efficiently.
*   **Logic**: If an employee is marked as "Double Shift Capable", the solver can assign them two consecutive shifts (e.g., Morning + Afternoon = Day Double) to count as 1 person covering 12 hours, potentially replacing 2 people.

### 3.5 Feature: Automatic Scheduler (The Solver)
**Description**: The core engine that takes all inputs and generates the optimal schedule.
*   **Output**: A valid roster satisfying all constraints.
*   **Visualization**: Interactive calendar/table view in Streamlit.
*   **Objective Function**: Minimize uncovered shifts, maximize employee preference/fairness (secondary).

---

## 4. UI/UX Design
### 4.1 Workflow
1.  **Setup**: Define Positions & Constraints.
2.  **Input**: Upload `availability.xlsx`.
3.  **Review**: Check parsed data.
4.  **Generate**: Click "Optimize".
5.  **Result**: View and Export Schedule.

### 4.2 Screens
*   **Dashboard**: Overview of current status.
*   **Configuration**: Settings for Positions/Constraints.
*   **Data Upload**: Drag-and-drop Excel.
*   **Schedule View**: Grid view of the generated roster.

---

## 5. Data Requirements
### 5.1 Input Format (`availability.xlsx`)
*   **To be confirmed**: Headers, merged cells?
*   **Assumed Columns**:
    *   `Employee Name`
    *   `Position`
    *   `Availability` (by day/shift)
    *   `Double Shift Capable` (Boolean)

---

## 6. Open Questions
*   [ ] What is the exact format of `availability.xlsx`? (Headers, column names?)
*   [ ] How are "Double Shifts" exactly defined (hours)? Is it always Morning + Afternoon? Or specific 12h blocks?
*   [ ] Are there break times mandated within a Double Shift?
