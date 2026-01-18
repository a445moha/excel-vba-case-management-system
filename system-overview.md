# System Overview

This Excel VBA system simulates a case management workflow with real-time intake, SLA tracking, staff workload balancing, and capacity planning.

## Excel File Technical Overview
1. **Analysis Sheet**: 
    - PivotTables and PivotCharts of tblContacts, tblCases and tblStaffLoad show complaint trends, escalation patterns, and staff workload.
2. **Case Tracking Sheet**: tblCases
    - tblCases is tblContacts automatically filtered to only show Open cases. 
    - Conditinal formatting applied to "Priority", "Days Until Deadline" and "Deadline Adherence" columns.
    - "Days Until Deadline" calculated using NETWORKDAYS().
    - "Deadline Adherence" Calculated using IF().
3. **All Contacts Sheet**: tblContacts
    - "Generate Data - Day" button generates one day of sample data using VBA module modGenerateDay.bas.
    - "Clear Data" button deletes all ListRows in tblContacts.
    - "Log a New Contact" buttons opens userform ManualEnterContactForm.frm where user enters details when a contacts is made, date, time, and deadline are automatically recorded using VBA.
    - "Update a Case Status" button opens ManualCaseClosed.frm where user enters ID and Outcome of a case, status and remaining columns of tblContacts are automatically filled in using VBA.
    - All buttons funtion via modules assigned to shapes.
    - Conditional formatting automatically highlights rows of cases closed after dealine in red.
4. **Staff Sheet**: tblStaff, tblHoursDist, tblStaffLoad, Stress Testing 
    - tblStaff automatically filled in using VBA code based on data from tblContacts.
    - "Click to Add a Staff Member" button adds a new staff member column to tblStaff using VBA module AddStaffMember.bas. Program adapts automatically.
    - tblHoursDist contains operations data used to calculate number of staff needed.
    - User inputs stress factor, monthly volume multiplied by said factor and number of employees required is outputted.
    - tblStaffLoad stores data from tblStaff over time using VBA module modLoadAnalysis.bas (currently does not automatically consider manual input).
5. **Planning**: tblHolidays, tblTypeDistribution
    - tblHolidays contains statutory holiday dates to be omitted from workdays by program.
    - tblTypeDistribution is purely my planning out of probabilities for accurate data generation using information available on OEB website.

## Modules & Forms Overview
#### Modules
1. **AddStaffMember.bas**
    - Adds a new column to tblStaff with sample staff name.
2. **DeleteAllRows.bas**
    - Deletes all ListRows in tblContacts
3. **macroRefreshTblCases.bas**
    - Automatically refreshes and sorts tblCases
    - Called upon worksheet activation.
4. **modAssignStaff.bas**
    - Reads tblStaff to find staff member with lowest load and assigns to them.
5. **modLoadAnalysis.bas**
    - Generates a new row of tblStaffLoad using current information from tblStaff.
6. **modUpdateStatuses.bas**
    - Updates statuses in simulated timeframe of generated data.
7. **modUpdateTblStaff.bas**
    - Updates tblStaff using current data from tblContacts.
8. **modGenerateDay.bas**
    - Uses a compilation of functions to generate one day worth of realistic data. 
#### Forms
1. **ManualEnterContactForm.frm**
    - User enters details when a contacts is made, date, time, and deadline are automatically recorded.
2. **ManualCaseClosedForm.frm**
    - User enters ID and Outcome of a case, status and remaining columns of tblContacts are automatically filled in.