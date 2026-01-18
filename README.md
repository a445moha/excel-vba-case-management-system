# Excel VBA Case Management & Workforce Planning System
A fully automated Excel-based case management system built with VBA to simulate real-time contact intake, deadline tracking, staff workload balancing, and workforce planning. The system dynamically assigns cases, tracks SLA adherence, highlights overdue cases, and models staffing requirements based on historical case distributions and stress factors.

#### Case Intake & Lifecycle Management
- VBA userforms to log contacts and complaints in real time
- Automatic capture of received date/time and resolution timing
- Case status tracking (Open/Closed) with SLA enforcement
- Overdue cases automatically highlighted across the system

#### Deadline & SLA Monitoring
- Dynamic “Days Until Deadline” calculation using business days
- Holiday-aware deadlines (custom 2024 holiday table)
- Conditional formatting for urgency and SLA violations
- Real-time adherence flags (“Deadline Met” vs “Overdue”)

#### Staff Workload & Assignment Engine
- Dynamic staff table with auto-adaptation when staff are added
- Tracks active cases, average handling time, and daily workload
- Automated case assignment logic based on capacity
- Historical staff load tracking for trend analysis

#### Workforce Planning & Capacity Modelling
- Monthly case volume simulation based on real regulatory data
- Case subtype distributions and processing-time modeling
- Adjustable workload stress factor
- Outputs required staffing levels under different demand scenarios

## Screenshots
The following screenshots illustrate the system’s functionality from a user and operations perspective. They are included to demonstrate workflow, automation, and decision-support features rather than low-level implementation details.

#### 1. Analytics Dashboard
![](https://drive.google.com/file/d/1DVvEuP4EIsQW4z-N3F-SZjk5fR_fiwVT/view?usp=sharing)
Executive-style dashboard built using PivotTables and PivotCharts to analyze complaint trends and operational performance.
Key insights shown:
- Tracks regulatory compliance and deadline adherence across all complaint categories
- Analyzes escalation behavior by complaint subtype and utility to surface risk patterns
- Visualizes staff workload over time to support capacity and staffing decisions
- Identifies process bottlenecks and compliance risks through dynamic, filter-driven charts
This dashboard is fully dynamic and updates automatically as new cases are generated or entered through the system.

#### 2. Case Tracking & Deadline Monitoring
![](https://drive.google.com/file/d/1mmkx-XKBzWY2Rc-InEiwNKsMB-qdUZz1/view?usp=sharing)
Displays all active and resolved cases with automated deadline tracking.
Key features shown:
- Days Until Deadline calculated dynamically
- Deadline Adherence status (Overdue / Deadline Met)
- Conditional formatting to visually flag overdue cases
- Live view of operational backlog and urgency

#### 3. All Contacts - System of Record & Automation Engine
![](https://drive.google.com/file/d/1ych23Pl0lZS2mWSLman3PkqLKQWHslCg/view?usp=drive_link)
Key features shown:
- Centralized case repository: Stores all contacts and complaints with full lifecycle tracking
- Real-time data capture: Automatically records system date and time for case creation and closure to simulate live operations
- Hybrid input model:
    - Manual VBA UserForms for realistic case entry and closure
    - Automated “Generate Day” function to simulate daily operational load
- Deadline enforcement:
    - Resolution deadlines calculated dynamically
    - Entire row flagged when deadlines are missed to surface compliance risk immediately
- Workload integration: Feeds live data into staff assignment logic, processing-time calculations, and downstream PivotCharts
- Scalable design: Table-driven structure allows new staff, case types, and volumes without breaking automation

#### 4. New Case / Contact Intake Form & Case Closure Form
![New Case/Contact Entry](https://drive.google.com/file/d/1axk7N8kOXJqZwU3-t_kIqSThvzuXCOfG/view?usp=drive_link)
UserForm used to enter new contacts and complaints in real time.
Key features shown:
- Automated capture of date and time received
- Dynamic assignment of priority and staff member
- Deadline calculation based on case type and business rules
- Eliminates manual data entry errors
![Case Closure Form](https://drive.google.com/file/d/1lidBGsSBHnCKsqLKFOOqSECd1wAGYnU2/view?usp=sharing)
UserForm used to close cases and finalize outcomes.
Key features shown:
- Automatic recording of resolution date
- Validation against resolution deadlines
- Status updates reflected instantly across tracking views
- Enables accurate SLA and compliance reporting

#### 5. Staff Workload & Capacity Planning
![](https://drive.google.com/file/d/1ZF0YCI1lDbLOWdCJNcKYjTs3Ds5UH7P-/view?usp=sharing)
Operational planning view used to monitor staff workload and staffing needs.
Key features shown:
- Per-staff case load and average processing time
- Historical workload tracking for trend analysis
- Workforce requirement calculations based on case volume
- Adjustable stress factor to model capacity under pressure

#### Implementation Notes
- Business logic is implemented using modular VBA (staff assignment, deadline calculation, workload modeling).
- All dates respect business days and statutory holidays.
- The system is designed to scale dynamically as staff members or case volume change.

*Core business logic is implemented in modular VBA code (staff assignment, deadline calculation, workload modeling), available in the /vba directory.*

OEB Statistics Used:
- https://www.oeb.ca/consumer-information-and-protection oebs-consumer-protection-role/complaint-statistics
- https://www.oeb.ca/sites/default/files/Compliance%20report-2024-2025-ENGLISH.pdf