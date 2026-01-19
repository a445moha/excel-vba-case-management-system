# Excel VBA Case Management & Workforce Planning System

A fully automated Excel system built with VBA to simulate real-time contact intake, SLA tracking, staff workload balancing, and workforce planning. Dynamically assigns cases, highlights overdue tasks, and models staffing needs under varying demand scenarios.

![Analytics Dashboard](<screenshots/Screenshot 2026-01-18 013425.png>)


## Table of Contents
- [All Files & How to Run](#all-files--how-to-run)
- [Core Features](#core-features)
- [System Overview](#system-overview)
- [**Demo Screenshots**](#demo-screenshots)
  - [Analytics Dashboard](#1-analytics-dashboard)
  - [Case Tracking & Deadline Monitoring](#2-case-tracking--deadline-monitoring)
  - [All Contacts – System of Record & Automation Engine](#3-all-contacts--system-of-record--automation-engine)
  - [New Case / Contact Intake & Case Closure Forms](#4-new-case--contact-intake--case-closure-forms)
  - [Staff Workload & Capacity Planning](#5-staff-workload--capacity-planning)
- [Implementation Notes](#implementation-notes)
- [References](#references)

## All Files & How to Run
- [All Files](./docs/all-files.md)


## Core Features
#### Case Intake & Lifecycle Management
- Log contacts and complaints in real time via VBA UserForms  
- Automatically record creation and resolution timestamps  
- Track case status (Open / Closed) with SLA enforcement  
- Highlight overdue cases automatically across the system 

#### Deadline & SLA Monitoring
- Dynamic “Days Until Deadline” using business days
- Flag deadline adherence (Overdue / Met) with conditional formatting  
- Support holiday-aware deadlines with custom holiday table  
- Provide a live view of backlog and operational urgency  

#### Staff Workload & Assignment Engine
- Track per-staff active cases, average handling time, and daily workload  
- Assign cases automatically based on capacity and priority  
- Monitor historical workload trends for trend analysis  
- Adjust dynamically as staff members or case volumes change 

#### Workforce Planning & Capacity Modelling
- Simulate monthly case volumes based on real regulatory data  
- Model processing times by case subtype and workload stress factor  
- Calculate required staffing under different demand scenarios  

## System Overview
- [Excel File Technical Overview](./docs/system-overview.md#excel-file-technical-overview)
- [All Modules & Forms Overview](./docs/system-overview.md#modules--forms-overview)
  - [Modules](./docs/system-overview.md#modules)
  - [Forms](./docs/system-overview.md#forms)

## Demo Screenshots
The following screenshots illustrate the system’s functionality from a user and operations perspective. They are included to demonstrate workflow, automation, and decision-support features rather than low-level implementation details.

#### 1. Analytics Dashboard

![Analytics Dashboard](<screenshots/Screenshot 2026-01-18 013425.png>)

Higher Quality Images: https://drive.google.com/file/d/1UNEaOhrkQVaAhlUB6RII3isPVOTDCpgV/view?usp=sharing, https://drive.google.com/file/d/1hs-WBmIecy26uAUeFbdh_qgKrnfqWN2d/view?usp=sharing 

Executive-style dashboard built using PivotTables and PivotCharts to analyze complaint trends and operational performance.
Key insights:
- Analyze complaint trends with PivotTables and PivotCharts  
- Visualize staff workload and escalation patterns  
- Identify bottlenecks and compliance risks dynamically  
- Updates automatically as new cases are added  

#### 2. Case Tracking & Deadline Monitoring

![Case Tracking](<screenshots/Screenshot 2026-01-18 113725.png>)

Displays all active and resolved cases with automated deadline tracking.
Key features:
- View active and resolved cases with automated deadlines  
- Highlight overdue cases with conditional formatting  
- Track live backlog and urgency  

#### 3. All Contacts - System of Record & Automation Engine

![alt text](<screenshots/Screenshot 2026-01-18 120245.png>)

Centralized repository for all contacts and complaints with lifecycle tracking.  
Key features:
- Central repository for all contacts and complaints  
- Capture timestamps automatically and simulate daily operational load  
- Enforce deadlines dynamically with row-level flags  
- Feed live data into staff assignment and reporting charts  
- Scalable table-driven design supports new staff, case types, and volumes 

#### 4. New Case / Contact Intake Form & Case Closure Form

![New Case/Contact Entry](<screenshots/Screenshot 2026-01-18 114728.png>)

UserForm for entering new contacts and complaints.  
Key features:
- Enter new contacts with automated priority and staff assignment  
- Calculate deadlines based on business rules  
- Reduce manual entry errors  

![Case Closure Form](<screenshots/Screenshot 2026-01-18 113208.png>)

UserForm for closing cases and finalizing outcomes.  
Key features:
- Record resolution dates automatically  
- Validate resolution against deadlines  
- Update status instantly across all views  
- Enable accurate SLA and compliance reporting 

#### 5. Staff Workload & Capacity Planning

![Staff Workload & Planning](<screenshots/Screenshot 2026-01-18 114016-1.png>)

Operational view for monitoring staff workload and staffing needs.  
Key features:
- Monitor per-staff case load and average processing time  
- Track historical workload trends  
- Calculate required staffing based on volume  
- Model capacity under pressure with adjustable stress factor

## Implementation Notes
- Modular VBA handles staff assignment, deadline calculation, and workload modeling  
- Dates respect business days and statutory holidays  
- Scales dynamically as staff or case volume changes  
- VBA code available in the `/vba` directory  


## References
Ontario Energy Board Statistics Used:
- [OEB Complaint Statistics](https://www.oeb.ca/consumer-information-and-protection/oebs-consumer-protection-role/complaint-statistics)
- [OEB Compliance Report 2024-2025](https://www.oeb.ca/sites/default/files/Compliance%20report-2024-2025-ENGLISH.pdf)
