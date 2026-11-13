# EMR Compliance Dashboard

A web-based audit and compliance tracking system built on **Google Apps Script** and **Google Sheets** for **Dawat-e-Islami Trust**. The application streamlines monthly audit query management, division-level responses, senior verification workflows, performance rankings, and automated email reminders across provinces and divisions.

- **Version:** 1.1 (October 2026)
- **Developed by:** Bluemoonways (`info.bluemoonways@gmail.com`)

---

## Key Features

- **Role-Based Access Control (5 Roles):** Granular permissions for `Admin`, `CO` (Provincial Officer), `Senior`, `Division`, and `Viewer`.
- **Bulk Query Entry:** Add audit queries across up to 17 standardized compliance sections in a single submission with built-in duplicate detection and section validation.
- **Interactive Executive Dashboard:**
  - Multi-month selection and province/division drill-downs.
  - Real-time KPI summary cards (`Total`, `Resolved`, `Awaiting Review`, `Returned`, `Pending Answer`).
  - **Aging Tracker:** Automatically calculates the age (in days) of the oldest unanswered query and highlights the responsible division(s).
  - **Visual Analytics:** Interactive SVG donut chart for section-wise distribution and stacked column charts comparing provinces or divisions.
- **Performance Ranking:** Tie-aware leaderboards ranking divisions and provinces by resolution percentage (`Resolved / Total`), highlighting top performers and areas needing attention.
- **Detailed & Section-Wise Reporting:**
  - **Detailed Report:** Complete query-by-query audit trail with answers, status badges, and verification timestamps.
  - **Section-Wise Report:** Aggregated metrics including resolution rate, return counts, average days to answer, and oldest pending query per section.
  - **Export Options:** Print-optimized layout and UTF-8 CSV export with formula-injection protection.
- **Security & Concurrency:**
  - HMAC-SHA256 signed session tokens (`7-day` expiry) stored client-side without persisting raw passwords.
  - Constant-time password comparison, brute-force lockout (5 failed attempts / 10-minute cooldown), and self-service password updates that invalidate other active sessions.
  - Server-side `LockService` concurrency control and stale-answer protection during senior reviews.
- **Automated Daily Reminders:** Scheduled 9:00 AM (`Asia/Karachi`) email notifications alerting divisions about pending/returned queries and seniors about answers awaiting review.

---

## User Roles & Permissions

| Role | Scope | Permissions |
| :--- | :--- | :--- |
| **Admin** | All Provinces & Divisions | Full access: add, edit, and delete queries (even after answered/verified); write or edit any answer; mark resolved or return; view all dashboards, rankings, and reports. |
| **CO** | Assigned Province only | Add, edit, and delete queries (before an answer is submitted) for divisions within their province; view provincial dashboards, rankings, and reports. |
| **Senior** | All Provinces & Divisions | Review submitted answers; mark queries as **Resolved (`Verified`)** or **Return (`Unverified`)** with mandatory remarks; view all dashboards, rankings, and reports. |
| **Division** | Assigned Division only | View own queries (prioritized by *Returned* and *Pending Answer*); submit or resubmit answers until resolved; view division reports. |
| **Viewer** | All Provinces & Divisions | Read-only access to all dashboards, queries, rankings, and reports. |

---

## Repository Structure

```text
├── Code.gs          # Google Apps Script backend: auth, API endpoints, caching, triggers, and Sheet menu
├── index.html       # Single-page frontend UI: responsive views, SVG charts, and client-side state
└── README.md        # Project documentation and deployment guide
