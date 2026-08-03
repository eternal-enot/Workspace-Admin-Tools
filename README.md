# Admin Tools for Google Workspace for Education

Automate account provisioning and lifecycle management directly from Google Sheets.

## Project Overview
This project provides a suite of Google Apps Script (GAS) tools designed to simplify the administration of corporate accounts for educational institutions. It consists of two main parts:
1. **Main Admin Dashboard (Main_Admin_Table):** Processes requests from Google Forms, automatically generates corporate email addresses, assigns temporary passwords, and routes users to the correct Organizational Units (OU).
2. **Target Group Sheets (Target_Bachelors, Target_Masters, Target_PhD, Target_Staff):** Spreadsheets for direct management of specific user groups (Bachelors, Masters, PhDs, Staff). These handle activity monitoring, tracking enrollment statuses, and archive management.

## Features
- **Automated Provisioning:** Generates standardized domain email addresses based on user input (first and last names).
- **Smart OU Routing:** Automatically assigns the correct Organizational Unit (OU) path based on the user's role (Student/PhD/Staff) and department.
- **Email Notifications:** Automatically sends welcome emails with corporate credentials to users' personal addresses upon successful deployment.
- **Lifecycle Management:** 
  - Regular checking of account statuses (active/suspended/deleted).
  - A convenient "Archive" system for managing expelled students or those moving to the next degree level (e.g., from Bachelor to Master).
  - Automated relocation of inactive users to archive OUs.
  - Sending advance warnings about impending account deletions and the ability to bulk restore users from the archive in one click.

## Prerequisites
- Admin access to Google Workspace.
- **Admin SDK Directory API** enabled in the Google Apps Script Advanced Services for each project.
- [Google Clasp](https://github.com/google/clasp) installed globally (`npm install -g @google/clasp`).

## Getting Started

1. **Clone the repository**
   ```bash
   git clone https://github.com/eternal-enot/Workspace-Admin-Tools.git
   cd Workspace-Admin-Tools
   ```

2. **Configure Secrets**
   Copy the example configuration file in the main directory and fill in your domain and target spreadsheet IDs.
   ```bash
   cp "Main_Admin_Table/Secrets.js.example" "Main_Admin_Table/Secrets.js"
   ```

3. **Link Your Sheets**
   For each module (folder), you need to authenticate and link it to the corresponding Google Apps Script project created in your spreadsheet.
   
   *Example for the Main Dashboard:*
   ```bash
   cd Main_Admin_Table
   clasp login
   clasp clone <YOUR_MAIN_SPREADSHEET_SCRIPT_ID> 
   # This will create a .clasp.json file.
   ```

4. **Deploy**
   Upload the local code to Google Apps Script:
   ```bash
   clasp push
   ```
   *Repeat steps 3 and 4 for each target folder (`Target_Bachelors`, `Target_Masters`, `Target_PhD`, `Target_Staff`), using their respective script IDs.*

## Repository Structure
- `/Main_Admin_Table/` — Main module for account routing, temporary password generation, and initial deployment.
- `/Target_Bachelors/` — Module for managing Bachelor students (includes logic for archiving and transferring graduates to Master's programs).
- `/Target_Masters/` — Module for managing Master students (includes logic for archiving and transferring graduates to PhD programs).
- `/Target_PhD/` — Module for managing PhD students.
- `/Target_Staff/` — Module for managing faculty and staff.
