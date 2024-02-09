# README for CTI Launch Spreadsheet Automation

## Overview

This repository is designed to automate various administrative tasks within the CTI Launch program using Google Apps Script. The script adds functionality to Google Sheets for managing student data, tracking attendance, synchronizing gradebook updates, and facilitating other educational administrative processes.

## Features

- **Custom Menu Integration:** Automatically adds a custom menu to the Google Sheets UI, providing easy access to the script's functions.
- **Automated Attendance Tracking:** Processes and records student attendance using data from external sources.
- **Gradebook Synchronization:** Allows for the synchronization of data with external gradebook systems such as Canvas.
- **Milestone Tracking:** Automatically updates student progress milestones based on gradebook information.
- **Contract Monitoring:** Facilitates the generation of reports for students who have not signed contracts.
- **Session Management:** Assists in organizing and updating deep work session assignments and attendance.
- **Email List Generation:** Simplifies the creation of email lists for use with email merge tools like YAMM (Yet Another Mail Merge).

## Setup Instructions

### Cloning the Repository

Instead of copying and pasting the code directly, it's recommended to clone this repository to your Google Apps Script project using `clasp`, Google's command-line tool for Apps Script projects. This approach enables version control and local development.

1. **Install `clasp` Tool:**
   Ensure you have Node.js and npm installed on your machine. Install `clasp` globally using npm:
   ```bash
   npm install -g @google/clasp
   ```
   
2. **Login to Google Apps Script:**
   Authenticate `clasp` with your Google account:
   ```bash
   clasp login
   ```
   
3. **Clone the Script to Your Local Machine:**
   Create a new Apps Script project in Google Drive or use an existing one. Then, clone the project using `clasp`:
   ```bash
   clasp clone <script-id>
   ```
   Replace `<script-id>` with your script's ID, found in the Apps Script URL.

4. **Pull the Repository:**
   Navigate to your local project directory and initialize a Git repository if not already done. Pull the code from this GitHub repository into your local project directory.

5. **Push Changes to Google Apps Script:**
   After making changes or pulling updates from the repository, use `clasp` to push the updates to your Google Apps Script project:
   ```bash
   clasp push
   ```

### Initial Configuration

- After cloning the script to your Google Apps Script project, open the Google Sheet you wish to automate.
- The script automatically adds a custom menu titled "Code Executions" to the Google Sheets UI. Access this menu to utilize the provided functions.

## Usage

Select the desired operation from the "Code Executions" menu in your Google Sheet. Each menu item triggers a specific function in the script, such as updating attendance, synchronizing gradebook data, or generating reports.

## Scheduled Execution

Optionally, set up scheduled triggers within the Google Apps Script environment to automate routine tasks:

- Go to `Triggers` in the Apps Script interface.
- Click `Add Trigger`, select your function, set the event source to "Time-driven", and configure the frequency.

## Use Cases

This script is ideal for educational administrators, instructors, and program coordinators involved in the CTI Launch program, streamlining the management of student data and administrative tasks.

## Version Control

Utilize Git for version control to manage your script development effectively, allowing for collaboration and tracking of changes over time.

## Security Considerations

Ensure the security and privacy of student data, especially when integrating with external systems or sharing information electronically.
