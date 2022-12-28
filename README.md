# Mailnotes

Simple Microsoft Outlook extension that allows you to add notes to an email.
The notes should be available across devices, they are stored using the Microsoft Offices APIs.

## Run locally (OS X)

### 1. Install the packages

`npm install`

### 2. Start the webpack server

`npm run dev-server`

runs on port 3000 per default

### 3. Install the Add-In in Outlook

- choose "add from file.."
- select the manifest.xml from the dist directory
- installation should succeed and you should see the new Mailnotes Add-In in the taskbar, clicking on it should open the sidebar

## Roadmap
- Save via CMD+s/Ctrl+s
- Show character limit / how many characters left