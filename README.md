# Mailnotes

![logo](./assets/icon-128.png)

Simple Microsoft Outlook extension that allows you to add notes to an email.
The notes should be available across devices, they are stored using the Microsoft Office APIs.

## Run locally (OS X)

### 1. Clone the Repo

### 2. Install the packages

`npm install`

### 3. Start the webpack server

`npm run dev-server`

runs on port 3000 per default

### 4. Install the Add-In in Outlook

- choose "add from file.."
- select the manifest.xml from the repository
- installation should succeed and you should see the new Mailnotes Add-In in the taskbar, clicking on it should open the taskpane

