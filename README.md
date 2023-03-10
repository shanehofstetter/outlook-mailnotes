# Mailnotes

![logo](./assets/icon-128.png)

[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/donate/?business=PB26QWEQQ3RE4&no_recurring=0&item_name=Support+my+open+source+work+on+github+%E2%9D%A4%EF%B8%8F&currency_code=USD)


Simple Microsoft Outlook extension that allows you to add notes to an email.
The notes are be available across devices as they are stored using the Microsoft Office APIs (no data is shared with 3rd party services).


https://user-images.githubusercontent.com/13404717/209876856-bc68c302-6571-4d4c-9f35-a57bb426f5df.mov

## Features
- Notes are available across devices/outlook clients
- Supports dark/light themes (automatically detected)
- Save note either using button or by pressing <kbd>ctrl+s</kbd> or <kbd>cmd+s</kbd>
- Pinnable Taskpane

## Tested/Supported Applications
- Outlook for Mac (Version 16.x)
- Outlook for Web (using Chrome)

Outlook for Windows should work as well, but is not tested (feedback is appreciated).

## Installation

The Add-In is not yet available in the Marketplace, currently it has to be installed as a custom add-in.
To do that, follow these steps:

1. Open Outlook
2. Open Add-In Management
3. Go to "My Add-Ins"
4. Click "Add custom Add-In"
5. Select "From URL"
6. Insert the following URL: https://outlookmailnotes.z1.web.core.windows.net/manifest.xml
7. Install

## Run locally (OS X)

**1. Clone the Repo**

**2. Install the packages**

`npm install`

**3. Start the webpack server**

`npm run dev-server`

runs on port 3000 per default

**4. Install the Add-In in Outlook**

- choose "add from file.."
- select the manifest.xml from the repository
- installation should succeed and you should see the new Mailnotes Add-In in the taskbar, clicking on it should open the taskpane


## Roadmap
- Autosave
- Add an icon to mail items in directory if a note exists
