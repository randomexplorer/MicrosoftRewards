# Bing Search Automator for Microsoft Rewards

A Tampermonkey userscript that automates Bing searches to help you earn Microsoft Rewards points easily and efficiently.

## Overview

This script automatically performs Bing searches with natural-looking search queries to help accumulate Microsoft Rewards points. It features a draggable floating button interface to control the automation and displays real-time progress.

## Features

- ✅ Performs up to 30 automated Bing searches with randomized intervals
- ✅ Generates natural-looking search queries across various topics
- ✅ Draggable floating control button with progress indicator
- ✅ Remembers position between sessions
- ✅ Pause/resume functionality
- ✅ Works in the background when tab is inactive

## Installation Requirements

### 1. Install a Userscript Manager

The script requires a userscript manager extension. We recommend Tampermonkey:

- [Tampermonkey for Chrome](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
- [Tampermonkey for Firefox](https://addons.mozilla.org/en-US/firefox/addon/tampermonkey/)
- [Tampermonkey for Edge](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)
- [Tampermonkey for Opera](https://addons.opera.com/en/extensions/details/tampermonkey-beta/)

### 2. Install the Script

#### Option 1: One-Click Installation
Click this link to install directly (requires a userscript manager):
[Install Bing Search Automator](microsoftRewards.user.js)

#### Option 2: Manual Installation
1. Click the Tampermonkey icon in your browser and select "Create a new script"
2. Delete any default code
3. Copy the entire content of the [`microsoftRewards.user.js`](MicrosoftRewards.user.js) file
5. Paste it into the Tampermonkey editor
6. Click File > Save or press Ctrl+S

## Usage

1. Navigate to [Bing.com](https://www.bing.com)
2. You'll see a blue "START" button in the corner of your screen
3. Click the button to start the automated searches
   - The button will turn red and display "STOP" while running
   - A counter will show your progress
4. Click again to stop automation at any time
5. The button can be dragged to any position on the screen

## Advanced Options

The script includes several configurable options in the `config` object:

- `minInterval`: Minimum time between searches (seconds)
- `maxInterval`: Maximum time between searches (seconds)
- `MAX_SEARCHES`: Maximum number of searches to perform

## Privacy & Safety Notes

- This script runs locally in your browser
- No data is collected or sent externally
- The script only interacts with Bing.com
- Using automation scripts may violate Microsoft's Terms of Service. Use at your own risk.

## Troubleshooting

**Script isn't working?**
- Make sure you're on a Bing.com domain
- Check if Tampermonkey is enabled
- Look for errors in the browser's console (F12)
- Try refreshing the page

## License

This project is open source and available under the [MIT License](LICENSE).

## Contributing

Feel free to fork this repository and submit pull requests with improvements or bug fixes.

## Disclaimer

This script is provided for educational purposes only. The author is not responsible for any consequences resulting from the use of this script, including but not limited to account suspension or termination. Use at your own risk.

---

*Last Updated: April 8, 2025*
