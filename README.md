# daggerheart-encounter-builder
A Google Apps Script for Sheets that helps automate building encounters and managing custom adversaries and environments.

Features
* Encounter Builder: includes adversary query UI and Battle Point calculations. Creates Sheets in a google docs spreadsheet that summarize adversary stat blocks and have HP and Stress boxes to mark.
* Custom Adversaries
* Custom Environments
* Import core Adversaries and Environments from a parse of the SRD (forked from seansbox/daggerheart-srd)

To use
* From a google docs sheet open Extensions -> Apps Script
* Erase everything in the editor and paste in the entirety of code.gs
* Save and give the script needed permissions. NOTE: the script needs permission to create and erase files on your drive! Do not allow that if you can't independantly trust this script.
* Then reload the sheet, you should see a Daggerheart menu appear

App features will create a "Encounter Data" folder in the same folder as the Sheet for storing data. Most settings in `config.json` can also be configured through the Settings dialog.

This encounter automation system is licensed under GPL 3.0.
Daggerheart game content and mechanics are property of Darrington Press.
