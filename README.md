# Outlook Add In by Psomas — the Contract Review Queue button

The static web page behind the **Add to Contract Queue** button in Outlook for the Psomas
Legal Department, served by GitHub Pages at
`https://psomaswebteam.github.io/contract-queue-addin/`.

- `manifest.xml` — what IT uploads in the Microsoft 365 admin center (Integrated apps →
  Upload custom apps → Office Add-in → link to manifest).
- `taskpane.html` / `taskpane.js` / `taskpane.css` — the side panel: reads the selected
  email, shows the Contract Review Queue questions, sends the answers to the
  "Contract Queue Intake" Power Automate flow, shows the result.
- `config.js` — generated from the desktop app's Python defaults (questions, settings,
  Paul's signature, email styles). Do not edit by hand.
- `assets/` — the button icons.

This site stores nothing. The flow's web address is typed in once per user and kept in
Outlook's roaming settings for that user; no secret lives in these files.

Published from the private repository PsomasWebTeam/Outlook-Add-In (`addin/` folder) by
`deploy_addin.py`.
