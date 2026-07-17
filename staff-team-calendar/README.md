# Master Wesco Staff Team Member Calendar

Standalone Azure Static Web Apps template connected to the Outlook calendar **Wesco Team Meeting Calendar**.

## Microsoft connection

- Entra application client ID: `cb882f2c-7555-427e-8484-92d6d0a21494`
- Wesco tenant ID: `5bfae2d1-ca2b-45bc-850a-90de45c02287`
- Delegated permissions: `User.Read`, `Calendars.ReadWrite`
- Outlook calendar owner: Theo Allen (`tallen@wesconc.com`)

The Entra app registration must contain the deployed site origin and `http://localhost:5500` as **Single-page application** redirect URIs.

## Existing Azure Static Web Apps deployment

The GitHub repository already contains an Azure Static Web Apps workflow with `app_location: "/"`. Publish this package at:

`staff-team-calendar/index.html`

The expected production address is:

`https://wonderful-field-08b34ac0f.azurestaticapps.net/staff-team-calendar/`

## Local review

From the HTML folder, run:

`python -m http.server 5500`

Then open:

`http://localhost:5500/master_wesco_staff_team_member_calendar.html`
