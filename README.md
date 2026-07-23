# Wesco Dashboard App

Static HTML dashboard and operational templates deployed from GitHub to Azure Static Web Apps.

## Website addresses

- Planned primary address: <https://www.wescodashboardapp.com>
- Azure deployment address: <https://wonderful-field-08b34ac0f.7.azurestaticapps.net>

The repository deploys through `.github/workflows/azure-static-web-apps-wonderful-field-08b34ac0f.yml` whenever `main` changes.

## Custom-domain configuration

This project uses Azure Static Web Apps, not GitHub Pages. Do not add a GitHub Pages `CNAME` file. The custom domain must be added under the Azure Static App's **Custom domains** settings, and the public DNS record for `www` must point to the same Azure Static App.

After Azure validates `www.wescodashboardapp.com`, set it as the default domain so requests to the generated Azure hostname redirect to the public website address.

## Microsoft sign-in

Pages using Microsoft sign-in need their exact deployed URLs listed as Single-page application redirect URIs in the Microsoft Entra app registration. For the Projects page, register both addresses if both will remain usable:

- `https://www.wescodashboardapp.com/Wesco-Projects.html`
- `https://wonderful-field-08b34ac0f.7.azurestaticapps.net/Wesco-Projects.html`

The Projects page requests the delegated Microsoft Graph permissions `User.Read`, `Sites.Read.All`, and `Files.ReadWrite.All`. The page uses the write permission only to upload files into the currently open folder under `Documents / We App / Projects`; it does not expose delete, rename, or move actions. An Entra administrator must grant consent before staff uploads will work.

## Deployment layout

- App location: `/`
- API location: blank
- Output location: blank
- Start file: `index.html`
