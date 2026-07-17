# Wesco Meeting Notebook — Azure package

This folder is a standalone Azure Static Web Apps package for the Wesco Meeting Notebook. It does not include the older Wesco Estimator Workflow.

## Azure deployment

- App location: `azure/Wesco Meeting Notebook`
- Output location: leave blank
- API location: leave blank
- Start file: `index.html`

After Azure deploys the site, share the generated `https://<name>.azurestaticapps.net` URL with staff.

## Saving behavior

The standalone review version saves meeting changes in the current browser. That is suitable for individual review and template use. For shared team records across devices, connect the notebook to Dataverse, a SharePoint List, or an Azure database/API.

The **Files & photos** tab accepts secure links to plans, estimates, change orders, contracts, vendor/customer documents, and job photos. Device-uploaded photos are browser-local; use SharePoint or OneDrive links when the whole team needs access.
