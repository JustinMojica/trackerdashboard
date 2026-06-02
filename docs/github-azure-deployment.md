# GitHub to Azure Deployment

The repo includes `.github/workflows/azure-app-service.yml` for deploying the tracker to Azure App Service.

## Required GitHub Settings

Add this repository secret:

- `AZURE_WEBAPP_PUBLISH_PROFILE`: the publish profile downloaded from the Azure App Service.

Add this repository variable after the secret is ready:

- `ENABLE_AZURE_DEPLOY=true`

The workflow also supports manual runs from **Actions > Deploy Azure App Service > Run workflow**.

## Azure App Service Settings

Keep these app settings configured in Azure:

- `MICROSOFT_TENANT_ID`
- `MICROSOFT_CLIENT_ID`
- `MICROSOFT_CLIENT_SECRET`
- `MICROSOFT_MAIL_FROM`
- `TRACKER_SESSION_SECRET`
- `TRACKER_ALLOWED_EMAIL_DOMAINS`
- `TRACKER_ADMIN_EMAILS`
- `TRACKER_PUBLIC_ORIGIN`
- `TRACKER_FRONTEND_ORIGIN`
- `TRACKER_USER_STORE`
- `TRACKER_USERS_SITE_ID`
- `TRACKER_USERS_LIST_ID`
- `TRACKER_PROJECT_STORE`
- `TRACKER_PROJECTS_SITE_ID`
- `TRACKER_PROJECTS_LIST_ID`
- `SCM_DO_BUILD_DURING_DEPLOYMENT=true`
- `WEBSITE_NODE_DEFAULT_VERSION=22`

Use `TRACKER_PROJECT_STORE=local` until the Audit Assignments Microsoft List has the `TrackerProjectJson` column and the list ID is configured.
