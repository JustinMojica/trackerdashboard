# GitHub to Azure Deployment

The repo includes `.github/workflows/azure-app-service.yml` for deploying the tracker to Azure App Service.

## Required GitHub Settings

Add this repository secret:

- `AZURE_WEBAPP_PUBLISH_PROFILE`: the publish profile downloaded from the Azure App Service.

Add this repository variable after the secret is ready:

- `ENABLE_AZURE_DEPLOY=true`

The workflow also supports manual runs from **Actions > Deploy Azure App Service > Run workflow**, but production deployments only run when the selected ref is `main` and `ENABLE_AZURE_DEPLOY=true`.

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

## Verification

After deployment, check:

```text
https://mosaic-audit-tracker-live.azurewebsites.net/
https://mosaic-audit-tracker-live.azurewebsites.net/api/auth/config
```

The homepage can return `200` while sign-in is still unavailable. The auth config endpoint must show `configured: true` before Microsoft sign-in and access requests will work.

If `/api/auth/config` lists missing values, restore those values in Azure App Service > Settings > Environment variables, then restart the app.

## Troubleshooting

GitHub Actions deploy failure:

- Check **Actions > Deploy Azure App Service**.
- If tests fail, fix the test/server issue first; deployment is intentionally blocked until checks pass.
- If deployment fails at the Azure deploy step, replace `AZURE_WEBAPP_PUBLISH_PROFILE` with a fresh publish profile.

Azure site loads but sign-in fails:

- Open `/api/auth/config`.
- Add any missing settings shown there to the Azure App Service environment variables.
- Confirm the Entra redirect URI matches the live callback URL.
- Restart the App Service.

Security rotation:

- If a publish profile is pasted into chat or a ticket, regenerate it in Azure and replace the GitHub secret.
- If a Microsoft client secret is exposed, create a new secret in Entra, update Azure, and delete the old secret.
