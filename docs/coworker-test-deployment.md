# Coworker Test Deployment

Use this when the tracker needs a public test URL instead of `localhost`.

## GitHub Source

Repository path:

```text
JustinMojica/trackerdashboard
```

Repository URL:

```text
https://github.com/JustinMojica/trackerdashboard
```

Use this GitHub path when Azure asks for the source repository. For the current
prototype, the deployment ZIP in `deploy-artifacts` is still the safest way to
publish the exact local version because the latest large React file has local
changes that are ahead of GitHub.

## Recommended Host

Use one Azure App Service Node.js app for the test deployment. This moves the
tracker off the local machine and gives coworkers a real HTTPS URL.

The app already serves both:

- the built React frontend from `dist`
- the secure Node API from `server/secureAccessServer.mjs`

That keeps Microsoft OAuth callbacks, cookies, and API calls on one public origin.

Do not use GitHub Pages for this version. GitHub Pages can host static files,
but this tracker needs the Node server for Microsoft sign-in, verification-code
email, admin approvals, and secure sessions.

## Azure Portal Path

1. Go to `https://portal.azure.com`.
2. Create an **App Service**.
3. Choose:
   - Publish: Code
   - Runtime stack: Node LTS
   - Operating system: Linux
   - Pricing plan: Free or Basic for testing
4. After Azure creates it, copy the app URL:

```text
https://<app-name>.azurewebsites.net
```

That URL is what you send to a coworker after the environment variables and
redirect URI are configured.

## Required App Settings

Set these in the hosting provider environment settings. Do not commit real secret values.

```text
MICROSOFT_TENANT_ID=...
MICROSOFT_CLIENT_ID=...
MICROSOFT_CLIENT_SECRET=...
MICROSOFT_MAIL_FROM=justin.mojica@mosaic-international.com
TRACKER_SESSION_SECRET=<long random value>
TRACKER_ALLOWED_EMAIL_DOMAINS=mosaic-international.com
TRACKER_ADMIN_EMAILS=justin.mojica@mosaic-international.com
TRACKER_PUBLIC_ORIGIN=https://<app-name>.azurewebsites.net
TRACKER_FRONTEND_ORIGIN=https://<app-name>.azurewebsites.net
TRACKER_USER_STORE=local
```

For the first coworker test, `TRACKER_USER_STORE=local` is acceptable. Move it to
`microsoft-lists` after the `Tracker Users` list is created and the list ID is known.

## Microsoft Entra Redirect URI

Add this Web redirect URI to the existing Audit Assignment Tracker app registration:

```text
https://<app-name>.azurewebsites.net/api/auth/callback
```

Keep the local redirect URI too if you still want local development:

```text
http://localhost:8787/api/auth/callback
```

## Build And Start

Azure should run this startup command from `package.json`:

```text
npm start
```

For ZIP deployments, enable build automation with:

```text
SCM_DO_BUILD_DURING_DEPLOYMENT=true
```

The server honors the platform `PORT` environment variable through
`server/productionStart.mjs`.

## Coworker Test Flow

1. Open `https://<app-name>.azurewebsites.net`.
2. Choose `Create one!` or `Request tracker access`.
3. Sign in with the company Microsoft account.
4. Enter the emailed verification code.
5. Admin signs in and approves the pending request.
6. Coworker signs in again and confirms they can see the correct tracker view.

## Checks

After deployment, sign in as admin and confirm the readiness panel shows:

- Secure access server: Ready
- Microsoft app token: Ready
- Graph admin consent: Ready
- Approval storage: Local for the first test, Microsoft Lists after migration
