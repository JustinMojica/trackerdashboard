# Coworker Test Deployment

Use this when the tracker needs a public test URL instead of `localhost`.

## Recommended Host

Use one Azure App Service Node.js app for the test deployment. This app already serves both:

- the built React frontend from `dist`
- the secure Node API from `server/secureAccessServer.mjs`

That keeps Microsoft OAuth callbacks, cookies, and API calls on one public origin.

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

Azure should run:

```text
npm install
npm run build
npm start
```

The server now honors the platform `PORT` environment variable. If no public
origin is configured but Azure provides `WEBSITE_HOSTNAME`, the server infers:

```text
https://<WEBSITE_HOSTNAME>
```

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
