# Excel Template Portal

A small Next.js app that generates an Excel template for student data.

## Note to self:
- used Codex to prompt the whole thing
- installed GIT locally
- created repo on GITHUB
- pushed to repo from CODEX
- connected to repo from Vercel
- deployed to Vercel

## Features

- Professional landing page with a school-name input
- Downloadable `.xlsx` template
- 20 dummy rows plus 100 ready blank rows
- Email validation in Excel
- Separate school-name field above the table

## Local Development

Install dependencies:

```powershell
npm install
```

Start the development server:

```powershell
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) or the port shown in the terminal.

## Production Build

Run a production build locally:

```powershell
npm run build
npm run start
```

## Deploy To Vercel

### Option 1: Deploy from GitHub

1. Push this project to a GitHub repository.
2. Sign in to [Vercel](https://vercel.com).
3. Click `Add New...` -> `Project`.
4. Import the GitHub repository.
5. Confirm the project is detected as a Next.js app.
6. Keep the default build settings unless you have a custom repo structure.
7. Click `Deploy`.

### Option 2: Deploy with the Vercel CLI

Install the Vercel CLI:

```powershell
npm install -g vercel
```

Log in and deploy:

```powershell
vercel
```

For a production deployment:

```powershell
vercel --prod
```

## Notes

- The app uses a Next.js route handler to generate the Excel file on the server.
- No environment variables are required for the current version.
