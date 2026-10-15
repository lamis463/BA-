# Beyond Academics Transcript Automation

**Owner:** Undergraduate Program Office, Suliman S. Olayan School of Business (OSB), AUB
**Language:** MATLAB (runs in MATLAB Online; no installs needed on the work PC)

## What it does

Students log their extracurricular activities through a Microsoft Form. Its
responses go to an Excel workbook in OneDrive. This script turns those
responses into each student's official **Beyond Academics transcript**, a
Word document kept in OneDrive. Each student has one file, and it is updated
every time they submit.

For each response the script hasn't seen before, it:

1. Checks that the student is OSB, agreed to the terms, and gave a valid
   AUB ID. Responses that fail a check are skipped and written to
   `review_log.csv`.
2. Downloads that student's transcript, or starts a new one from the
   official template on their first submission.
3. Fills in the header (name, concentration and date issued). It adds the
   activity (semester, title and hours/days) to the matching category table.
4. Uploads the file back to the same place in OneDrive.

Processed responses are listed in `processed_log.csv`, so running the script
again never adds the same activity twice.

## Data flow

```
Microsoft Form ──► FormResponses.xlsx (OneDrive)
                        │  GET (Graph)
                        ▼
                MATLAB script ── edits .docx XML in a temp folder
                        │  PUT (Graph)
                        ▼
      Beyond Academics Transcripts/<Name>_<AUBID>.docx (OneDrive)
```

All files stay in the owner's OneDrive. Nothing is sent anywhere except
Microsoft's own login and Graph endpoints:

| Purpose | Endpoint |
|---|---|
| Sign-in (device code flow) | `login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode` and `/token` |
| Read a file's metadata and download URL | `GET graph.microsoft.com/v1.0/me/drive/root:/{path}` |
| Download a file | the short-lived `@microsoft.graph.downloadUrl` returned above |
| Create or replace a file (under 4 MB) | `PUT graph.microsoft.com/v1.0/me/drive/root:/{path}:/content` |

The script reads 3 files: the responses workbook, the template, and its own
logs. It writes into 2 folders: the transcripts folder and the
`_automation` log folder.

## Azure / Entra app registration (for IT)

| Setting | Value |
|---|---|
| Account type | Single tenant (AUB only) |
| Platform | Mobile and desktop applications (public client) |
| **Allow public client flows** | **Yes**. This is required for the device code sign-in. |
| Redirect URI | Not used by device code flow. `http://localhost` is harmless if one is required. |
| API permissions (delegated) | `Files.ReadWrite`, `offline_access` |
| Client secret / certificate | **None**. It's a public client, so no secret is stored anywhere. |

**Which account signs in:** the form and its files are owned by the
**sbstserv** account. With `Files.ReadWrite`, the script must sign in as
sbstserv, because that permission only covers the signed-in user's own
OneDrive. To sign in as a personal staff account instead, which only has
shared access to those files, the app would need `Files.ReadWrite.All`.

**Sign-in and tokens:** the user signs in at `microsoft.com/devicelogin` in a
normal browser, so MFA and conditional access apply as usual. The script
never sees a password. Only the refresh token is cached, in MATLAB's private
preferences folder (`prefdir`). It is never stored in this repository (see
`.gitignore`). Revoking the user's sessions or deleting the app registration
cuts off access straight away.

## Files

| File | Purpose |
|---|---|
| `run_transcripts.m` | The main script. Run this. |
| `ba_config.m` | All settings: tenant/client IDs and OneDrive paths. This is the only file to edit. |
| `GraphClient.m` | Sign-in and OneDrive read/write through Microsoft Graph. |
| `TranscriptDocx.m` | Edits the transcript `.docx` (Office Open XML) directly. Word isn't needed. |
| `tests/test_transcript_docx.m` | Offline test that fills a copy of the template with sample entries. |

## Setup

1. IT registers the app (see above) and provides the **Application (client) ID**
   and **Directory (tenant) ID**.
2. Paste both into `ba_config.m`. Check the OneDrive paths there. They are
   relative to the root of the signed-in account's OneDrive.
3. Put the official template at the `templatePath` given in `ba_config.m`.
4. Keep `outputFolder` pointed at a **TEST** folder for the first runs.
5. Open MATLAB Online, add this folder to MATLAB Drive (or clone this repo
   with MATLAB's Git support), and run `run_transcripts`.
6. On the first run, open the link it prints, enter the code, and sign in as
   the account that owns the files. Later runs reuse the session.

## How it gets triggered (open question)

The script processes **every new response each time it runs**. MATLAB Online
can't run on a schedule, so on the current setup a person clicks **Run** (for
example once a day). To make it fully hands-off, it needs to run on a
schedule on an always-on machine, such as a server or VM with MATLAB, or a
MATLAB Production Server/Compiler deployment. That's what to agree with IT.

## Known limits

- Each category box on the template shows about 3 entries. Extra entries are
  still added, but they can spill past the box. Students over the limit are
  flagged in `review_log.csv`.
- The name and concentration are written on the first submission and are not
  overwritten afterwards. "Date Issued" is refreshed on every update.
- Simple upload supports files under 4 MB. A transcript is about 180 KB.
