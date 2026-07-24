# Job Number Enforcer for Outlook

Forces a job number onto every meeting/appointment before Outlook will send it.
It's an Outlook **Smart Alerts** add-in running a hard block on send. If the
subject (or notes) doesn't contain a valid job number, the meeting stays in
Drafts and can't go out.

Works on: Outlook on the web, new Outlook for Windows, classic Outlook for
Windows (v2206+), Outlook for Mac (v16.65+).
**Does not work on Outlook mobile (iOS/Android)** — Microsoft doesn't support
on-send add-ins there yet. See "The mobile gap" below.

---

## What counts as a valid job number

Out of the box: the letter **J followed by 4–6 digits**, e.g. `J1234`, anywhere
in the meeting subject (or the notes/body). To change that rule, edit the one
`JOB_NUMBER_REGEX` line at the top of `commands.js` and re-upload the file to
your host. Examples are in the comments there. If job numbers must be validated
against a live list from another system, that's a bigger change — tell me and
I'll wire it up.

---

## Files

| File | What it is |
|---|---|
| `manifest.xml` | The add-in definition your admin uploads to Microsoft 365. |
| `commands.html` | Loads Office.js and the handler. Must be hosted. |
| `commands.js` | The on-send check. Must be hosted. **Edit config here.** |
| `assets/icon-*.png` | Placeholder icons. Must be hosted. Swap for your logo if you like. |

---

## Deploy it — two steps

### Step 1 — Host the files on HTTPS

The three files plus the `assets` folder must sit on any HTTPS web host at a
fixed URL. Free options that work: **Azure Static Web Apps**, **GitHub Pages**,
Netlify, Cloudflare Pages — or any web server your IT already runs. Upload so
the URLs resolve like this:

```
https://YOUR-DOMAIN/commands.html
https://YOUR-DOMAIN/commands.js
https://YOUR-DOMAIN/assets/icon-64.png   (etc.)
```

Then open `manifest.xml` and replace **every** occurrence of
`https://YOUR-DOMAIN.example.com` with your real host. There are several — do a
find-and-replace on the whole file.

### Step 2 — Push it to everyone (admin)

1. Go to the **Microsoft 365 admin center** → **Settings** → **Integrated apps**.
2. Choose **Upload custom apps** → **Upload manifest file** and select your edited `manifest.xml`.
3. Under **Assign users**, pick **Entire organization** (or the specific group).
4. Under **Deployment method**, choose **Fixed (default)** — this installs it for
   everyone and stops users from removing it, which is what makes the block
   un-bypassable.
5. Finish and accept the permissions. Rollout to all mailboxes typically takes
   up to 6–24 hours (sometimes longer for the first deployment).

That's it. After it lands, anyone trying to send a meeting without a job number
gets stopped with the message you set in `commands.js`.

---

## Test before org-wide rollout (recommended)

In step 2, assign to **Specific users** (just you) first. Then in Outlook on the
web, create a meeting with no job number and hit Send — you should be blocked.
Add `J1234` to the subject and it sends. Once happy, switch the assignment to the
whole org.

---

## The mobile gap (and how to close it if you need to)

The hard block covers desktop and web. On Outlook **mobile**, on-send add-ins
don't run, so a meeting created on a phone won't be checked. If you need mobile
covered too, the standard belt-and-braces is a **Power Automate** flow on "When a
new event is created" that flags or emails the organiser about any meeting
missing a job number — it can't stop the save, but nothing slips through
unnoticed. Say the word and I'll build that flow as well.

---

## Changing the rule later

Everything about *what's required* lives in the top block of `commands.js`
(`JOB_NUMBER_REGEX`, whether the body counts, and the error message). Edit,
re-upload that one file to your host, done — no need to touch the manifest or
redeploy in the admin center unless you rename files or change URLs.
