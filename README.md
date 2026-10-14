# Interactio for Outlook

> Hosted at **https://raimondasnauseda.github.io/ms-outlook-integration/** (GitHub Pages, `main` → `/docs`). `manifest.xml` in this repo is already built for that URL.

An Outlook **online-meeting provider add-in**. When someone creates a meeting in Outlook and picks **Interactio** (instead of Teams), the add-in adds their own saved Interactio link to the meeting, with a "Join Interactio meeting" link. It's the Outlook counterpart of *Interactio for Google Calendar*.

- **One link per user.** Each user saves their link once. It's stored in their mailbox's add-in settings (`roamingSettings`), so it follows them across Outlook web, Windows, Mac and mobile. There's no backend.
- **Only Interactio links.** Accepts HTTPS links on `interactio.com` / `interactio.io` and their subdomains. The validation logic is the same as in the Google add-in.
- **New users** who pick Interactio before saving a link get a setup dialog first. They never get someone else's link.
- **Narrow permission.** It only asks for `ReadWriteItem`, which means it can edit only the meeting that's open. It makes no Microsoft Graph calls.
- **Static files.** Everything is served from GitHub Pages (the `docs/` folder).

## Layout

```
docs/                  ← published by GitHub Pages
  link.js              link validation + meeting-block HTML (shared)
  commands.html/.js    UI-less function: insertInteractioMeeting
  settings.html/.js    settings task pane; also the setup dialog (?mode=dialog)
  assets/              grayscale button icons (#919191) + colour list icons
manifest.template.xml  add-in-only (XML) manifest: web, new + classic Windows, Mac, iOS, Android
scripts/build-manifest.js   fills the Pages URL into manifest.xml
scripts/make-icons.py       regenerates every icon from one logo PNG
test/run-tests.js      43 offline tests (validation, block merge, mocked Office.js flow)
```

Why XML only: the unified JSON manifest can't be sideloaded or line-of-business deployed on Mac or mobile. The XML manifest covers every client.

## What gets added to the meeting

A short divider, **Interactio meeting**, then "Join Interactio meeting: <link>" where the clickable text is the URL itself, and a closing divider. The URL appears only once, so plain-text views of the invite (Google Calendar, text-only mail) don't show it twice. To change the wording, edit `buildMeetingBlock()` in `docs/link.js`, and update the matching text in the replace regex in `mergeMeetingBlock()`.

- **Same link already in the body:** nothing is added.
- **An older Interactio block** (for example after the user changed their link): it's replaced, not duplicated.

## 1. Icons

All icons are generated from the Interactio logo in `src-assets/interactio-logo.png`:

- **Button icons** are grey (`#919191`) with the swirl cut out, as Microsoft requires.
- **Add-in list icons** use the magenta logo.

If the logo changes, run:

```bash
python3 scripts/make-icons.py src-assets/interactio-logo.png
```

## 2. Host on GitHub Pages

1. Create a repo, for example `interactio-outlook-addin`, under the Interactio GitHub org (public or internal), and push this folder.
2. In the repo, go to **Settings → Pages → Build and deployment**, choose **Deploy from a branch**, then branch `main` and folder `/docs`.
3. Note the URL, for example `https://<org>.github.io/interactio-outlook-addin/`. Check that `…/settings.html` opens.
4. Generate the manifest for that URL:

```bash
npm run manifest -- https://<org>.github.io/interactio-outlook-addin/
npm test
npx office-addin-manifest validate manifest.xml   # needs internet access to Microsoft's validator
```

If the repo or URL ever changes, regenerate `manifest.xml` and redeploy it. Keep the `<Id>` GUID unchanged.

## 3. Test (sideload)

In Outlook on the web, go to **Get Add-ins → My add-ins → Add a custom add-in → Add from file** and pick `manifest.xml`. Then:

1. Open **New event**. The **Interactio** button is on the ribbon, or under *Apps*. Click it: with no link saved, the setup dialog appears. Save a link, and the block is added.
2. Click it again. Nothing is duplicated.
3. Open **Interactio link** (the task pane). Change the link and click Interactio again: the block is replaced. Then click **Remove** and check that the setup dialog comes back.

Sideloaded add-ins are *user-installed*. Microsoft only shows the add-in **in place of the Teams toggle** when an admin installs it (step 4). The logic is the same, so sideloading is good enough for testing.

**First thing to check:** this manifest adds a second, task-pane button ("Interactio link") next to the function button in the meeting organizer. Microsoft allows only one *function* command there; a task-pane button should be fine. If the online-meeting toggle misbehaves after admin install, delete the `interactioSettingsButton` control. Users can still change their link through the setup dialog after pressing **Remove**. Re-adding that option would need a read-surface pane.

## 4. Roll out to the whole tenant (admin)

1. In the Microsoft 365 admin center, go to **Settings → Integrated apps → Upload custom apps**.
2. Choose **Office Add-in**, then **Upload manifest file (.xml)**, and pick `manifest.xml`.
3. Assign it to **Entire organization**, then Deploy. It can take up to 24 hours to reach every mailbox (usually faster on the web).
4. In a new meeting, users then see Interactio where the Teams toggle is, or as a choice next to it.

**Updating:**

- **Code changes** (anything in `docs/`): push to `main`. Pages redeploys and clients pick it up without reinstalling.
- **Manifest changes** (labels, icons, buttons): bump `<Version>` in `manifest.template.xml`, regenerate, then **Update** the app in Integrated apps.

## Default provider ("Add online meeting to all meetings")

Interactio **can't currently be the automatic default.** Microsoft limits "Make every meeting online" to specific partner add-ins: Zoom, Webex, GoTo, Google Meet, JioMeet (and previously BlueJeans). The Exchange setting `Set-MailboxCalendarConfiguration -DefaultOnlineMeetingProvider` only accepts a fixed list of providers. Microsoft's guidance is that other providers aren't supported for this feature.

What is possible: with the add-in admin-installed, Interactio appears on the compose screen as the online-meeting option. If you also want Teams out of the way, an admin can turn off the Teams Outlook meeting add-in through a Teams meeting policy (`AllowOutlookAddIn = False`) for chosen users. That removes Teams meetings from Outlook for those users, so treat it as a separate decision.

## Later: public Microsoft Marketplace

1. Submit through Partner Center (AppSource). Use the privacy, terms and support URLs on interactio.com.
2. After it's published, open a GitHub issue in `OfficeDev/office-js` titled "Outlook: Register the online-meeting template for Interactio" to get native **Join**/**Chat** buttons on events. Classic Outlook for Windows never shows those buttons.
3. Unlike Google's Marketplace, the same manifest and `<Id>` can be published, so no second project is needed.
