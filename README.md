# Outlook Calendar Panel for Obsidian

A sidebar in Obsidian that shows one day of your Outlook calendar. Drag an
event into a note to insert it as a Markdown line: it lands on a line of its
own, indented like the line you dropped it on. Or use an event's **+** button
to add it below the cursor. The command **Insert the day's agenda** adds every
event of the day at once.

- **‹ ›** step through days, and **Today** jumps back.
- Open a daily note and the panel shows that note's day. Leave it and the
  panel goes back to today, unless you picked a day yourself. The date is read
  from the file name with the **Daily note date format** setting.
- The event happening now is highlighted, and past ones fade. Events across
  midnight show as "From 22:00" or "Until 06:00".
- **Join** opens an online meeting. Click an event to see its organiser, the
  start of its description and an **Open in Outlook** link.
- Cancelled events are hidden, and so are declined ones unless you turn them on.
- An event the active note already holds has no **+** button. A daily note
  counts only for its own day.
- In a note, a small calendar icon at the end of each line that holds an
  event shows that event in the panel: the panel moves to its day, scrolls to
  it and it flashes. The icon is drawn by the editor and never written into
  the note; it appears in editing and live preview, not in reading view.

A note line holds an event when, trimmed, it is exactly what adding the event
would insert (its first line, for a template of several lines), or when it
is a list item whose text is the event's time and title, such as
`- 09:00–10:00 Standup` or `- [ ] 09:00 standup`. Case, extra spaces, the
kind of dash and a tail of tags, dates, emojis and links do not count. A title alone, such as `- Standup`, does not match. The icons cover
the days the panel has shown or read from its cache since Obsidian started.
When a line matches an event on several days, as a recurring one does, the
icon shows the one on the note's own day if it is a daily note, else the
panel's day, else the one nearest to today.

The plugin only reads your calendar. It never creates, changes or answers
anything in Outlook.

This plugin is not made by or affiliated with Microsoft. Outlook is a
trademark of Microsoft.

## Performance

The plugin is built not to slow Obsidian down:

- Loading the plugin only registers the panel, its commands and settings. It
  makes no network request and does not read the sign-in.
- The panel reads the sign-in and the last fetched events only once Obsidian
  has finished starting and is idle, shows those events at once, then
  fetches. Nothing waits for Microsoft.
- A panel that is not on screen (a collapsed sidebar or a background tab)
  does nothing. It catches up when you show it.
- Opening a note reads its file name. While the panel is on screen, typing
  and saving in the active note restart one shared 300 ms timer; when it
  fires, the panel reads the note's text once and compares it with the last
  check, and only a changed note is matched against the shown events. With
  the panel closed or hidden, these handlers return at once.
- The icons beside event lines only look at the lines on screen, only when
  the text, the visible part or the fetched events change, and each distinct
  line text is matched once, so moving the cursor costs nothing and typing
  costs one lookup for the line being typed. Until the panel has read its
  cache, there are no icons, so an editor opened at startup reads neither
  the sign-in nor the cache.
- One timer a minute moves the "now" highlight and refreshes a day older than
  **Refresh after** (15 minutes by default), and only while the panel is on
  screen.
- Plain JavaScript with no dependencies: one small `main.js`.

To check, open **Settings → Community plugins** and use Obsidian's startup
time report.

## Disclosures

- **Account required:** a Microsoft account (work, school or personal) with
  an Outlook calendar, and an app registration you create yourself (see Setup).
- **Network use:** the plugin talks only to Microsoft:
  `login.microsoftonline.com` to sign in and `graph.microsoft.com` to read
  your calendar. Nothing is sent anywhere else.
- **Clipboard:** the sign-in dialog writes the sign-in code to the clipboard
  when you select **Copy code**. The plugin never reads the clipboard.
- **No telemetry, ads or payments.** Event data is not written to your vault
  unless you drag or add an event into a note.
- **Cached on your device:** the last seven fetched days are kept in
  Obsidian's local storage for this vault (not in the vault's files), so the
  panel shows them at once at startup. Signing out clears them.
- **Stored on your device:** the sign-in token, in Obsidian's secret storage.
  The client ID, tenant and other settings are in the plugin's `data.json` in
  your vault.

## Setup

The plugin signs in to Microsoft itself. It needs an app registration in
Microsoft Entra so that Microsoft knows which app is asking. You register your
own, once.

**Already use the Microsoft To Do Panel plugin?** You can reuse its app
registration: add the **Calendars.Read** permission to it (step 7), then enter
the same client ID and tenant here. You still sign in once in this plugin,
because each plugin keeps its own sign-in.

### 1. Register an app

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com)
   and go to **App registrations → New registration**.
2. Name it, for example "Obsidian Outlook Calendar Panel".
3. Under **Supported account types**, pick one:
   - **Accounts in this organizational directory only** for a work or school
     account. You will then use your tenant ID in the plugin.
   - **Personal Microsoft accounts only**, or **Any organizational directory
     and personal Microsoft accounts**, for an Outlook.com or Hotmail account.
4. Leave **Redirect URI** empty and select **Register**.
5. On the app's **Overview** page, copy the **Application (client) ID**, and
   for a work account the **Directory (tenant) ID**.
6. Under **Authentication**, set **Allow public client flows** to **Yes** and
   save. The plugin signs in with a device code, which needs this.
7. Under **API permissions**, select **Add a permission → Microsoft Graph →
   Delegated permissions**, add **Calendars.Read**, and save. If your
   organisation requires it, ask an administrator to grant consent.

### 2. Connect the plugin

1. Open the panel and select **Sign in**, or open **Settings → Outlook Calendar Panel**.
2. Enter the client ID and tenant: your directory (tenant) ID for an app
   limited to your organisation, `consumers` for personal accounts only, or
   `common` for both.
3. Select **Sign in**. The dialog shows the client ID and tenant to check;
   select **Continue** and it shows a short code. Open the Microsoft page,
   enter the code and sign in. The dialog closes by itself and the panel fills.

Open the panel from the ribbon icon or with the command **Outlook Calendar Panel: Open panel**.

### If sign-in fails

| Message starts with | Fix |
| --- | --- |
| `AADSTS50059` | Set **Tenant** to your directory (tenant) ID. |
| `AADSTS700016` | Check the client ID and tenant. |
| `AADSTS7000218` | Turn on **Allow public client flows** in the app registration. |
| `AADSTS65001` | The permission is not granted yet; an administrator may need to approve it. |

Some organisations block device code sign-in altogether. If yours does, the
plugin cannot sign in there.

## Install

Requires Obsidian 1.11.4 or newer, for secret storage.

- **With BRAT:** add `Au2mater/obsidian-outlook-calendar-panel`.
- **Manually:** download `main.js`, `manifest.json` and `styles.css` from the
  latest release into `<vault>/.obsidian/plugins/outlook-calendar-panel/`,
  then enable the plugin under **Settings → Community plugins**.

## Settings

The settings tab shows your Microsoft account and the event template. The
rest is under **Advanced**, collapsed.

| Setting | Default | Notes |
| --- | --- | --- |
| Client ID | empty | From your app registration. |
| Tenant | `common` | Tenant ID, `common`, `organizations` or `consumers`. |
| Event template | `- {{time}} {{title}}` | `{{time}}`, `{{start}}`, `{{end}}`, `{{title}}`, `{{date}}`, `{{location}}`, `{{organizer}}`, `{{calendar}}`, `{{link}}`, `{{join}}` |
| Time format (advanced) | `HH:mm` | Moment.js format, such as `h:mm A`. |
| Daily note date format (advanced) | `YYYY-MM-DD` | Moment.js format of daily note file names. |
| Follow the daily note (advanced) | on | Show the day of the daily note you open. |
| Show declined events (advanced) | off | Shown struck through. |
| Refresh after (advanced) | `15` minutes | How old a shown day may get before it is fetched again. |
| Permissions (advanced) | `Calendars.Read` | Microsoft Graph permissions to ask for. |

`{{time}}` is `09:00–10:00`, `All day`, `From 22:00` or `Until 06:00`.
`{{link}}` opens the event in Outlook on the web, and `{{join}}` is its online
meeting link. Placeholders with nothing to fill in disappear, with the spaces
around them.

### Calendars

The panel shows your default calendar until you choose others under
**Calendars** in the settings, which lists your own calendars and the ones
shared with you, each with its Outlook colour. Events show a stripe in their
calendar's colour, and with more than one calendar chosen, the calendar's
name. A calendar that cannot be read is named above the events, and the
others still show.

Calendars shared with you by someone else may need the **Calendars.Read.Shared**
permission: add it to the app registration and to **Permissions** under
Advanced, then sign in again.

## Development

Plain JavaScript with no dependencies. The source is `src/plugin.js`;
`npm run build` checks it and writes `main.js`, the file Obsidian loads.

```bash
./install.sh "/path/to/vault/.obsidian"   # builds, then copies
npm test
```

The tests load `src/plugin.js` against a small stand-in for the Obsidian API.
They cover startup cost (no request and no secret read on load, no fetch
before the layout is ready or while the panel is hidden), sign-in, token
refresh, day windows across summer time changes, all-day and multi-day
events, the cache, following the daily note, inserting events, matching note
lines to events, the editor icons and showing an event from a note.

To release, bump `version` in `manifest.json` and `package.json`, add it to
`versions.json` and commit. Push the tag first and the branch once the release
exists:

```bash
git tag 1.0.0 && git push origin 1.0.0   # the workflow builds and publishes
git push origin main
```
