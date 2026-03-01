# Power Platform Blog Digest — Power Automate Solution

> **A daily dual-channel digest of new Microsoft Power Platform blog posts — delivered automatically to your inbox *and* your Microsoft Teams channel every morning.**

[![Power Automate](https://img.shields.io/badge/Power%20Automate-Cloud%20Flow-0078d4?logo=microsoftpowerautomate&logoColor=white)](https://make.powerautomate.com)
[![Microsoft Teams](https://img.shields.io/badge/Microsoft%20Teams-Notification-6264a7?logo=microsoftteams&logoColor=white)](https://teams.microsoft.com)
[![Power Platform](https://img.shields.io/badge/Power%20Platform-Solution-742774?logo=microsoftpowerapps&logoColor=white)](https://make.powerapps.com)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## What It Does

Every day at **8:00 AM UTC** this flow:

1. Fetches the [Microsoft Power Platform Blog RSS feed](https://www.microsoft.com/en-us/power-platform/blog/feed/) for posts published in the last 24 hours
2. Builds a **styled HTML email** with a branded header, one card per post (title, date, summary, Read More link), and a footer
3. Sends the digest to a configurable recipient via **Office 365 Outlook**
4. Posts the same digest as an **Adaptive Card** to a configured **Microsoft Teams** chat or channel
5. If no new posts exist, exits gracefully with status **Succeeded** — no empty notifications
6. On any failure, the **Error Handler Scope** sends an alert email to a dedicated admin address

---

## Solution Contents

```
PowerPlatformBlogDigest_1_0_0_3.zip          ← Importable Power Platform solution
├── solution.xml                              ← Solution manifest (name, version, publisher)
├── customizations.xml                        ← Flow registration metadata
├── [Content_Types].xml                       ← OPC content types
├── environmentvariabledefinitions/           ← AlertEmail & NotificationEmail env var schemas
└── Workflows/
    └── PowerPlatformBlogDigest-*.json        ← Full cloud flow definition

Screenshots/                                  ← Article screenshots (referenced by blog-post.html)
blog-post.html                               ← Standalone HTML article with embedded screenshots
blog-post.docx                               ← Word document version with embedded screenshots
generate_docx.py                             ← Python script that generated blog-post.docx
```

---

## Flow Architecture

```
Recurrence Trigger (Daily 8:00 AM UTC)
│
├── Initialize Variable: HTML_Body (string, empty)
│
└── Main Scope
    ├── List all RSS feed items
    ├── Filter Array (last 24 hours)
    │
    └── Condition: length(posts) > 0?
        │
        ├── TRUE ──► Apply to Each Post
        │               ├── Compose Post HTML card
        │               └── Append to HTML_Body
        │           Compose Email Header
        │           Compose Email Footer
        │           Compose Full Email Body  (Header + HTML_Body + Footer)
        │           Compose Final Card       (Teams Adaptive Card payload)
        │           Send Notification Email  (Office 365 Outlook → NotificationEmail)
        │           Post card in a chat or channel  (Microsoft Teams)
        │
        └── FALSE ─► Terminate (Succeeded) — no notifications sent

Error Handler Scope  [runs only when Main Scope fails or times out]
    ├── Compose Error Details (HTML error report with code, message, timestamp)
    └── Send Error Alert Email  (Office 365 Outlook → AlertEmail)
```

---

## Connections Required

| Connector | Purpose |
|-----------|---------|
| **RSS** | Fetches Power Platform blog feed items |
| **Office 365 Outlook** | Sends digest email and error alert emails |
| **Microsoft Teams** | Posts Adaptive Card digest to a chat or channel |

All three connections are created or selected during the solution import wizard.

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `NotificationEmail` | `user@example.com` | Recipient for the daily digest email |
| `AlertEmail` | `admin@example.com` | Admin address for error alert emails |

These are **Power Platform environment variables** — set them in the solution's
**Environment Variables** section after import. No need to edit the flow actions directly.

---

## Prerequisites

- A **Power Platform environment** with a Dataverse database provisioned
- An **Office 365 / Microsoft 365** licence (for the Outlook connector)
- A **Microsoft Teams** licence (for the channel card notification)
- **Maker** (or higher) role in the target environment

---

## How to Import & Use

### 1. Clone or download

```bash
git clone https://github.com/rkneela0912/PowerPlatformLatestUpdates_MicrosoftBlogPosts_Monitor.git
```

Or download `PowerPlatformBlogDigest_1_0_0_3.zip` directly from the repository.

### 2. Import into Power Platform

1. Go to [make.powerautomate.com](https://make.powerautomate.com) and select your environment
2. Navigate to **Solutions → Import solution**
3. Upload `PowerPlatformBlogDigest_1_0_0_3.zip`
4. On the **Connections** step, create or select:
   - An **RSS** connection (no credentials — it's a public feed)
   - An **Office 365 Outlook** connection (sign in with your Microsoft 365 account)
   - A **Microsoft Teams** connection (sign in with your Microsoft 365 account)
5. Click **Import** and wait for the import to complete

### 3. Set environment variable values

1. Open the imported solution in [make.powerapps.com](https://make.powerapps.com)
2. Navigate to **Environment Variables**
3. Set `NotificationEmail` to the recipient address for the daily digest
4. Set `AlertEmail` to your admin / monitoring email address

### 4. Configure the Teams action

1. Open the **Power Platform Blog Digest** flow and click **Edit**
2. Find the **Post card in a chat or channel** action (inside the Main Scope, after Send Notification Email)
3. Select your target **Team** and **Channel** (or switch to a Chat if preferred)
4. **Save** the flow

### 5. Turn on the flow

Click **Turn on** in the flow detail page. The flow will run automatically at the next 8:00 AM UTC.

### 6. Test immediately (optional)

Click **Run → Run flow** to trigger a manual test run and verify both the email and Teams card land as expected.

---

## How It Was Built

### Why I Built This

Staying current with the Microsoft Power Platform blog is part of my daily routine as a Power Platform architect. The blog publishes important announcements, deprecation notices, and best-practice guides — but there is no native "daily digest" subscription. I wanted a zero-maintenance solution that would surface only today's posts, formatted for quick scanning, visible both in my inbox and in my team's Teams channel — without any external services or code deployments.

### Design Principles

**Dual-channel delivery** — The same post data is rendered twice in a single flow run: once as a rich HTML email via Office 365 Outlook, and once as an Adaptive Card posted to Microsoft Teams. Individual recipients get a detailed email for in-depth reading; the team channel gets an at-a-glance card for instant awareness.

**Scope-based error handling** — The entire business logic lives inside `Main_Scope`. A second `Error_Handler_Scope` is wired to run only when `Main_Scope` fails or times out (`runAfter: ["Failed", "TimedOut"]`). This cleanly separates the happy path from the error path — analogous to try/catch.

**Graceful no-op on quiet days** — Instead of sending empty notifications, the `else` branch uses a `Terminate` action with `runStatus: "Succeeded"`. This keeps run history clean (all green) and avoids notification fatigue.

**HTML generated entirely in-flow** — All email styling uses inline CSS inside `Compose` actions. No external templates, no Azure Functions, no storage accounts. The flow is fully self-contained and portable.

**Environment variables over hard-coded values** — `NotificationEmail` and `AlertEmail` are declared as solution-level environment variables. Importers configure them once in the Environment Variables pane without touching flow internals, making the solution plug-and-play across environments.

### Key Expressions Used

| Expression | Where | Purpose |
|------------|-------|---------|
| `addDays(utcNow(), -1)` | RSS filter | Fetch only last-24-hour posts |
| `length(body('List_all_RSS_feed_items'))` | Condition | Check if any posts were returned |
| `formatDateTime(..., 'dddd, MMMM dd yyyy')` | Header & cards | Human-readable date display |
| `items('Apply_to_Each_Post')?['title']` | Post card Compose | Safe null-coalescing access to RSS fields |
| `result('Main_Scope')?[0]?['error']?['code']` | Error handler | Extract error details from failed scope |

---

## Notification Output

### Email Digest (Office 365 Outlook)

The digest email uses Microsoft's brand blue (`#0078d4`) throughout and is fully self-contained HTML — no linked stylesheets, renders reliably in all email clients:

- **Header** — Gradient banner (`#0078d4 → #005a9e`) with today's date
- **Post cards** — White card, 4 px blue left border, title linked, publish date, summary, "Read More" button
- **Footer** — "View Full Blog Archive" link

### Teams Adaptive Card (Microsoft Teams)

An Adaptive Card is posted to the configured chat or channel via the EPMPoint Workflows bot. It presents the same post list — title, publish date, summary, and Read More buttons — giving the team at-a-glance awareness directly in their Teams feed.

---

## Versioning

| Version | Date | Notes |
|---------|------|-------|
| `1.0.0.1` | 2026-02-28 | Initial release |
| `1.0.0.2` | 2026-02-28 | Packaging refinements |
| `1.0.0.3` | 2026-02-28 | Added Microsoft Teams Adaptive Card notification; environment variables; current release |

---

## Author

**Ranjith Neela**
Power Platform Architect — EPMPoint
[iamranjithneela@gmail.com](mailto:iamranjithneela@gmail.com)

---

## License

MIT — free to use, modify, and redistribute with attribution.
