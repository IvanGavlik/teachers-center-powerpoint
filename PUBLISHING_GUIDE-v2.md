 Steps to Publish

  1. Prepare Your Accounts

  - Join the Microsoft Partner Network at https://aka.ms/joinmarketplace
  - Sign in to Partner Center at https://partner.microsoft.com/dashboard/home
  - Enroll in the Microsoft 365 and Copilot program

  2. Prepare Your Add-in for Production

  Things to fix in your current project before submitting:

  - SSL required — your Render deployment already uses HTTPS, so you're good
  - Update manifest.xml — make sure ProviderName matches your Partner Center publisher name (currently says "Contoso")
  - Privacy policy, Terms of Use, Support pages — you already have these at /assets/privacy.html, /assets/terms.html, /assets/support.html (all must be
  HTTPS, no 404s)
  - Validate your manifest using Microsoft's https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest

  3. Prepare Store Listing Assets

  You'll need:
  - App name: Teacher Assistant PowerPoint
  - Summary: Short description (1-2 sentences)
  - Description: Detailed description of features
  - Icons: You already have all sizes (16, 32, 64, 80, 128)
  - Screenshots: 2-3 screenshots showing the add-in in action (dialog, preview, settings)
  - Categories: Pick 1-3 (e.g., "Education", "Productivity")
  - Search keywords (optional)

  4. Prepare Test Instructions for Reviewers

  This is critical — apps without clear instructions get auto-rejected. You need:
  - Step-by-step instructions on how to test the add-in
  - Test account/credentials if needed
  - Explain: open ribbon → click button → dialog opens → pick type → enter prompt → preview → insert slides
  - Note that it requires a running backend (provide the Render URL)

  5. Submit via Partner Center

  1. Go to Marketplace offers → Microsoft 365 and Copilot tab
  2. Click + New offer → select Office Add-in
  3. Name your app, select publisher
  4. Upload manifest.xml for package testing
  5. Fill in categories, legal links (privacy, terms, support)
  6. Create store listing (summary, description, screenshots, icons)
  7. Set availability date
  8. Add certification notes (test instructions)
  9. Click Review and publish

  6. Wait for Review

  - Expect 3-4 business days for initial response
  - Full process typically takes 4-6 weeks (often requires resubmissions)
  - You'll get email notifications about status

  ---
  Checklist Before Submitting
  ┌────────────────────────────────────────────┬────────────────────────────────────┐
  │                    Item                    │               Status               │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ ProviderName in manifest matches publisher │ Needs update (currently "Contoso") │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Privacy policy page (HTTPS, no 404)        │ Done                               │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Terms of use page (HTTPS, no 404)          │ Done                               │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Support page (HTTPS, no 404)               │ Done                               │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Icons (16, 32, 64, 80, 128)                │ Done                               │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Backend deployed and running               │ Done (Render)                      │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Screenshots for store listing              │ Need to create                     │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Test instructions for reviewers            │ Need to write                      │
  ├────────────────────────────────────────────┼────────────────────────────────────┤
  │ Partner Center account                     │ Need to set up                     │
  └────────────────────────────────────────────┴────────────────────────────────────┘
  ---
  Sources:
  - https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource
  - https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide
  - https://learn.microsoft.com/en-us/partner-center/marketplace-offers/appsource-submission-faq
  - https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish