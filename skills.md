# Project Skills

This document lists the development skills and technologies used by the Outlook Subject Viewer add-in.

## Outlook Add-in Development

- Build classic Outlook add-ins with the Office add-in-only XML manifest.
- Configure Outlook message read-mode activation with `MessageReadCommandSurface`.
- Target Outlook Mailbox requirement set 1.3 for Outlook 2016 compatibility.
- Define task-pane commands, icons, permissions, and HTTPS resource URLs.

## Office.js Mailbox API

- Initialize the add-in with `Office.onReady`.
- Read the selected Outlook message through `Office.context.mailbox.item`.
- Display the message subject using `item.subject`.
- Display sender information using `item.from.displayName` and `item.from.emailAddress`.
- Handle missing messages, empty subjects, and unavailable sender details safely.

## Frontend Development

- Create accessible task-pane markup with HTML.
- Use semantic labels and live regions for message details and status updates.
- Style the pane with CSS for the Outlook desktop webview.
- Support long subjects, long sender values, and wrapped text without layout overflow.

## Outlook 2016 Compatibility

- Use ES5-compatible JavaScript syntax and browser APIs.
- Avoid framework and bundler dependencies for this small add-in.
- Support older Outlook webview behavior with a fallback to `Office.initialize`.
- Use `textContent` when rendering message values to avoid interpreting message data as HTML.

## HTTPS Local Development

- Serve add-in resources over HTTPS on `https://localhost:3000`.
- Use a trusted Office add-in development certificate.
- Verify that the manifest, task pane, JavaScript, CSS, and icon resources are reachable over HTTPS.
- Use cache-busting query strings when Outlook has cached an older task-pane version.

## Manifest Validation and Testing

- Validate `manifest.xml` with the Office add-in manifest validator.
- Check JavaScript syntax with Node.js.
- Check source files with VS Code diagnostics.
- Sideload the manifest into Outlook on the web or classic Outlook for Windows.
- Test normal, empty, long, and special-character subjects and sender values.

## Source Control

- Use Git to track the add-in source, manifest, documentation, and image assets.
- Keep commits focused on a single project change.
- Push the local `master` branch to the configured GitHub remote.
