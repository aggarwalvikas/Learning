# Outlook Subject Viewer

This is a classic Outlook add-in-only task pane add-in. When opened for a selected message, it reads the message sender and subject through Office.js and displays them in the pane.

## Requirements

- Outlook 2016 for Windows with a supported update level
- A mailbox that allows custom Outlook add-ins
- HTTPS hosting for the add-in files
- A trusted development certificate when using localhost

The manifest requests Mailbox requirement set 1.3. Outlook 2016 builds and webview components vary, so verify the add-in on the exact Outlook installation that will use it.

## Local hosting

Serve this folder from an HTTPS web server on port 3000 so these URLs are available:

- `https://localhost:3000/src/taskpane.html`
- `https://localhost:3000/assets/icon-80.png`

If the add-in is hosted elsewhere, replace every `https://localhost:3000` value in `manifest.xml` with the HTTPS origin of that server. The certificate must be trusted by Windows and Outlook.

For a quick static server, use any HTTPS-capable local web server that can serve this directory. This project intentionally has no npm build step because it uses plain HTML, CSS, and ES5-compatible JavaScript.

## Sideload into Outlook

1. Start the HTTPS server.
2. Open Outlook on the web and go to the Add-Ins for Outlook dialog, or in classic Outlook choose **File > Info > Manage Add-ins**.
3. Choose **My add-ins**, then **Add a custom add-in > Add from File**.
4. Select `manifest.xml` and accept the installation prompts.
5. Open a received message in classic Outlook 2016.
6. Use **Show Subject** on the message ribbon to open the task pane.

Classic Outlook may cache manually sideloaded add-ins. Restart Outlook if the add-in does not appear immediately.

## Test cases

Verify the task pane with:

- A message with a normal subject
- A message with a visible sender name and email address
- A message with an empty subject
- A long subject that wraps in the pane
- Subject text containing HTML-like characters or non-Latin characters
- Switching between messages and reopening the pane
- A state where no readable message is available

The add-in only reads the current item. It does not change the message, access the message body, or send data to a server.
