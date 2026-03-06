# prove-webextension-wiped-out-by-getFileAsync

Prove that the webextensions folder in a Word document is being wiped out by `getFileAsync` in Office.js.

When a Word document contains a web extension (Office Add-in), calling `Office.context.document.getFileAsync(Office.FileType.Compressed)` returns a copy of the document with the `webextensions` folder stripped out. This repository provides a minimal Office Add-in to reproduce and verify this behavior.

## How It Works

1. The add-in provides a **"Download from getFileAsync()"** button in a task pane.
2. Clicking the button calls `getFileAsync` to retrieve the document as a compressed (`.docx`) file.
3. The file data is uploaded to a local HTTPS server, which serves it back as a download.
4. You can then inspect the downloaded `.docx` (which is a ZIP archive) to confirm that the `webextensions` folder is missing.

## Prerequisites

- [Node.js](https://nodejs.org/)
- Office Add-in dev certificates in `~/.office-addin-dev-certs/` (`localhost.key` and `localhost.crt`). You can generate these using [`office-addin-dev-certs`](https://www.npmjs.com/package/office-addin-dev-certs):
  ```bash
  npx office-addin-dev-certs install
  ```
- Microsoft Word (desktop)

## Getting Started

1. Start the server:
   ```bash
   node server.js
   ```
   The server runs at `https://localhost:3008`.

2. Sideload the add-in in Word using the manifest file `manifest-prove-webextension-wiped-out-by-getFileAsync.xml`.

3. Open a Word document that contains a web extension.

4. Open the add-in task pane and click **"Download from getFileAsync()"**.

5. Open the downloaded `.docx` file as a ZIP archive and verify that the `webextensions` folder is missing.

## Project Structure

| File | Description |
|------|-------------|
| `server.js` | HTTPS server that serves the add-in and handles file upload/download |
| `index.html` | Add-in task pane UI with the getFileAsync button |
| `download.html` | Download page opened in the browser to save the file |
| `manifest-prove-webextension-wiped-out-by-getFileAsync.xml` | Office Add-in manifest for sideloading |
