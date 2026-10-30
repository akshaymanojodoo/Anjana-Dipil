# Wedding invitation with Google Sheets guestbook

## 1. Google Sheet + Apps Script (about 3 minutes)
1. Create a new Google Sheet (name it e.g. "Anjana Dipil Wishes").
2. Extensions > Apps Script. Delete the sample code and paste `Code.gs`. Save.
3. Deploy > New deployment > type **Web app**. Execute as: **Me**. Who has access: **Anyone**. Deploy and authorise.
4. Copy the Web app URL (ends with `/exec`).
5. Paste that URL into `config.js` as `sheetUrl`.
6. Test: open the URL in a browser; you should see `{"wishes":[]}`.

Wishes appear in the "Wishes" tab. To hide one from the site, set its Approved column to FALSE.
If you edit Code.gs later, use Deploy > Manage deployments > Edit > New version (the URL stays the same).

## 2. GitHub Pages
```
git init && git add . && git commit -m "Wedding invitation"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```
Then repo Settings > Pages > Deploy from branch > `main` / root. Site: `https://<user>.github.io/<repo>/`
