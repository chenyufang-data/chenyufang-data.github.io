# cyfang.org

Chenyu Fang's homepage: one static page, no build step, no tracking.

| File | What it is |
|---|---|
| `index.html` | the page (light and dark mode, phone layout, link-preview tags) |
| `og.png` | the 1200x630 card LinkedIn, Slack and email show when the link is pasted |
| `favicon.svg` | the "CF" tab icon |
| `resume/` | web copies of the three résumés, phone number removed |
| `build_resumes.py` | rebuilds `resume/` from `Documents/JobNavigator/chenyufang/Chenyu_Fang_resume_{MLE,AIE,ASci}.docx` (needs Word) |
| `CNAME` | tells GitHub Pages the site is cyfang.org |

## Update

- Text: edit `index.html`. Every number on the page comes from the résumés; keep them in step.
- Résumés: after changing a .docx, run `python build_resumes.py`, then commit and push.

## Deploy (GitHub Pages, free)

1. On GitHub, create a public repository named `chenyufang-data.github.io` (empty: no README).
2. Push this folder to it:
   ```
   git init -b main
   git add .
   git commit -m "Homepage"
   git remote add origin https://github.com/chenyufang-data/chenyufang-data.github.io.git
   git push -u origin main
   ```
3. Repository **Settings > Pages**: Source "Deploy from a branch", branch `main`, folder `/ (root)`.
   Custom domain: `cyfang.org` (the CNAME file sets it).
4. GoDaddy, **My Products > cyfang.org > DNS** (the name servers are GoDaddy's):
   - If the domain is connected to GoDaddy Website Builder, disconnect it there first (and cancel that plan if
     you pay for it), so its two A records (76.223.105.230 and 13.248.243.5) can be removed.
   - Records for `@`: delete those two A records, add four A records:
     `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`.
   - `www`: a CNAME to `chenyufang-data.github.io`.
   - **Leave the `gridcast` and `coach` records alone**: they run the live demos.
5. Back in **Settings > Pages**, wait for the DNS check to pass (minutes to an hour), then tick **Enforce HTTPS**.
6. Recommended: GitHub **profile Settings > Pages > Add a domain** to verify cyfang.org (a TXT record at
   GoDaddy), so nobody else can claim it on GitHub Pages.

Check: https://cyfang.org shows the page, https://www.cyfang.org redirects to it, and pasting the link into a
LinkedIn message shows the card.
