# Cairn GTM deployment

## 1. GitHub

Recommended repository: `cairn-gtm-site`

Upload the contents of this folder at the repository root. The site has no build step.

## 2. Vercel

Import the GitHub repository into Vercel.

- Framework preset: Other
- Root directory: repository root
- Build command: leave blank
- Output directory: leave blank
- Install command: leave blank

Deploy first to the generated `*.vercel.app` preview URL and verify desktop + mobile before changing DNS.

## 3. Domain cutover

After preview QA, add both `cairngtm.com` and `www.cairngtm.com` to the Vercel project. Follow the exact DNS records Vercel gives for the project, then set the preferred primary domain and redirect the other hostname to it.

Do not remove the current site until the Vercel domain check passes.

## 4. Post-launch checks

- Homepage loads on `https://cairngtm.com/`
- Mobile nav opens/closes
- Pipeline Analyzer CTA opens the Vercel analyzer
- Pipeline Calculator CTA opens the GitHub Pages calculator
- Writing cards open the intended Substack posts
- `https://cairngtm.com/robots.txt` loads
- `https://cairngtm.com/sitemap.xml` loads
- Social preview uses `/og-image.png`