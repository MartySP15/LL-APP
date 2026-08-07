# Liveable Layouts — Orders & Invoicing

A shared orders/invoicing tool: customers, orders, PDF-ready invoices, and payment
tracking, backed by a real database so your whole team sees the same data — with
live updates when a teammate makes a change.

- **Frontend:** React + Vite, hosted free on GitHub Pages
- **Database:** Supabase (free tier), shared across everyone using the link

## 1. Create your Supabase project

1. Go to [supabase.com](https://supabase.com) and create a free account/project.
2. Once the project is ready, open **SQL Editor → New query**, paste in the contents
   of [`supabase/schema.sql`](./supabase/schema.sql), and run it. This creates the
   `customers`, `orders`, and `business_settings` tables plus auto-numbering for
   orders/invoices.
3. Go to **Project Settings → API**. You'll need two values from there:
   - **Project URL**
   - **anon public** key

> Note on access: the schema uses open read/write policies so anyone with your
> deployed link can use the app — good for a small trusted team. If you later want
> individual client logins, Supabase Auth can be layered on top; ask and I can wire
> that up.

## 2. Run it locally (optional, to test first)

```bash
npm install
cp .env.example .env
# edit .env and paste in your Project URL + anon key
npm run dev
```

Open the printed local URL — you should see the app load with no customers/orders yet.

## 3. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/orders-invoicing-app.git
git push -u origin main
```

If you name your repo something other than `orders-invoicing-app`, update the
`base` path in [`vite.config.js`](./vite.config.js) to match, e.g.
`base: "/your-repo-name/"`.

## 4. Add your Supabase keys as GitHub secrets

In your repo: **Settings → Secrets and variables → Actions → New repository secret**,
and add both:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

(Same values as your `.env` file. These get baked into the build safely — the anon
key is meant to be public-facing, protection comes from the database policies.)

## 5. Turn on GitHub Pages

**Settings → Pages → Build and deployment → Source → GitHub Actions.**

That's it — the included workflow (`.github/workflows/deploy.yml`) builds and
deploys automatically on every push to `main`. After the first push, check the
**Actions** tab for progress; once it's green, your app is live at:

```
https://YOUR_USERNAME.github.io/orders-invoicing-app/
```

Share that link with your team or clients. Anyone who opens it sees the same
customers, orders, and invoices — updates sync live across everyone with the tab open.

## Making future changes

Edit the code, commit, and push to `main` — GitHub Actions rebuilds and redeploys
automatically. No manual steps needed after the first setup.

## Project structure

```
src/
  App.jsx            # entire app: UI + Supabase data layer
  supabaseClient.js   # Supabase connection (reads .env)
  main.jsx            # React entry point
supabase/
  schema.sql           # run once in Supabase's SQL Editor
.github/workflows/
  deploy.yml            # auto-build + deploy to GitHub Pages
```
