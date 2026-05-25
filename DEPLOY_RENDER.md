# Deploying to Render

Steps to prepare a Render account and deploy this repository.

1. Create a Render account
   - Visit https://render.com and sign up using GitHub.

2. Connect the repository
   - In Render, choose "New" → "Web Service" and connect to this GitHub repo.
   - When prompted, you can let Render use the `render.yaml` in the repository to create the web service and the managed Postgres database automatically.

3. Environment variables
   - Copy values from your local `.env` into the Render service's Environment settings.
   - Required keys (from `.env.example`):
     - `SECRET_KEY`
     - `FERNET_KEY`
     - `DOWNLOAD_SECRET`
     - `DATABASE_URL` (if you already provisioned a DB)
     - `CLOUDINARY_CLOUD_NAME`
     - `CLOUDINARY_API_KEY`
     - `CLOUDINARY_API_SECRET`

4. Create a Postgres database on Render (if not using `render.yaml` automated step)
   - In Render dashboard, choose "New" → "Postgres" to create a managed Postgres instance.
   - After creation, copy the `DATABASE_URL` Render provides and paste it into the Web Service's Environment variables.

5. Deploy
   - Trigger a manual deploy from the Render dashboard or push to the connected Git branch.
   - Logs are available in the service page; ensure the service starts and connects to the DB.

Notes & security
   - Do NOT commit a real `.env` file to git. Use Render's environment settings and secrets.
   - `FERNET_KEY` should be a base64 urlsafe key from `Fernet.generate_key()`.
   - `SECRET_KEY` and `DOWNLOAD_SECRET` must be long, random secrets.
   - Cloudinary credentials are required for file storage (we store encrypted blobs in Cloudinary `raw` resources).

If you want, I can scaffold an `alembic` migration next so the database schema can be applied in production. Would you like that? 
