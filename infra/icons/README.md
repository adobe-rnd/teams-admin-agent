# Custom app icons (admin bot)

To use your admin bot icon in Teams (e.g. the icon shown when users type `@admin-bot`):

1. Add two PNG files here:
   - **color.png** — 192×192 px (full-color icon for app list and @mention picker)
   - **outline.png** — 32×32 px (outline/monochrome for small contexts)

2. Run the packaging script from the repo root:
   ```bash
   ./infra/package-manifest.sh
   ```

3. Upload the new `dist/teams-admin-agent.zip` in **Teams Admin Center → Manage apps → Upload** (or sideload via **Teams → Apps → Manage your apps → Upload a custom app**).

If these files are missing, the script uses placeholder purple icons instead.
