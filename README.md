# School Fund Planner

Local Windows desktop app for preparing a school budget from scratch:

1. Enter the school details, fiscal year, annual NEP allocation, and MOOE categories.
2. Add planned SOB items by month, quantity, unit, and unit cost.
3. Review the quarterly APP generated from the monthly SOB entries.

Budget data is saved on this computer in a SQLite database, `fund-planner.db`, in the app data folder (`%APPDATA%\school-fund-planner\` on Windows). Allocation and APP exports are saved as new `.xlsx` files.

## Run the desktop app

```sh
npm install
npm run dev
```

## Build

```sh
npm run build
```

## School list and logos

School names and DepEd School IDs come from SwiSS (`src/renderer/schools.js`). Picking a school fills in its School ID.
Logos live in `src/image/school_logo/` (file names as in the Supabase `school_logo` bucket, e.g. `ieces.png`). A school without a logo file shows the plain dot instead.
