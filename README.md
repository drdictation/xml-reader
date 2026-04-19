# Genie XML Reader

Read-only browser viewer for Genie `patient_summary` XML exports.

## Local development

This project expects Node `22.x`.

If you use `nvm`:

```bash
nvm install
nvm use
```

Or install Node 22 from [nodejs.org](https://nodejs.org/) and confirm:

```bash
node -v
npm -v
```

```bash
npm install
npm run dev
```

Then open `http://localhost:3000` and load one XML file from the `Example XML Files` folder.

## Future appointment export

Generate a calendar and CSV of future appointments found in a folder of Genie XML files.

```bash
npm run future-appointments
```

Defaults:

- Scans `Example XML Files`
- Treats `2026-03-21` as the cutoff date
- Writes `output/future-appointments.ics`
- Writes `output/future-appointments.csv`

You can also pass a different folder and cutoff date:

```bash
npm run future-appointments -- "/path/to/xml/folder" 2026-03-21
```

The exporter reads `appt_list` records, ignores cancelled or DNA appointments, and uses `apptduration` as the appointment length in seconds.

## Version 1 scope

- Browser-only parsing
- One XML file at a time
- Read-only display
- Structured sections plus raw XML fallback
- No Supabase storage
