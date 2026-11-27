# Refreshing dashboard data

The dashboard loads `dashboard-data.json` beside `index.html`; end users do not
upload files. Each month, regenerate the JSON from the updated workbook and
publish both files together:

```powershell
python -m pip install -r requirements.txt
python build_dashboard_data.py "C:\path\to\updated-workbook.xlsx"
```

The converter checks every workbook sheet and ignores `NA`, `N/A`, and `Blank`
values rather than treating them as zero. Country rows are grouped by region;
when a region has an `All` summary row, the summary is used instead of adding
the country rows a second time. Country rows are also retained for the Country
filter. The converter normalizes the known `Beligum` spelling to `Belgium`.
It imports populated database, consent, New Subscribers, New UnSubscribers,
marketable, campaign, form, eStore, Engaged Contacts, Touched Contacts, and
audience data. The dashboard status tooltip identifies sheets
with repeated placeholder values
and sheets with no numeric data. Product columns include `Blank` when the
workbook provides it. Campaign detail includes total and unique clicks when
provided by the source sheet; open counts remain unavailable when the source
sheet omits its unique-open column. Review converter warnings for rows with
missing keys and the dashboard data-status note after each refresh.

The dashboard's Country filter uses country-specific rows where available and
preserves region totals for the `All Countries` selection. It applies to
country-bearing metrics, campaign and form details, eStore, engagement, and
audience dimensions.

FY26 data for January through September is temporarily hidden across the
dashboard. The source JSON remains unchanged; the display restriction is
configured by `TEMPORARILY_HIDDEN_PERIODS` in `index.html` and can be removed
when those periods should be shown again.

Serve the dashboard from a web server; opening `index.html` directly from disk
does not allow the browser to fetch its adjacent JSON file. Keep the hosting
location access-controlled when the workbook contains sensitive data. The
dashboard revalidates the JSON on each visit so monthly updates stay current
without forcing browsers to discard their cached copy. For faster repeat loads,
enable ETag or Last-Modified validation and gzip/Brotli compression on the host.
