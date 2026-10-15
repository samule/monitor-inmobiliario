# Monitor inmobiliario de CABA

Seven charts for construction costs, rental supply, rents, property prices,
salary access and estimated rental returns in Buenos Aires City.

The site is a static report. It contains public source data from IDECBA,
Argentina's labour authority and ArgentinaDatos (an unofficial exchange rate API).
The report shows the source dates and calculation methods.

Dates, chart filters and CSV downloads work in the browser. Opening the site does
not get new source data.

GitHub Actions runs the Python updater each Monday at 09:15 in Sao Paulo
(12:15 UTC). The job downloads the sources, checks the data, generates the report,
saves originals and run records in Git, and publishes the result with GitHub Pages.
The computer that created the project does not need to be on.

Use **Actions > Update and publish report > Run workflow** to start an extra update.
The job uses GitHub's temporary built-in token. It does not use a personal token.
If a source fails or data validation fails, the job fails and keeps the last published site.
The Actions run contains the error and a diagnostics artifact.

Source series remain monthly or quarterly. A weekly check does not create weekly observations.
GitHub can delay scheduled jobs. Check the generation date in the report and the Actions history.

For a local run with Python 3.12:

```sh
python -m pip install -r requirements.txt
python scripts/update.py --refresh
python scripts/build_site.py
```

Only `_site` is deployed. Python code and configuration remain in the repository.

The rental return chart is an estimate. Its method is stated next to the chart.

Construction cost per m² (chart 8) uses the official IDECBA Total Ciudad in ARS divided by the same month’s average seller MEP rate. Views show the total, four housing models, monthly USD changes, and nominal ARS. Only closed months with at least ten exchange-rate observations are included. Land, VAT, permits, financing and developer profit are excluded. This is a construction cost for apartment-building models, not a sale price or a quote for an individual house.
