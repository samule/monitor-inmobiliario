# Monitor inmobiliario de CABA

Twelve charts for construction costs, rental supply, rents, property prices,
salary access, rental returns, property deeds, UVA mortgage payments, salaries and
the basic basket in dollars, and mortgage credit in Buenos Aires City.

The site is a static report. It contains public source data from IDECBA,
Argentina's labour authority, the Colegio de Escribanos de la Ciudad de Buenos Aires,
the central bank (BCRA), INDEC, the St. Louis Fed (FRED) and ArgentinaDatos
(an unofficial exchange rate API). The report shows the source dates and calculation methods.

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

Charts 9 to 12 were added to compare the monitor with a real-estate video (August 2025):

- **9 · Property deeds.** Monthly sale deeds in the City and the average deed amount in USD.
  The primary source is the monthly report of the Colegio de Escribanos, read from its public
  WordPress API. IDECBA publishes the same series and fills months without a report.
  A year of IDECBA data is discarded when its median differs from the Colegio by more than 5%
  (`deeds_tolerance`). In November 2026 this discards 2026, where IDECBA shows about a fifth
  of the Colegio figures.
- **10 · UVA payment vs rent.** A UVA loan payment is fixed in UVA, so in pesos it follows the
  UVA (BCRA series 31). The chart compares it with the 2-room rent and RIPTE from April 2016.
  Other views estimate the first payment of a new loan each quarter. The loan settings
  (`mortgage_rate_pct`, `mortgage_years`, `mortgage_ltv`) are assumptions in `config/settings.json`,
  not a bank offer.
- **11 · Salary and basic basket in USD.** RIPTE and INDEC's basic total basket per adult
  equivalent, divided by the monthly average MEP or blue dollar.
- **12 · Mortgage credit.** Deeds with a mortgage, their share of sales, and the national
  mortgage loan balance (BCRA series 112) in USD and in constant pesos.

Chart 7 also shows returns for 1, 2 and 3 rooms with the US 10-year Treasury yield and the
assumed UVA rate as references. Chart 8 adds a blue-dollar view that starts in 2015.
There is no public official series for the number of properties listed for sale,
so the monitor has no stock-of-listings chart.
