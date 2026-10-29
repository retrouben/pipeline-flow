# Pipeline Flow – Tableau Dashboard Extension

Animated Segment → Category → Region Sankey for the *Pipeline Command Center* dashboard.
Flow width = sales, colour = profit margin (red < 0% → amber → teal ≥ 20%). Three insight chips
(biggest flow, margin drag, concentration) are computed live from the filtered data, and clicking a
flow selects the matching marks on the native sheets so the dashboard's filter actions cross-filter.

The extension reads the `Flow Data` worksheet (Segment, Category, Sales Region, SUM(Sales), SUM(Profit))
via the Tableau Extensions API. Opened outside Tableau it renders demo data.

Hosted via GitHub Pages; point the `.trex` / dashboard zone at `https://<user>.github.io/tableau-pipeline-flow/index.html`.
