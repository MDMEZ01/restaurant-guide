# Maintenance tools

- `heights.py <xlsx>` — refits row heights in every tab (city band rows keep 24 pt).
- `extract.py <xlsx> <data.json>` — builds the site data from the workbook and lists any hours it could not parse.

After `extract.py`, paste data.json into `index.html` between `<script id="data" type="application/json">` and the next `</script>`, replacing `</` with `<\/`.
