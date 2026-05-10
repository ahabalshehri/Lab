# Laboratory And Blood Bank Dashboard

Static GitHub Pages dashboard for laboratory workflow monitoring, weekly reports, hospital comparison, and TAT indicators.

## Local Preview

Open `index.html` directly in a browser, or serve this folder with any static file server.

## GitHub Pages

Upload these files to the `ahabalshehri/Lab` repository:

- `index.html`
- `css/style.css`
- `js/app.js`
- `README.md`

Then enable GitHub Pages:

1. Open the repository on GitHub.
2. Go to `Settings`.
3. Open `Pages`.
4. Choose `Deploy from a branch`.
5. Select branch `main` and folder `/root`.
6. Save.

The public link should become:

```text
https://ahabalshehri.github.io/Lab/
```

## Excel Columns

Recommended columns:

```text
hospital
sample_id
test_name
order_time
collection_time
lab_received_time
result_time
department
priority
```

Arabic or alternative column names are partly supported by the import script.
