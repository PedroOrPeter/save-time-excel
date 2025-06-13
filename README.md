# Save-Time Excel

This repository provides a simple and effective solution to help your team filter `.xlsx` files efficiently.  
It leverages **Google Apps Script**, along with **HTML** and **CSS**, to create a lightweight, web-based interface for filtering Excel data directly from your Google Drive.

## 🚀 What is Google Apps Script?

According to Google:

> *"Apps Script is a cloud-based JavaScript platform powered by Google Drive that lets you integrate with and automate tasks across Google products."*

In this project, we use Apps Script to work directly with Excel files stored in your Google Drive.  

Whether you want to integrate this tool into your workflow, understand the logic behind it, or explore how it was built — you're in the right place!

## 💡 Why Use This?

- No need for complex Excel formulas.
- Simple and user-friendly web interface.
- Easily customizable for different filtering needs.
- Fast deployment inside your Google Workspace.

## 📁 Tech Stack

- **Google Apps Script** – for backend logic and integration with Google Sheets.
- **HTML + CSS** – for building the filtering form and UI.

## 🔍 The Logic Behind It

Before jumping into code, we need to understand the real problem — and why it matters.

Filtering data in large spreadsheets can be:
- **Confusing** for non-technical team members.
- **Time-consuming**, especially when dealing with complex formulas.
- **A productivity bottleneck**, which ultimately means **lost time, money, and focus**.

If your team frequently works with large `.xlsx` files, applying filters manually becomes inefficient.  
This tool simplifies the filtering process through a guided form, allowing users to focus on what matters most — the **results**.

Picture your sheet as a grid, with **rows** (X) and **columns** (Y).  
Each column has a header, which we can use as its index — this lets us identify and retrieve the values we need with mathematical precision.

But how do we implement this?

Simple — first we create two constants that will enable us to access our data:

```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheetProducts = ss.getSheetByName("Products");
```

getDataRange and getValues will save the values in our fullData const. 
The values from headers will be our first line
Products = fullData.slice(1) will be everything from second line to max number of lines

```javascript
const fullData = sheetProducts.getDataRange().getValues();
const headers = fullData[0];
const products = fullData.slice(1);
```

---

Here:
- getDataRange() grabs the range of cells with data.
- getValues() converts them into a 2D array.

The first row (fullData[0]) contains the column headers.
The rest (fullData.slice(1)) contains the actual data.

We have now:
>- The headers for each column.
>- All the data from the main sheet (Products).

We need to enable users to filter this data. We have a few options:
>- minimum price
>- maximum price
>- color
>- size
>- gender

Instead of adding another sheet with filter criteria, we chose Option 2 — a custom HTML form:

```javascript
const filters = getData(data);
const columns = readSheet(headers);
const filtered = apply(products, columns, filters);
```
Here’s getData — a function from our script:

```javascript
// Retrieve filter values from the HTML form
function getData(data) {
  const minPrice = parseFloat(data.minPrice) || 0;
  const maxPrice = parseFloat(data.maxPrice) || Number.MAX_VALUE;
  const color = (data.color || "").toString().toLowerCase();
  const size = (data.size || "").toString().toLowerCase();
  const gender = (data.gender || "").toString().toLowerCase();

  return { minPrice, maxPrice, color, size, gender };
}
```

Now readSheet(headers) will retrieve the index from each of our Headers:
```javascript
function readSheet(headers) {
  const idPrice = headers.indexOf("SALE_PRICE");
  const idColor = headers.indexOf("COLOR");
  const idSize = headers.indexOf("SIZE");
  const idGender = headers.indexOf("GENDER");

  return { idPrice, idColor, idSize, idGender };
}
```

The apply function expects 3 parameters:
- Products
- Columns
- Filters

 It then applies filter criteria (min and max price, color, size, and gender) and returns the matching products
 ```javascript
// Apply filters from the user's filter form. This function contains filtering logic, but it's easy to follow.
function apply(products, columns, filters) {
  return products.filter(prod => {
    const price = parseFloat(prod[columns.idPrice]);
    const productColor = prod[columns.idColor].toString().toLowerCase();
    const productSize = prod[columns.idSize].toString().toLowerCase();
    const productGender = prod[columns.idGender].toString().toLowerCase();

    return (
      price >= filters.minPrice &&
      price <= filters.maxPrice &&
      (filters.color === "" || productColor === filters.color) &&
      (filters.size === "" || productSize === filters.size) &&
      (filters.gender === "" || productGender === filters.gender)
    );
  });
}
```
Now we create the "Results" sheet and populate it with the filter results
```javascript
 const lastSheet = ss.getSheetByName("Results");
  if (lastSheet) ss.deleteSheet(lastSheet);

  const resultsSheet = ss.insertSheet("Results");
  resultsSheet.appendRow(headers);
  filtered.forEach(row => resultsSheet.appendRow(row)); 

  SpreadsheetApp.getUi().alert("Done! You can see the results in you sheet 'Results'!");
```

## Extra!

To create a interface, use this!
```javascript
// This function builds a UI to make it easier for the user to filter products
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Products Filter')
    .addItem('Open interface', 'showFullScreenInterface')
    .addToUi();
}

// This function opens the interface in a modal dialog with almost full-screen dimensions
function showFullScreenInterface() {
  var html = HtmlService.createHtmlOutputFromFile('interface.html')
      .setWidth(1000)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Product Filter');
}
```


## 👨‍💻 How to Use

Coming soon...

---
