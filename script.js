// Main function called from interface.html
// It calls other functions and then displays the results
// by creating a new sheet with the filter results
function filterProducts(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProducts = ss.getSheetByName("Products");

  const fullData = sheetProducts.getDataRange().getValues();
  const headers = fullData[0];
  const products = fullData.slice(1);

  const filters = getData(data);
  const columns = readSheet(headers);

  const filtered = apply(products, columns, filters);

  const lastSheet = ss.getSheetByName("Results");
  if (lastSheet) ss.deleteSheet(lastSheet);

  const resultsSheet = ss.insertSheet("Results");
  resultsSheet.appendRow(headers);
  filtered.forEach(row => resultsSheet.appendRow(row)); 

  SpreadsheetApp.getUi().alert("Done! You can see the results in you sheet 'Results'!");
}

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

// Retrieve filter values from the HTML form
function getData(data) {
  const minPrice = parseFloat(data.minPrice) || 0;
  const maxPrice = parseFloat(data.maxPrice) || Number.MAX_VALUE;
  const color = (data.color || "").toString().toLowerCase();
  const size = (data.size || "").toString().toLowerCase();
  const gender = (data.gender || "").toString().toLowerCase();

  return { minPrice, maxPrice, color, size, gender };
}

// Read headers and return their index positions
function readSheet(headers) {
  const idPrice = headers.indexOf("SALE_PRICE");
  const idColor = headers.indexOf("COLOR");
  const idSize = headers.indexOf("SIZE");
  const idGender = headers.indexOf("GENDER");

  return { idPrice, idColor, idSize, idGender };
}

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
