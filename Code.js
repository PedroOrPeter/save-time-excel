/**
 * Main entry point for the web app. Returns the HTML content from 'Index' file.
 * @return {HtmlOutput} The HTML output for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index');  
}

/**
 * Retrieves product data from the "Products" sheet.
 * @return {Object} An object containing headers and products data.
 * @throws {Error} If the "Products" sheet doesn't exist.
 */
function getProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Products");

  if (!sheet) {
    throw new Error("Base sheet does not exist.");
  }

  var fullData = sheet.getDataRange().getValues();
  var headers = fullData[0];
  var products = fullData.slice(1);

  return {
    headers: headers,
    products: products
  };
}

/**
 * Filters a list of products based on specified criteria.
 * @return {Array} The filtered array of products.
 */
function filterProductsList(products, headers, filter) {
  // Get column indices
  var idxPrice = headers.indexOf("SALE_PRICE");
  var idxColor = headers.indexOf("COLOR");
  var idxSize = headers.indexOf("SIZE");
  var idxGender = headers.indexOf("GENDER");

  // Parse filter values with defaults
  var minPrice = parseFloat(filter.minPrice) || 0;
  var maxPrice = parseFloat(filter.maxPrice) || Number.MAX_VALUE;
  var color = (filter.color || '').toString().toLowerCase().trim();
  var size = (filter.size || '').toString().toLowerCase().trim();
  var gender = (filter.gender || '').toString().toLowerCase().trim();

  // Filter products based on criteria
  return products.filter(function (prod) {
    var price = parseFloat(prod[idxPrice]) || 0;
    var productColor = prod[idxColor].toString().toLowerCase().trim();
    var productSize = prod[idxSize].toString().toLowerCase().trim();
    var productGender = prod[idxGender].toString().toLowerCase().trim();

    return (
      price >= minPrice &&
      price <= maxPrice &&
      (color === '' || productColor === color) &&
      (size === '' || productSize === size) &&
      (gender === '' || productGender === gender)
    );
  });
}

/**
 * Creates a new sheet with filtered products data.
 * @return {string} The name of the newly created sheet.
 */
function createFilteredSheet(filteredProducts, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create sheet with timestamp in name
  var newSheetName = "Filtered " + new Date().toLocaleString("en-US");
  var newSheet = ss.insertSheet(newSheetName);
  
  // Add headers and data
  newSheet.appendRow(headers);
  filteredProducts.forEach(function (item) {
    newSheet.appendRow(item);
  });

  return newSheetName;
}

/**
 * Main function to filter products and create a new sheet with results.
 * @param {Object} data - The filter criteria (same as filterProductsList).
 * @return {Object} An object containing result information.
 */
function filterProducts(data) {
  // Get all products data
  var result = getProducts();
  var headers = result.headers;
  var products = result.products;

  // Filter products and create new sheet
  var filtered = filterProductsList(products, headers, data);
  var newSheetName = createFilteredSheet(filtered, headers);

  return {
    message: "Data filtered successfully!",
    newSheetName: newSheetName,
    products: filtered,
    headers: headers
  };
}
