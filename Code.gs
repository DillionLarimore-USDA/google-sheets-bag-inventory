// Google Apps Script for Bag Approval and Inventory Management

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bag Management')
    .addItem('Create Bag', 'createBag')
    .addItem('View Bags Log', 'viewBagsLog')
    .addToUi();
}

function createBag() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = ss.getActiveSheet();
    
    // Check if sheet name is "Inventory" or similar (adjust if needed)
    if (!inventorySheet.getName().toLowerCase().includes('inventory')) {
      SpreadsheetApp.getUi().alert('Please make sure you are on the Inventory sheet');
      return;
    }

    // Get all data from inventory sheet
    const data = inventorySheet.getDataRange().getValues();
    
    if (data.length < 2) {
      SpreadsheetApp.getUi().alert('No inventory data found');
      return;
    }

    // Extract headers and find column indices
    const headers = data[0];
    const itemIndex = headers.findIndex(h => h.toLowerCase().includes('item'));
    const quantityIndex = headers.findIndex(h => h.toLowerCase().includes('quantity'));
    const priceIndex = headers.findIndex(h => h.toLowerCase().includes('price'));

    if (itemIndex === -1 || quantityIndex === -1 || priceIndex === -1) {
      SpreadsheetApp.getUi().alert('Could not find Item, Quantity, or Price columns');
      return;
    }

    // Create bag contents and calculate total price
    const bagContents = [];
    let totalPrice = 0;
    const timestamp = new Date();

    for (let i = 1; i < data.length; i++) {
      const item = data[i][itemIndex];
      const quantity = data[i][quantityIndex];
      const price = data[i][priceIndex];

      if (item && quantity > 0) {
        bagContents.push({
          item: item,
          quantity: 1, // Bagging 1 of each item
          price: price,
          lineTotal: 1 * price
        });
        totalPrice += (1 * price);
      }
    }

    if (bagContents.length === 0) {
      SpreadsheetApp.getUi().alert('No items with quantity > 0 to bag');
      return;
    }

    // Update quantities in inventory (decrease by 1)
    for (let i = 1; i < data.length; i++) {
      const quantity = data[i][quantityIndex];
      if (quantity > 0) {
        inventorySheet.getRange(i + 1, quantityIndex + 1).setValue(quantity - 1);
      }
    }

    // Create or get Bags log sheet
    let bagsSheet = ss.getSheetByName('Bags Log');
    if (!bagsSheet) {
      bagsSheet = ss.insertSheet('Bags Log');
      bagsSheet.appendRow(['Bag #', 'Date', 'Items Count', 'Total Price', 'Details']);
    }

    // Add bag record
    const bagNumber = bagsSheet.getLastRow(); // Simple bag numbering
    const details = bagContents.map(b => `${b.item} (Qty: ${b.quantity} @ $${b.price})`).join('; ');
    
    bagsSheet.appendRow([
      bagNumber,
      timestamp,
      bagContents.length,
      totalPrice,
      details
    ]);

    // Show confirmation
    SpreadsheetApp.getUi().alert(
      `✓ Bag Created!\n\n` +
      `Items: ${bagContents.length}\n` +
      `Total Price: $${totalPrice.toFixed(2)}\n\n` +
      `Quantities decreased by 1 for each item.`
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

function viewBagsLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bagsSheet = ss.getSheetByName('Bags Log');
  
  if (!bagsSheet) {
    SpreadsheetApp.getUi().alert('No bags have been created yet');
    return;
  }
  
  ss.setActiveSheet(bagsSheet);
  SpreadsheetApp.getUi().alert('Switched to Bags Log sheet');
}
