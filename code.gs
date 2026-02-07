// ============================================
// Html Service
// ============================================
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Inventory Management")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"
    )
    .addMetaTag("apple-mobile-web-app-capable", "yes") // apple-specific
    .addMetaTag("mobile-web-app-capable", "yes"); // old andriod-specific
  //.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// Constants
// ============================================
const ITEM_SHEET_NAME = "Items";
const ACCESSORY_SHEET_NAME = "Accessories";
const ITEM_ACCESSORY_SHEET_NAME = "Item_Accessory_Mapping"; // Many-to-many mapping
const REQUEST_SHEET_NAME = "Requests";
const REQUEST_ITEM_SHEET_NAME = "Request_Item";
const REQUEST_ITEM_ACCESSORY_SHEET_NAME = "Request_Item_Accessory";
const STOCK_LEDGER_SHEET_NAME = "StockLedger";
const USER_SHEET_NAME = "Users";
const REQUEST_ACTIVITY_SHEET_NAME = "Request_Activity";
const SYSTEM_ACTIVITY_SHEET_NAME = "System_Activity";
const INVENTORY_ACTIVITY_SHEET_NAME = "Inventory_Activity";
const SESSION_SHEET_NAME = "Sessions";
const DURATION_CACHE_SEC = 15; // 5 minutes
const ENABLE_CACHE = false; // Set to false to disable caching

// ============================================
// Core Functions
// ============================================

/*
@ Get Authorized User and Permission by Gmail
*/
function getAuthorized(){
    const email = Session.getActiveUser().getEmail();
    const ss = getActiveSheet();
    
    let userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet) {
        throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
    }

    const data = userSheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header row
    const emailColIndex = headers.indexOf("Email");
    const row = data.find(row => row[emailColIndex] === email);

    if (!row) {
        return { perm: 'Guest' };
    }

    return { perm: row[headers.indexOf("Permission") ] || 'Guest' };
}

/*
@ Get Active Spreadsheet
*/
const getActiveSheet = () => {
    return SpreadsheetApp.getActiveSpreadsheet();
}

function getNextId(sheet, offset = 0) {
    const lastRow = sheet.getLastRow();

    if (lastRow < 1) return 1 + offset;

    const range = sheet.getRange(1, 1, lastRow, 1).getValues();
    let lastId = 0;

    for (let i = range.length - 1; i >= 0; i--) {
        if (range[i][0] !== "" && !isNaN(Number(range[i][0]))) {
            lastId = Number(range[i][0]);
            break;
        }
    }

    return lastId + 1 + offset;
}

// ============================================
// Debug Functions
// ============================================
/*
@ [DEBUG]
@ Repair Sheets Structure
*/
const repairSheets = () => {
  const ss = getActiveSheet();

  let userSheet = ss.getSheetByName(USER_SHEET_NAME);
  let itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
  let accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
  let itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
  let requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
  let requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
  let requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);
  let stockLedgerSheet = ss.getSheetByName(STOCK_LEDGER_SHEET_NAME);
  let sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);

    if (!userSheet) {
        userSheet = ss.insertSheet(USER_SHEET_NAME);
        userSheet.getRange(1,1,1,4).setValues([
            ["Email","Password","Permission","Active"]
        ])

        Logger.log(`Created sheet: ${USER_SHEET_NAME}`);
    }

    // Create sheets if not exist
    if (!itemsSheet) {
        itemsSheet = ss.insertSheet(ITEM_SHEET_NAME);
        itemsSheet.getRange(1,1,1,11).setValues([
            ["Item_Id","Item_Name","Item_Desc","Total_Qty","Available_Qty","Image","Active","Created_By","Created_At","Modified_By","Modified_At"]
        ])

        Logger.log(`Created sheet: ${ITEM_SHEET_NAME}`);
    }

    if (!accessorySheet) {
        accessorySheet = ss.insertSheet(ACCESSORY_SHEET_NAME);
        accessorySheet.getRange(1,1,1,10).setValues([
            ["Accessory_Id","Accessory_Name","Accessory_Desc","Total_Qty","Available_Qty","Active","Created_By","Created_At","Modified_By","Modified_At"]
        ])

        Logger.log(`Created sheet: ${ACCESSORY_SHEET_NAME}`);
    }

    // Item-Accessory Many-to-Many mapping
    if (!itemAccessorySheet) {
        itemAccessorySheet = ss.insertSheet(ITEM_ACCESSORY_SHEET_NAME);
        itemAccessorySheet.getRange(1,1,1,6).setValues([
            ["Mapping_Id","Item_Id","Accessory_Id","Created_By","Created_At","Active"]
        ])

        Logger.log(`Created sheet: ${ITEM_ACCESSORY_SHEET_NAME}`);
    }

    if (!requestSheet) {
        requestSheet = ss.insertSheet(REQUEST_SHEET_NAME);
        requestSheet.getRange(1,1,1,11).setValues([
            ["Request_Id","Requirer_Name","Status","Request_Date","Distributed_Date","Return_Date","Remark","Created_By","Created_At","Modified_By","Modified_At"]
        ])

        Logger.log(`Created sheet: ${REQUEST_SHEET_NAME}`);
    }

    if (!requestItemSheet) {
        requestItemSheet = ss.insertSheet(REQUEST_ITEM_SHEET_NAME);
        requestItemSheet.getRange(1,1,1,7).setValues([
            ["Request_Id","Item_Index","Item_Id", "Item_Name","Qty","Returned_Qty","Status"]
        ])

        Logger.log(`Created sheet: ${REQUEST_ITEM_SHEET_NAME}`);
    }

    if(!requestItemAccessorySheet) {
        requestItemAccessorySheet = ss.insertSheet(REQUEST_ITEM_ACCESSORY_SHEET_NAME);
        requestItemAccessorySheet.getRange(1,1,1,8).setValues([
            ["Request_Id","Item_Index","Accessory_Index","Accessory_Id", "Accessory_Name","Qty","Returned_Qty","Status"]
        ])

        Logger.log(`Created sheet: ${REQUEST_ITEM_ACCESSORY_SHEET_NAME}`);
    }

    if (!stockLedgerSheet) {
        stockLedgerSheet = ss.insertSheet(STOCK_LEDGER_SHEET_NAME);
        stockLedgerSheet.getRange(1,1,1,9).setValues([
            ["Ledger_Id","Request_Id","Ref_Id","Ref_Type","Item_Type","Qty_Change","Balance_After","Action_By","Action_At"]
        ])

        Logger.log(`Created sheet: ${STOCK_LEDGER_SHEET_NAME}`);
    }

    // Create 3 separate activity sheets
    let requestActivitySheet = ss.getSheetByName(REQUEST_ACTIVITY_SHEET_NAME);
    if (!requestActivitySheet) {
        requestActivitySheet = ss.insertSheet(REQUEST_ACTIVITY_SHEET_NAME);
        requestActivitySheet.getRange(1,1,1,4).setValues([
            ["Log_Id","Email","Activity","Action_At"]
        ])
        Logger.log(`Created sheet: ${REQUEST_ACTIVITY_SHEET_NAME}`);
    }

    let systemActivitySheet = ss.getSheetByName(SYSTEM_ACTIVITY_SHEET_NAME);
    if (!systemActivitySheet) {
        systemActivitySheet = ss.insertSheet(SYSTEM_ACTIVITY_SHEET_NAME);
        systemActivitySheet.getRange(1,1,1,4).setValues([
            ["Log_Id","Email","Activity","Action_At"]
        ])
        Logger.log(`Created sheet: ${SYSTEM_ACTIVITY_SHEET_NAME}`);
    }

    let inventoryActivitySheet = ss.getSheetByName(INVENTORY_ACTIVITY_SHEET_NAME);
    if (!inventoryActivitySheet) {
        inventoryActivitySheet = ss.insertSheet(INVENTORY_ACTIVITY_SHEET_NAME);
        inventoryActivitySheet.getRange(1,1,1,4).setValues([
            ["Log_Id","Email","Activity","Action_At"]
        ])
        Logger.log(`Created sheet: ${INVENTORY_ACTIVITY_SHEET_NAME}`);
    }

    if (!sessionSheet) {
        sessionSheet = ss.insertSheet(SESSION_SHEET_NAME);
        sessionSheet.getRange(1,1,1,5).setValues([
            ["Session_Id","Email","Permission","Created_At","Last_Activity"]
        ])
        Logger.log(`Created sheet: ${SESSION_SHEET_NAME}`);
    }
}

// ============================================
// Item Functions
// ============================================

/*
@ Check Admin Permission
*/
function checkAdminPermission(email) {
    const ss = getActiveSheet();
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    
    if (!userSheet) {
        return false;
    }
    
    const userData = userSheet.getDataRange().getValues();
    const headers = userData[0];
    const emailCol = headers.indexOf("Email");
    const permissionCol = headers.indexOf("Permission");
    const activeCol = headers.indexOf("Active");
    
    for (let i = 1; i < userData.length; i++) {
        if (userData[i][emailCol] === email) {
            const isActive = userData[i][activeCol] === true || userData[i][activeCol] === 'TRUE';
            return userData[i][permissionCol] === 'Admin' && isActive;
        }
    }
    
    return false;
}

/*
@ Create Item (Admin only)
*/
function createItem(itemData, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
        
        if (!itemsSheet) {
            throw new Error(`Sheet "${ITEM_SHEET_NAME}" not found`);
        }
        
        const itemId = getNextId(itemsSheet, 0);
        const timestamp = new Date();
        
        itemsSheet.appendRow([
            itemId,
            itemData.Item_Name,
            itemData.Item_Desc || '',
            itemData.Total_Qty || 0,
            itemData.Available_Qty || 0,
            itemData.Image || '',
            itemData.Active !== false ? true : false,
            userEmail,
            timestamp,
            userEmail,
            timestamp
        ]);
        
        logInventoryActivity(`Admin ${userEmail} created item: ${itemData.Item_Name} (ID: ${itemId})`);
        
        return { success: true, message: "Item created successfully", itemId: itemId };
    } catch (error) {
        Logger.log("Error in createItem: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update Item (Admin only)
*/
function updateItem(itemId, itemData, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
        
        if (!itemsSheet) {
            throw new Error(`Sheet "${ITEM_SHEET_NAME}" not found`);
        }
        
        const data = itemsSheet.getDataRange().getValues();
        const headers = data[0];
        
        // Find item row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == itemId) {
                const rowNum = i + 1;
                const timestamp = new Date();
                
                // Update fields
                if (itemData.Item_Name !== undefined) itemsSheet.getRange(rowNum, 2).setValue(itemData.Item_Name);
                if (itemData.Item_Desc !== undefined) itemsSheet.getRange(rowNum, 3).setValue(itemData.Item_Desc);
                if (itemData.Total_Qty !== undefined) itemsSheet.getRange(rowNum, 4).setValue(itemData.Total_Qty);
                if (itemData.Available_Qty !== undefined) itemsSheet.getRange(rowNum, 5).setValue(itemData.Available_Qty);
                if (itemData.Image !== undefined) itemsSheet.getRange(rowNum, 6).setValue(itemData.Image);
                if (itemData.Active !== undefined) itemsSheet.getRange(rowNum, 7).setValue(itemData.Active);
                
                itemsSheet.getRange(rowNum, 10).setValue(userEmail);
                itemsSheet.getRange(rowNum, 11).setValue(timestamp);
                
                logInventoryActivity(`Admin ${userEmail} updated item ID: ${itemId}`);
                
                return { success: true, message: "Item updated successfully" };
            }
        }
        
        return { success: false, message: "Item not found" };
    } catch (error) {
        Logger.log("Error in updateItem: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Delete Item (Admin only)
*/
function deleteItem(itemId, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
        
        if (!itemsSheet) {
            throw new Error(`Sheet "${ITEM_SHEET_NAME}" not found`);
        }
        
        const data = itemsSheet.getDataRange().getValues();
        
        // Find and delete item row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == itemId) {
                itemsSheet.deleteRow(i + 1);
                logInventoryActivity(`Admin ${userEmail} deleted item ID: ${itemId}`);
                return { success: true, message: "Item deleted successfully" };
            }
        }
        
        return { success: false, message: "Item not found" };
    } catch (error) {
        Logger.log("Error in deleteItem: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Create Accessory (Admin only) - Now supports many-to-many item mapping
*/
function createAccessory(accessoryData, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
        
        if (!accessorySheet) {
            throw new Error(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
        }
        
        if (!itemAccessorySheet) {
            throw new Error(`Sheet "${ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }
        
        const accessoryId = getNextId(accessorySheet, 0);
        const timestamp = new Date();
        
        // Create accessory (no Item_Id column anymore)
        accessorySheet.appendRow([
            accessoryId,
            accessoryData.Accessory_Name,
            accessoryData.Accessory_Desc || '',
            accessoryData.Total_Qty || 0,
            accessoryData.Available_Qty || 0,
            accessoryData.Active !== false ? true : false,
            userEmail,
            timestamp,
            userEmail,
            timestamp
        ]);
        
        // Create mappings if Item_Ids provided
        if (accessoryData.Item_Ids && Array.isArray(accessoryData.Item_Ids) && accessoryData.Item_Ids.length > 0) {
            accessoryData.Item_Ids.forEach(itemId => {
                const mappingId = getNextId(itemAccessorySheet, 0);
                itemAccessorySheet.appendRow([
                    mappingId,
                    itemId,
                    accessoryId,
                    userEmail,
                    timestamp,
                    true
                ]);
            });
        }
        
        logInventoryActivity(`Admin ${userEmail} created accessory: ${accessoryData.Accessory_Name} (ID: ${accessoryId})`);
        
        return { success: true, message: "Accessory created successfully", accessoryId: accessoryId };
    } catch (error) {
        Logger.log("Error in createAccessory: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update Accessory (Admin only) - Now supports updating item mappings
*/
function updateAccessory(accessoryId, accessoryData, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
        
        if (!accessorySheet) {
            throw new Error(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
        }
        
        if (!itemAccessorySheet) {
            throw new Error(`Sheet "${ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }
        
        const data = accessorySheet.getDataRange().getValues();
        const headers = data[0];
        
        // Find accessory row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == accessoryId) {
                const rowNum = i + 1;
                const timestamp = new Date();
                
                // Update fields (note: column indices changed after removing Item_Id)
                if (accessoryData.Accessory_Name !== undefined) accessorySheet.getRange(rowNum, 2).setValue(accessoryData.Accessory_Name);
                if (accessoryData.Accessory_Desc !== undefined) accessorySheet.getRange(rowNum, 3).setValue(accessoryData.Accessory_Desc);
                if (accessoryData.Total_Qty !== undefined) accessorySheet.getRange(rowNum, 4).setValue(accessoryData.Total_Qty);
                if (accessoryData.Available_Qty !== undefined) accessorySheet.getRange(rowNum, 5).setValue(accessoryData.Available_Qty);
                if (accessoryData.Active !== undefined) accessorySheet.getRange(rowNum, 6).setValue(accessoryData.Active);
                
                accessorySheet.getRange(rowNum, 9).setValue(userEmail);
                accessorySheet.getRange(rowNum, 10).setValue(timestamp);
                
                // Update item mappings if provided
                if (accessoryData.Item_Ids !== undefined && Array.isArray(accessoryData.Item_Ids)) {
                    // Delete existing mappings for this accessory
                    const mappingData = itemAccessorySheet.getDataRange().getValues();
                    for (let j = mappingData.length - 1; j >= 1; j--) {
                        if (mappingData[j][2] == accessoryId) {
                            itemAccessorySheet.deleteRow(j + 1);
                        }
                    }
                    
                    // Create new mappings
                    accessoryData.Item_Ids.forEach(itemId => {
                        const mappingId = getNextId(itemAccessorySheet, 0);
                        itemAccessorySheet.appendRow([
                            mappingId,
                            itemId,
                            accessoryId,
                            userEmail,
                            timestamp,
                            true
                        ]);
                    });
                }
                
                logInventoryActivity(`Admin ${userEmail} updated accessory ID: ${accessoryId}`);
                
                return { success: true, message: "Accessory updated successfully" };
            }
        }
        
        return { success: false, message: "Accessory not found" };
    } catch (error) {
        Logger.log("Error in updateAccessory: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Delete Accessory (Admin only)
*/
function deleteAccessory(accessoryId, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
        
        if (!accessorySheet) {
            throw new Error(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
        }
        
        // Delete accessory
        const data = accessorySheet.getDataRange().getValues();
        
        // Find and delete accessory row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == accessoryId) {
                accessorySheet.deleteRow(i + 1);
                
                // Delete all mappings for this accessory
                if (itemAccessorySheet) {
                    const mappingData = itemAccessorySheet.getDataRange().getValues();
                    for (let j = mappingData.length - 1; j >= 1; j--) {
                        if (mappingData[j][2] == accessoryId) {
                            itemAccessorySheet.deleteRow(j + 1);
                        }
                    }
                }
                
                logInventoryActivity(`Admin ${userEmail} deleted accessory ID: ${accessoryId}`);
                return { success: true, message: "Accessory deleted successfully" };
            }
        }
        
        return { success: false, message: "Accessory not found" };
    } catch (error) {
        Logger.log("Error in deleteAccessory: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Get Item-Accessory Mappings (Admin only)
*/
function getItemAccessoryMappings() {
    try {
        const ss = getActiveSheet();
        const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
        const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        
        if (!itemAccessorySheet) {
            return { success: true, mappings: [] };
        }
        
        const mappingData = itemAccessorySheet.getDataRange().getValues();
        const itemData = itemsSheet ? itemsSheet.getDataRange().getValues() : [];
        const accessoryData = accessorySheet ? accessorySheet.getDataRange().getValues() : [];
        
        const mappings = [];
        
        for (let i = 1; i < mappingData.length; i++) {
            const mapping = {
                Mapping_Id: mappingData[i][0],
                Item_Id: mappingData[i][1],
                Accessory_Id: mappingData[i][2],
                Created_By: mappingData[i][3],
                Created_At: formatDate(mappingData[i][4]),
                Active: mappingData[i][5]
            };
            
            // Find item name
            for (let j = 1; j < itemData.length; j++) {
                if (itemData[j][0] == mapping.Item_Id) {
                    mapping.Item_Name = itemData[j][1];
                    break;
                }
            }
            
            // Find accessory name
            for (let j = 1; j < accessoryData.length; j++) {
                if (accessoryData[j][0] == mapping.Accessory_Id) {
                    mapping.Accessory_Name = accessoryData[j][1];
                    break;
                }
            }
            
            mappings.push(mapping);
        }
        
        return { success: true, mappings: mappings };
    } catch (error) {
        Logger.log("Error in getItemAccessoryMappings: " + error.toString());
        return { success: false, message: error.toString(), mappings: [] };
    }
}

/*
@ Link Accessory to Items (Admin only)
*/
function linkAccessoryToItems(accessoryId, itemIds, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
        
        if (!itemAccessorySheet) {
            throw new Error(`Sheet "${ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }
        
        const timestamp = new Date();
        
        // Delete existing mappings for this accessory
        const mappingData = itemAccessorySheet.getDataRange().getValues();
        for (let j = mappingData.length - 1; j >= 1; j--) {
            if (mappingData[j][2] == accessoryId) {
                itemAccessorySheet.deleteRow(j + 1);
            }
        }
        
        // Create new mappings
        itemIds.forEach(itemId => {
            const mappingId = getNextId(itemAccessorySheet, 0);
            itemAccessorySheet.appendRow([
                mappingId,
                itemId,
                accessoryId,
                userEmail,
                timestamp,
                true
            ]);
        });
        
        logInventoryActivity(`Admin ${userEmail} linked accessory ${accessoryId} to ${itemIds.length} items`);
        
        return { success: true, message: "Accessory linked successfully" };
    } catch (error) {
        Logger.log("Error in linkAccessoryToItems: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

// ============================================
// User Management Functions
// ============================================

/*
@ Get Users (Admin only)
*/
function getUsers() {
    try {
        const ss = getActiveSheet();
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        
        if (!userSheet) {
            throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
        }
        
        const data = userSheet.getDataRange().getValues();
        
        if (data.length === 0) {
            return { success: true, data: [] };
        }
        
        const headers = data.shift();
        
        const users = data
            .filter(row => row.some(cell => cell !== ""))
            .map(row => ({
                Email: row[0],
                Permission: row[2],
                Active: row[3] === true || row[3] === 'TRUE'
            }));
        
        return { success: true, data: users };
    } catch (error) {
        Logger.log("Error in getUsers: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

/*
@ Create User (Admin only)
*/
function createUser(userData, adminEmail) {
    try {
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        
        if (!userSheet) {
            throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
        }
        
        // Normalize email: trim and convert to lowercase
        const normalizedEmail = userData.Email.trim().toLowerCase();
        
        // Validate email format
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(normalizedEmail)) {
            return { success: false, message: "Invalid email format" };
        }
        
        // Check if email already exists (case-insensitive)
        const data = userSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0].toLowerCase() === normalizedEmail) {
                return { success: false, message: "Email already exists" };
            }
        }
        
        // Hash password
        const passwordHash = hashPassword(userData.Password);
        
        userSheet.appendRow([
            normalizedEmail,
            passwordHash,
            userData.Permission,
            userData.Active !== false ? true : false
        ]);
        
        logSystemActivity(`Admin ${adminEmail} created user: ${normalizedEmail}`);
        
        return { success: true, message: "User created successfully" };
    } catch (error) {
        Logger.log("Error in createUser: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update User (Admin only)
*/
function updateUser(originalEmail, userData, adminEmail) {
    try {
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        
        if (!userSheet) {
            throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
        }
        
        const data = userSheet.getDataRange().getValues();
        const headers = data[0];
        
        // Find user row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === originalEmail) {
                const rowNum = i + 1;
                
                // Update permission
                if (userData.Permission !== undefined) {
                    userSheet.getRange(rowNum, 3).setValue(userData.Permission);
                }
                
                // Update active status
                if (userData.Active !== undefined) {
                    userSheet.getRange(rowNum, 4).setValue(userData.Active);
                }
                
                // Update password if provided
                if (userData.Password && userData.Password.trim() !== '') {
                    const passwordHash = hashPassword(userData.Password);
                    userSheet.getRange(rowNum, 2).setValue(passwordHash);
                }
                
                logSystemActivity(`Admin ${adminEmail} updated user: ${originalEmail}`);
                
                return { success: true, message: "User updated successfully" };
            }
        }
        
        return { success: false, message: "User not found" };
    } catch (error) {
        Logger.log("Error in updateUser: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Toggle User Active Status (Admin only)
*/
function toggleUserActive(userEmail, newActiveState, adminEmail) {
    try {
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        
        if (!userSheet) {
            throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
        }
        
        const data = userSheet.getDataRange().getValues();
        
        // Normalize email for case-insensitive comparison
        const normalizedEmail = userEmail.toLowerCase();
        
        // Find user row (case-insensitive)
        for (let i = 1; i < data.length; i++) {
            if (data[i][0].toLowerCase() === normalizedEmail) {
                const rowNum = i + 1;
                userSheet.getRange(rowNum, 4).setValue(newActiveState);
                
                logSystemActivity(`Admin ${adminEmail} ${newActiveState ? 'activated' : 'deactivated'} user: ${data[i][0]}`);
                
                return { success: true, message: `User ${newActiveState ? 'activated' : 'deactivated'} successfully` };
            }
        }
        
        return { success: false, message: "User not found" };
    } catch (error) {
        Logger.log("Error in toggleUserActive: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Change User Password (Admin only)
*/
function changeUserPassword(userEmail, newPassword, adminEmail) {
    try {
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        // Validate new password
        if (!newPassword || newPassword.trim().length < 6) {
            return { success: false, message: "Password must be at least 6 characters long" };
        }
        
        const ss = getActiveSheet();
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        
        if (!userSheet) {
            throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
        }
        
        const data = userSheet.getDataRange().getValues();
        
        // Normalize email for case-insensitive comparison
        const normalizedEmail = userEmail.toLowerCase();
        
        // Find user row (case-insensitive)
        for (let i = 1; i < data.length; i++) {
            if (data[i][0].toLowerCase() === normalizedEmail) {
                const rowNum = i + 1;
                
                // Hash and update password
                const passwordHash = hashPassword(newPassword);
                userSheet.getRange(rowNum, 2).setValue(passwordHash);
                
                logSystemActivity(`Admin ${adminEmail} changed password for user: ${data[i][0]}`);
                
                return { success: true, message: "Password changed successfully" };
            }
        }
        
        return { success: false, message: "User not found" };
    } catch (error) {
        Logger.log("Error in changeUserPassword: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Load Items with filter and pagination
*/
function loadItems (filterName = '', filterActive = 'ALL', page = 1, pageSize = 50)  {
    const cacheKey = 'loadItems';

    try {
        // Try to get from cache first (if enabled)
        let items = ENABLE_CACHE ? getCachedItems(cacheKey) : null;
        
        if (!items) {
            // Get fresh data from sheet
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);

            if (!itemsSheet) {
                Logger.log(`Sheet "${ITEM_SHEET_NAME}" not found`);
                return { data: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
            }

            const data = itemsSheet.getDataRange().getValues();
            
            if (data.length === 0) {
                return { data: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
            }

            const headers = data.shift(); // Remove header row

            // Convert to objects
            items = data
                .filter(row => row.some(cell => cell !== "")) // Filter empty rows
                .map(row => {
                    const item = {
                        Item_Id: row[0],
                        Item_Name: row[1],
                        Item_Desc: row[2],
                        Total_Qty: row[3],
                        Available_Qty: row[4],
                        Image: row[5],
                        Active: row[6],
                        Created_By: row[7],
                        Created_At: formatDate(row[8]),
                        Modified_By: row[9],
                        Modified_At: formatDate(row[10]),
                    };

                    return item;
                })

            // Cache the data (if enabled)
            if (ENABLE_CACHE) {
                setCache(cacheKey, items, DURATION_CACHE_SEC); // Cache for 5 minutes
            }
        }

        items = items.filter(item => {
                    // filter by name
                    const matchName =
                        !filterName ||
                        item.Item_Name?.toString().toLowerCase()
                            .includes(filterName.toLowerCase());

                    // filter by active
                    const matchActive =
                        filterActive === "ALL" ||
                        (filterActive === "TRUE" && (item.Active === true || item.Active === "TRUE")) ||
                        (filterActive === "FALSE" && (item.Active === false || item.Active === "FALSE"));

                    return matchName && matchActive;
                });

        // Apply pagination
        const total = items.length;
        const totalPages = Math.ceil(total / pageSize);
        const startIndex = (page - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const paginatedItems = items.slice(startIndex, endIndex);

        const result = {
            data: paginatedItems,
            total: total,
            page: page,
            pageSize: pageSize,
            totalPages: totalPages
        };

        console.log("[GAS] Load Items :", result)

        return result;

    } catch (error) {
        Logger.log("Error in loadItems: " + error.toString());
        return { data: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
    }
};

// ========================= Accessory Functions =========================
/*
@ Load Accessories by Item IDs
*/
function loadAccessories (itemIds = [], includeInactive = false)  {
    itemIds = arrayParser(itemIds);

    const cacheKey = `loadAccessories`;

    try {
        // Try to get from cache first (if enabled)
        let accessories = ENABLE_CACHE ? getCachedItems(cacheKey) : null;

        if (!accessories) {
            // Get fresh data from sheet
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
            const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);

            if (!accessorySheet) {
                Logger.log(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
                return { data: [] };
            }

            const data = accessorySheet.getDataRange().getValues();
            
            if (data.length === 0) {
                return { data: [] };
            }

            const headers = data.shift(); // Remove header row

            // Get item-accessory mappings
            let mappings = [];
            if (itemAccessorySheet) {
                const mappingData = itemAccessorySheet.getDataRange().getValues();
                mappings = mappingData.slice(1).map(row => ({
                    Mapping_Id: row[0],
                    Item_Id: row[1],
                    Accessory_Id: row[2],
                    Created_By: row[3],
                    Created_At: row[4],
                    Active: row[5]
                }));
            }

            // Convert to objects with mapped items
            accessories = data
                .filter(row => row.some(cell => cell !== "")) // Filter empty rows
                .map(row => {
                    const accessory = {
                        Accessory_Id: row[0],
                        Accessory_Name: row[1],
                        Accessory_Desc: row[2],
                        Total_Qty: row[3],
                        Available_Qty: row[4],
                        Active: row[5],
                        Created_By: row[6],
                        Created_At: formatDate(row[7]),
                        Modified_By: row[8],
                        Modified_At: formatDate(row[9]),
                        Item_Ids: [] // Will be populated from mappings
                    };

                    // Add item IDs from mappings
                    mappings.forEach(mapping => {
                        if (mapping.Accessory_Id == accessory.Accessory_Id && mapping.Active) {
                            accessory.Item_Ids.push(mapping.Item_Id);
                        }
                    });

                    return accessory;
                });

            // Cache the data (if enabled)
            if (ENABLE_CACHE) {
                setCache(cacheKey, accessories, DURATION_CACHE_SEC); // Cache for 5 minutes
            }
        }

        // Filter by itemIds and active status
        if (includeInactive) {
            // For management page - show all accessories
            accessories = accessories.filter(acc =>
                itemIds.length === 0 || acc.Item_Ids.some(id => itemIds.includes(String(id).trim()))
            );
        } else {
            // For normal usage - show only active accessories
            accessories = accessories.filter(acc =>
                (itemIds.length === 0 || acc.Item_Ids.some(id => itemIds.includes(String(id).trim()))) &&
                (acc.Active === true || String(acc.Active).toUpperCase() === "TRUE")
            );
        }

        const result = {
            success: true,
            data: accessories
        }

       console.log("[GAS] Load Accessories :", result)
                
        return result;

    }
    catch (error) {
        Logger.log("Error in loadAccessories: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

// ============================================
// Request Functions
// ============================================

/*
@ Get Request List
*/
function getRequestList() {
    const ss = getActiveSheet();
    const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
    const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
    const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);

    if (!requestSheet) {
        throw new Error(`Sheet "${REQUEST_SHEET_NAME}" not found`);
    }

    if (!requestItemSheet) {
        throw new Error(`Sheet "${REQUEST_ITEM_SHEET_NAME}" not found`);
    }

    if (!requestItemAccessorySheet) {
        throw new Error(`Sheet "${REQUEST_ITEM_ACCESSORY_SHEET_NAME}" not found`);
    }

    let requestList = [];

    const requestData = requestSheet.getDataRange().getValues();
    const requestHeaders = requestData[0];

    const itemsData = requestItemSheet.getDataRange().getValues();
    const itemsHeaders = itemsData[0];

    const accessoryData = requestItemAccessorySheet.getDataRange().getValues();
    const accessoryHeaders = accessoryData[0];
    
    for (let i = 1; i < requestData.length; i++) {
        const row = requestData[i];

        const requestStatus = row[requestHeaders.indexOf("Status")];

        const request = {
            Request_Id: row[requestHeaders.indexOf("Request_Id")],
            Requirer_Name: row[requestHeaders.indexOf("Requirer_Name")],
            Status: row[requestHeaders.indexOf("Status")],
            Request_Date: formatDate(row[requestHeaders.indexOf("Request_Date")], 'yyyy-MM-dd'),
            Distributed_Date: formatDate(row[requestHeaders.indexOf("Distributed_Date")], 'yyyy-MM-dd'),
            Return_Date: formatDate(row[requestHeaders.indexOf("Return_Date")], 'yyyy-MM-dd'),
            Remark: row[requestHeaders.indexOf("Remark")],
            Created_By: row[requestHeaders.indexOf("Created_By")],
            Created_At: formatDate(row[requestHeaders.indexOf("Created_At")]),
            Modified_By: row[requestHeaders.indexOf("Modified_By")],
            Modified_At: formatDate(row[requestHeaders.indexOf("Modified_At")]),
            Items: []
        };

        for (let j = 1; j < itemsData.length; j++) {
            const itemRow = itemsData[j];
            const itemRequestId = itemRow[itemsHeaders.indexOf("Request_Id")];

            if (itemRequestId != request.Request_Id) continue;

            const item = {
                Request_Id: itemRequestId,
                Item_Index: itemRow[itemsHeaders.indexOf("Item_Index")],
                Item_Id: itemRow[itemsHeaders.indexOf("Item_Id")],
                Item_Name: itemRow[itemsHeaders.indexOf("Item_Name")],
                Qty: Number(itemRow[itemsHeaders.indexOf("Qty")]),
                Returned_Qty: Number(itemRow[itemsHeaders.indexOf("Returned_Qty")]),
                Status: itemRow[itemsHeaders.indexOf("Status")],
                Accessories: []
            };

            for (let k = 1; k < accessoryData.length; k++) {
                const accRow = accessoryData[k];

                if (
                    accRow[accessoryHeaders.indexOf("Request_Id")] == request.Request_Id &&
                    accRow[accessoryHeaders.indexOf("Item_Index")] == item.Item_Index
                ) {
                    item.Accessories.push({
                        Request_Id: accRow[accessoryHeaders.indexOf("Request_Id")],
                        Item_Index: accRow[accessoryHeaders.indexOf("Item_Index")],
                        Accessory_Index: accRow[accessoryHeaders.indexOf("Accessory_Index")],
                        Accessory_Id: accRow[accessoryHeaders.indexOf("Accessory_Id")],
                        Accessory_Name: accRow[accessoryHeaders.indexOf("Accessory_Name")],
                        Qty: Number(accRow[accessoryHeaders.indexOf("Qty")]),
                        Returned_Qty: Number(accRow[accessoryHeaders.indexOf("Returned_Qty")]),
                        Status: accRow[accessoryHeaders.indexOf("Status")]
                    });
                }
            }

            request.Items.push(item);
        }

        requestList.push(request);
    }

    console.log(JSON.stringify(requestList));
    return { success: true, data: requestList };
}

/*
@ Get Row Request
*/
function getRowRequest(requestId = null) {
    try {
        const ss = getActiveSheet();
        const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);

        if (!itemsSheet) {
            throw new Error(`Sheet "${ITEM_SHEET_NAME}" not found`);
        }

        if (!requestSheet) {
            throw new Error(`Sheet "${REQUEST_SHEET_NAME}" not found`);
        }

        if (!requestItemSheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_SHEET_NAME}" not found`);
        }

        if (!requestItemAccessorySheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }

        let request = null;

        const requestData = requestSheet.getDataRange().getValues();
        const requestHeaders = requestData[0];

        const itemsData = requestItemSheet.getDataRange().getValues();
        const itemsHeaders = itemsData[0];

        const accessoryData = requestItemAccessorySheet.getDataRange().getValues();
        const accessoryHeaders = accessoryData[0];

        for (let i = 1; i < requestData.length; i++) {
            const row = requestData[i];
            if (row[requestHeaders.indexOf("Request_Id")] != +requestId) continue;

            request = {
                Request_Id: row[requestHeaders.indexOf("Request_Id")],
                Requirer_Name: row[requestHeaders.indexOf("Requirer_Name")],
                Status: row[requestHeaders.indexOf("Status")],
                Request_Date: formatDate(row[requestHeaders.indexOf("Request_Date")], 'yyyy-MM-dd'),
                Distributed_Date: formatDate(row[requestHeaders.indexOf("Distributed_Date")], 'yyyy-MM-dd'),
                Return_Date: formatDate(row[requestHeaders.indexOf("Return_Date")], 'yyyy-MM-dd'),
                Remark: row[requestHeaders.indexOf("Remark")],
                Created_By: row[requestHeaders.indexOf("Created_By")],
                Created_At: formatDate(row[requestHeaders.indexOf("Created_At")]),
                Modified_By: row[requestHeaders.indexOf("Modified_By")],
                Modified_At: formatDate(row[requestHeaders.indexOf("Modified_At")]),
                Items: []
            };
            
            for (let j = 1; j < itemsData.length; j++) {
                const itemRow = itemsData[j];
                const itemRequestId = itemRow[itemsHeaders.indexOf("Request_Id")];

                if (itemRequestId != request.Request_Id) continue;

                const item = {
                    Request_Id: itemRequestId,
                    Item_Index: itemRow[itemsHeaders.indexOf("Item_Index")],
                    Item_Id: itemRow[itemsHeaders.indexOf("Item_Id")],
                    Item_Name: itemRow[itemsHeaders.indexOf("Item_Name")],
                    // get Item_Image from itemsSheet in column index 5
                    Item_Image: (() => {
                        const itemId = itemRow[itemsHeaders.indexOf("Item_Id")];
                        const itemsDataFull = itemsSheet.getDataRange().getValues();
                        const itemsHeadersFull = itemsDataFull[0];
                        for (let k = 1; k < itemsDataFull.length; k++) {
                            const itemFullRow = itemsDataFull[k];
                            if (itemFullRow[itemsHeadersFull.indexOf("Item_Id")] == itemId) {
                                return itemFullRow[itemsHeadersFull.indexOf("Image")];
                            }
                        }
                        return null;
                    })(),
                    Qty: Number(itemRow[itemsHeaders.indexOf("Qty")]),
                    Returned_Qty: Number(itemRow[itemsHeaders.indexOf("Returned_Qty")]),
                    Status: itemRow[itemsHeaders.indexOf("Status")],
                    Accessories: []
                };

                for (let k = 1; k < accessoryData.length; k++) {
                    const accRow = accessoryData[k];
                    
                    if (
                        accRow[accessoryHeaders.indexOf("Request_Id")] == request.Request_Id &&
                        accRow[accessoryHeaders.indexOf("Item_Index")] == item.Item_Index
                    ) {
                        item.Accessories.push({
                            Request_Id: accRow[accessoryHeaders.indexOf("Request_Id")],
                            Item_Index: accRow[accessoryHeaders.indexOf("Item_Index")],
                            Accessory_Index: accRow[accessoryHeaders.indexOf("Accessory_Index")],
                            Accessory_Id: accRow[accessoryHeaders.indexOf("Accessory_Id")],
                            Accessory_Name: accRow[accessoryHeaders.indexOf("Accessory_Name")],
                            Qty: Number(accRow[accessoryHeaders.indexOf("Qty")]),
                            Returned_Qty: Number(accRow[accessoryHeaders.indexOf("Returned_Qty")]),
                            Status: accRow[accessoryHeaders.indexOf("Status")]
                        });
                    }
                }
                request.Items.push(item);
            }
            break; // Exit loop after finding the request
        }

        console.log(JSON.stringify(request));

        return { success: true, data: request };
    }
    catch (error) {
        Logger.log("Error in getRowRequest: " + error.toString());
        throw error;
    }
}

/*
@ Submit Request
*/
function submitRequest(submitRequestData) {
    try {
        const ss = getActiveSheet();
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);

        if (!requestSheet) {
            throw new Error(`Sheet "${REQUEST_SHEET_NAME}" not found`);
        }

        if (!requestItemSheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_SHEET_NAME}" not found`);
        }

        if (!requestItemAccessorySheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }

        // ============================================
        // 1. New or Edit Request
        // ============================================
        if (submitRequestData.IsNew) {
            let request = [];
            let requestId = getNextId(requestSheet, 0);
            let status = 'Submit';
            const timestamp = formatDate(new Date());
            const userEmail = Session.getActiveUser().getEmail();

            // Check stock availability first (before any writes)
            const stockValidation = validateStockAvailable(submitRequestData.Items);

            if (!stockValidation.success) {
                throw new Error(`Insufficient stock: ${stockValidation.message}`);
            }

            request.push([
                requestId,
                submitRequestData.Requirer_Name,
                status,
                submitRequestData.Request_Date,
                null,
                submitRequestData.Return_Date,
                submitRequestData.Remark,
                userEmail,
                timestamp,
                userEmail,
                timestamp
            ]);

            // insert to item
            let requestItems = [];
            let requestItemsAccessories = [];
            submitRequestData.Items.forEach((item, itemIndex) => {
                requestItems.push([
                    requestId,
                    itemIndex + 1,
                    item.Item_Id,
                    item.Item_Name,
                    item.Qty,
                    0,
                    status
                ]);

                // insert to accessory
                item.Accessories.forEach((accessory, accessoryIndex) => {
                    requestItemsAccessories.push([
                        requestId,
                        itemIndex + 1,
                        accessoryIndex + 1,
                        accessory.Accessory_Id,
                        accessory.Accessory_Name,
                        accessory.Qty,
                        0,
                        status
                    ]);
                });
            });

            if (request.length > 0)
                requestSheet.getRange(requestSheet.getLastRow() + 1, 1, request.length, request[0].length).setValues(request);

            if (requestItems.length > 0)
                requestItemSheet.getRange(requestItemSheet.getLastRow() + 1, 1, requestItems.length, requestItems[0].length).setValues(requestItems);

            if (requestItemsAccessories.length > 0)
                requestItemAccessorySheet.getRange(requestItemAccessorySheet.getLastRow() + 1, 1, requestItemsAccessories.length, requestItemsAccessories[0].length).setValues(requestItemsAccessories);

            logRequestActivity("Create Request");
        } else {
            // Edit existing request - only allowed if status is 'Submit' (Pending)
            const result = editRequest(submitRequestData);
            if (!result.success) {
                throw new Error(result.message);
            }
        }

        Logger.log("Submit Request Data: " + JSON.stringify(submitRequestData));

        return { success: true, message: "Request submitted successfully" };
    }
    catch (error) {
        Logger.log("Error in submitRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Edit Request - Only allowed for requests with status 'Submit' (Pending)
@ Users can edit their own requests, admins can edit any
*/
function editRequest(editRequestData) {
    try {
        const userEmail = Session.getActiveUser().getEmail();
        const ss = getActiveSheet();
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);

        // Find the request
        const requestData = requestSheet.getDataRange().getValues();
        const headers = requestData[0];
        const requestIdIndex = headers.indexOf("Request_Id");
        const statusIndex = headers.indexOf("Status");
        const createdByIndex = headers.indexOf("Created_By");
        
        let rowIndex = -1;
        let rowData = null;

        for (let i = 1; i < requestData.length; i++) {
            if (requestData[i][requestIdIndex] == editRequestData.Request_Id) {
                rowIndex = i + 1;
                rowData = requestData[i];
                break;
            }
        }

        if (rowIndex === -1) {
            return { success: false, message: "Request not found" };
        }

        // Check if request is in editable status
        if (rowData[statusIndex] !== "Submit") {
            return { success: false, message: "Only pending requests can be edited" };
        }

        // Check permission: User can edit own requests, Admin can edit any
        const isAdmin = checkAdminPermission(userEmail);
        const createdBy = rowData[createdByIndex];
        
        if (!isAdmin && createdBy !== userEmail) {
            return { success: false, message: "You can only edit your own requests" };
        }

        // Check stock availability
        const stockValidation = validateStockAvailable(editRequestData.Items);
        if (!stockValidation.success) {
            return { success: false, message: `Insufficient stock: ${stockValidation.message}` };
        }

        // Delete existing items and accessories for this request
        const requestItemData = requestItemSheet.getDataRange().getValues();
        for (let i = requestItemData.length - 1; i >= 1; i--) {
            if (requestItemData[i][0] == editRequestData.Request_Id) {
                requestItemSheet.deleteRow(i + 1);
            }
        }

        const requestItemAccessoryData = requestItemAccessorySheet.getDataRange().getValues();
        for (let i = requestItemAccessoryData.length - 1; i >= 1; i--) {
            if (requestItemAccessoryData[i][0] == editRequestData.Request_Id) {
                requestItemAccessorySheet.deleteRow(i + 1);
            }
        }

        // Update main request
        const timestamp = formatDate(new Date());
        requestSheet.getRange(rowIndex, headers.indexOf("Requirer_Name") + 1).setValue(editRequestData.Requirer_Name);
        requestSheet.getRange(rowIndex, headers.indexOf("Request_Date") + 1).setValue(editRequestData.Request_Date);
        requestSheet.getRange(rowIndex, headers.indexOf("Return_Date") + 1).setValue(editRequestData.Return_Date);
        requestSheet.getRange(rowIndex, headers.indexOf("Remark") + 1).setValue(editRequestData.Remark);
        requestSheet.getRange(rowIndex, headers.indexOf("Modified_By") + 1).setValue(userEmail);
        requestSheet.getRange(rowIndex, headers.indexOf("Modified_At") + 1).setValue(timestamp);

        // Insert new items
        let requestItems = [];
        let requestItemsAccessories = [];
        const status = 'Submit';

        editRequestData.Items.forEach((item, itemIndex) => {
            requestItems.push([
                editRequestData.Request_Id,
                itemIndex + 1,
                item.Item_Id,
                item.Item_Name,
                item.Qty,
                0,
                status
            ]);

            // insert accessories
            item.Accessories.forEach((accessory, accessoryIndex) => {
                requestItemsAccessories.push([
                    editRequestData.Request_Id,
                    itemIndex + 1,
                    accessoryIndex + 1,
                    accessory.Accessory_Id,
                    accessory.Accessory_Name,
                    accessory.Qty,
                    0,
                    status
                ]);
            });
        });

        if (requestItems.length > 0)
            requestItemSheet.getRange(requestItemSheet.getLastRow() + 1, 1, requestItems.length, requestItems[0].length).setValues(requestItems);

        if (requestItemsAccessories.length > 0)
            requestItemAccessorySheet.getRange(requestItemAccessorySheet.getLastRow() + 1, 1, requestItemsAccessories.length, requestItemsAccessories[0].length).setValues(requestItemsAccessories);

        logRequestActivity("Edit Request");

        return { success: true, message: "Request updated successfully" };
    } catch (error) {
        Logger.log("Error in editRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

function validateStockAvailable(items) {
    const ss = getActiveSheet();
    const itemsData = ss.getSheetByName(ITEM_SHEET_NAME).getDataRange().getValues();
    const itemsMap = {};

    const accessoryData = ss.getSheetByName(ACCESSORY_SHEET_NAME).getDataRange().getValues();
    const accessoryMap = {};

    // build items map
    for (let i = 1; i < itemsData.length; i++) {
        itemsMap[itemsData[i][0]] = itemsData[i][4]; // "Item_Id":"Available_Qty"
    }

    // build accessory map
    for (let i = 1; i < accessoryData.length; i++) {
        accessoryMap[accessoryData[i][0]] = accessoryData[i][5]; // "Accessory_Id":"Available_Qty"
    }

    // check each item
    for (let item of items) {
        const availableQty = itemsMap[item.Item_Id] || 0;
        if (availableQty < item.Qty) {
            return { success: false, message: `Insufficient stock for item ID ${item.Item_Id}` };
        }

        // check each accessory
        for (let accessory of item.Accessories) {
            const accAvailableQty = accessoryMap[accessory.Accessory_Id] || 0;
            if(accAvailableQty < accessory.Qty) {
                return { success: false, message: `Insufficient stock for accessory ID ${accessory.Accessory_Id}` };
            }
        }
    }

    return { success: true };
}

/*
@ Manage Stock: Support partial return
@ action: 'decrease' (for borrow) or 'increase' (for return)
*/
function manageStock(items, action = 'decrease') {
    const ss = getActiveSheet();
    const itemsData = ss.getSheetByName(ITEM_SHEET_NAME).getDataRange().getValues();
    const itemHeaders = itemsData[0];
    const itemIdCol = itemHeaders.indexOf("Item_Id");
    const itemsMap = {};

    const accessoryData = ss.getSheetByName(ACCESSORY_SHEET_NAME).getDataRange().getValues();
    const accessoryHeaders = accessoryData[0];
    const accessoryIdCol = accessoryHeaders.indexOf("Accessory_Id");
    const accessoryMap = {};

    // build items map
    for (let i = 1; i < itemsData.length; i++) {
        itemsMap[itemsData[i][itemIdCol]] = { 
            Item_Id: itemsData[i][itemIdCol], 
            Item_Row_Index: i + 1, 
            Available_Qty: itemsData[i][itemHeaders.indexOf("Available_Qty")] 
        };
    }

    // build accessory map
    for (let i = 1; i < accessoryData.length; i++) {
        accessoryMap[accessoryData[i][accessoryIdCol]] = { 
            Accessory_Id: accessoryData[i][accessoryIdCol], 
            Accessory_Row_Index: i + 1, 
            Available_Qty: accessoryData[i][accessoryHeaders.indexOf("Available_Qty")] 
        };
    }

    let updateItems = [];
    let updateAccessories = [];

    // check each item
    for (let item of items) {
        const itemData = itemsMap[item.Item_Id];
        if (!itemData) {
            Logger.log(`Warning: Item ID ${item.Item_Id} not found in stock`);
            continue;
        }

        // Use Return_Qty for returns, Qty for borrows
        const qtyToUpdate = action === 'increase' ? (item.Return_Qty || item.Qty) : item.Qty;
        
        if (itemData.Item_Id == item.Item_Id) {
            if (action === 'decrease') {
                updateItems.push({ 
                    Item_Id: item.Item_Id, 
                    Item_Row_Index: itemData.Item_Row_Index, 
                    Available_Qty: itemData.Available_Qty, 
                    Update_Qty: qtyToUpdate, 
                    New_Available_Qty: itemData.Available_Qty - qtyToUpdate 
                });
            } else if (action === 'increase') {
                updateItems.push({ 
                    Item_Id: item.Item_Id, 
                    Item_Row_Index: itemData.Item_Row_Index, 
                    Available_Qty: itemData.Available_Qty, 
                    Update_Qty: qtyToUpdate, 
                    New_Available_Qty: itemData.Available_Qty + qtyToUpdate 
                });
            }
        }

        // check each accessory
        if (item.Accessories && item.Accessories.length > 0) {
            for (let accessory of item.Accessories) {
                const accData = accessoryMap[accessory.Accessory_Id];
                if (!accData) {
                    Logger.log(`Warning: Accessory ID ${accessory.Accessory_Id} not found in stock`);
                    continue;
                }

                // Use Return_Qty for returns, Qty for borrows
                const accQtyToUpdate = action === 'increase' ? (accessory.Return_Qty || accessory.Qty) : accessory.Qty;
                
                if (accData.Accessory_Id == accessory.Accessory_Id) {
                    if (action === 'decrease') {
                        updateAccessories.push({ 
                            Accessory_Id: accessory.Accessory_Id, 
                            Accessory_Row_Index: accData.Accessory_Row_Index, 
                            Available_Qty: accData.Available_Qty, 
                            Update_Qty: accQtyToUpdate, 
                            New_Available_Qty: accData.Available_Qty - accQtyToUpdate 
                        });
                    } else if (action === 'increase') {
                        updateAccessories.push({ 
                            Accessory_Id: accessory.Accessory_Id, 
                            Accessory_Row_Index: accData.Accessory_Row_Index, 
                            Available_Qty: accData.Available_Qty, 
                            Update_Qty: accQtyToUpdate, 
                            New_Available_Qty: accData.Available_Qty + accQtyToUpdate 
                        });
                    }
                }
            }
        }
    }

    // Update items stock
    updateItems.forEach(item => {
        Logger.log(`Updating Item ID ${item.Item_Id}: ${item.Available_Qty} ${action === 'decrease' ? '-' : '+'} ${item.Update_Qty} = ${item.New_Available_Qty}`);
        ss.getSheetByName(ITEM_SHEET_NAME).getRange(item.Item_Row_Index, itemHeaders.indexOf("Available_Qty") + 1).setValue(item.New_Available_Qty);
        ss.getSheetByName(ITEM_SHEET_NAME).getRange(item.Item_Row_Index, itemHeaders.indexOf("Modified_At") + 1).setValue(formatDate(new Date()));
    });

    // Update accessories stock
    updateAccessories.forEach(accessory => {
        Logger.log(`Updating Accessory ID ${accessory.Accessory_Id}: ${accessory.Available_Qty} ${action === 'decrease' ? '-' : '+'} ${accessory.Update_Qty} = ${accessory.New_Available_Qty}`);
        ss.getSheetByName(ACCESSORY_SHEET_NAME).getRange(accessory.Accessory_Row_Index, accessoryHeaders.indexOf("Available_Qty") + 1).setValue(accessory.New_Available_Qty);
        ss.getSheetByName(ACCESSORY_SHEET_NAME).getRange(accessory.Accessory_Row_Index, accessoryHeaders.indexOf("Modified_At") + 1).setValue(formatDate(new Date()));
    });

    return { 
        success: true, 
        updatedItems: updateItems.length, 
        updatedAccessories: updateAccessories.length 
    };
}

/*
@ Distribute Request (Admin only)
*/
function distributeRequest(request) {
    try {
        // Check admin permission
        const userEmail = Session.getActiveUser().getEmail();
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required to distribute requests" };
        }
        
        const ss = getActiveSheet();
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);

        if (!requestSheet) {
            throw new Error(`Sheet "${REQUEST_SHEET_NAME}" not found`);
        }

        if (!requestItemSheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_SHEET_NAME}" not found`);
        }

        if (!requestItemAccessorySheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }

        console.log(JSON.stringify(request));

        const requestData = requestSheet.getDataRange().getValues();
        const headers = requestData[0];
        const requestIdIndex = headers.indexOf("Request_Id");
        let rowIndex = -1;
        let rowData = null;

        for (let i = 1; i < requestData.length; i++) {
            const dataRequestId = +requestData[i][requestIdIndex];
            if (dataRequestId === +request.Request_Id) {
                rowIndex = i + 1; // +1 because data array is 0-based and sheet rows are 1-based
                rowData = requestData[i];
                break;
            }
        }

        if (rowIndex === -1) {
            throw new Error(`Request ID ${request.Request_Id} not found`);
        }

        if (rowData[headers.indexOf("Status")] !== "Submit") {
            throw new Error(`Request ID ${request.Request_Id} is not in 'Submit' status`);
        }

        const statusIndex = headers.indexOf("Status");
        const distributedDateIndex = headers.indexOf("Distributed_Date");
        const modifiedByIndex = headers.indexOf("Modified_By");
        const modifiedAtIndex = headers.indexOf("Modified_At");

        // reduce stock
        manageStock(request.Items, 'decrease');

        requestSheet.getRange(rowIndex, statusIndex + 1).setValue("Distributed");
        requestSheet.getRange(rowIndex, distributedDateIndex + 1).setValue(formatDate(new Date(), 'yyyy-MM-dd'));
        requestSheet.getRange(rowIndex, modifiedByIndex + 1).setValue(userEmail);
        requestSheet.getRange(rowIndex, modifiedAtIndex + 1).setValue(formatDate(new Date()));

        const requestItemData = requestItemSheet.getDataRange().getValues();
        const requestItemRequestIdIndex = requestItemData[0].indexOf("Request_Id");
        const requestItemAccessoryData = requestItemAccessorySheet.getDataRange().getValues();
        const requestItemAccessoryRequestIdIndex = requestItemAccessoryData[0].indexOf("Request_Id");

        // update status of requestItemData in column index 6
        for (let i = 1; i < requestItemData.length; i++) {
            const dataRequestId = +requestItemData[i][requestItemRequestIdIndex];
            if (dataRequestId === +request.Request_Id) {
                requestItemSheet.getRange(i + 1, 7).setValue("Distributed");
            }
        }

        // update status of requestItemAccessoryData in column index 7
        for (let i = 1; i < requestItemAccessoryData.length; i++) {
            const dataRequestId = +requestItemAccessoryData[i][requestItemAccessoryRequestIdIndex];
            if (dataRequestId === +request.Request_Id) {
                requestItemAccessorySheet.getRange(i + 1, 8).setValue("Distributed");
            }
        }

        logRequestActivity("Distribute Request");

        return { success: true, message: `Request ID ${request.Request_Id} distributed successfully` };
    }
    catch (error) {
        Logger.log("Error in distributeRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Cancel Request
*/
function cancelRequest(requestId) {
    try {
        const ss = getActiveSheet();
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);

        if (!requestSheet) {
            throw new Error(`Sheet "${REQUEST_SHEET_NAME}" not found`);
        }

        if (!requestItemSheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_SHEET_NAME}" not found`);
        }

        if (!requestItemAccessorySheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }

        const requestData = requestSheet.getDataRange().getValues();
        const headers = requestData[0];
        const requestIdIndex = headers.indexOf("Request_Id");
        let rowIndex = -1;
        let rowData = null;

        for (let i = 1; i < requestData.length; i++) {
            const dataRequestId = +requestData[i][requestIdIndex];
            if (dataRequestId === +requestId) {
                rowIndex = i + 1;
                rowData = requestData[i];
                break;
            }
        }
        
        if (rowIndex === -1) {
            throw new Error(`Request ID ${requestId} not found`);
        }

        const statusIndex = headers.indexOf("Status");

        if (rowData[statusIndex] !== "Submit") {
            throw new Error(`Request ID ${requestId} is not in 'Submit' status`);
        }

        const modifiedByIndex = headers.indexOf("Modified_By");
        const modifiedAtIndex = headers.indexOf("Modified_At");
        requestSheet.getRange(rowIndex, statusIndex + 1).setValue("Cancelled");
        requestSheet.getRange(rowIndex, modifiedByIndex + 1).setValue(Session.getActiveUser().getEmail());
        requestSheet.getRange(rowIndex, modifiedAtIndex + 1).setValue(formatDate(new Date()));
        
        const requestItemData = requestItemSheet.getDataRange().getValues();
        const requestItemRequestIdIndex = requestItemData[0].indexOf("Request_Id");
        const requestItemAccessoryData = requestItemAccessorySheet.getDataRange().getValues();
        const requestItemAccessoryRequestIdIndex = requestItemAccessoryData[0].indexOf("Request_Id");

        // update status of requestItemData in column index 6
        for (let i = 1; i < requestItemData.length; i++) {
            const dataRequestId = +requestItemData[i][requestItemRequestIdIndex];
            if (dataRequestId === +requestId) {
                requestItemSheet.getRange(i + 1, 7).setValue("Cancelled");
            }
        }
        // update status of requestItemAccessoryData in column index 7
        for (let i = 1; i < requestItemAccessoryData.length; i++) {
            const dataRequestId = +requestItemAccessoryData[i][requestItemAccessoryRequestIdIndex];
            if (dataRequestId === +requestId) {
                requestItemAccessorySheet.getRange(i + 1, 8).setValue("Cancelled");
            }
        }

        logRequestActivity("Cancel Request");

        return { success: true, message: `Request ID ${requestId} cancelled successfully` };
    }
    catch (error) {
        Logger.log("Error in cancelRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Return Request: Support partial return (Admin only)
*/
function returnRequest(returnRequestData) {
    try {
        // Check admin permission
        const userEmail = Session.getActiveUser().getEmail();
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required to return requests" };
        }
        
        const ss = getActiveSheet();
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);
        
        if (!requestSheet) {
            throw new Error(`Sheet "${REQUEST_SHEET_NAME}" not found`);
        }
        if (!requestItemSheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_SHEET_NAME}" not found`);
        }
        if (!requestItemAccessorySheet) {
            throw new Error(`Sheet "${REQUEST_ITEM_ACCESSORY_SHEET_NAME}" not found`);
        }

        Logger.log("Return Request Data: " + JSON.stringify(returnRequestData));

        const requestData = requestSheet.getDataRange().getValues();
        const headers = requestData[0];
        const requestIdIndex = headers.indexOf("Request_Id");
        const statusIndex = headers.indexOf("Status");
        const modifiedByIndex = headers.indexOf("Modified_By");
        const modifiedAtIndex = headers.indexOf("Modified_At");
        
        let rowIndex = -1;
        let rowData = null;

        // Find the request
        for (let i = 1; i < requestData.length; i++) {
            const dataRequestId = requestData[i][requestIdIndex];
            if (dataRequestId == returnRequestData.Request_Id) {
                rowIndex = i + 1;
                rowData = requestData[i];
                break;
            }
        }

        if (rowIndex === -1) {
            throw new Error(`Request ID ${returnRequestData.Request_Id} not found`);
        }

        if (rowData[statusIndex] !== "Distributed" && rowData[statusIndex] !== "Partial_Returned") {
            throw new Error(`Request ID ${returnRequestData.Request_Id} is not in 'Distributed' or 'Partial_Returned' status. Current status: ${rowData[statusIndex]}`);
        }

        // Check if this is a partial return
        const isPartialReturn = returnRequestData.Return_Type === 'partial';
        
        // Check if all items are fully returned
        let allItemsFullyReturned = true;
        for (let item of returnRequestData.Items) {
            if (item.Return_Qty < item.Borrowed_Qty) {
                allItemsFullyReturned = false;
                break;
            }
            // Check accessories
            if (item.Accessories && item.Accessories.length > 0) {
                for (let acc of item.Accessories) {
                    if (acc.Return_Qty < acc.Borrowed_Qty) {
                        allItemsFullyReturned = false;
                        break;
                    }
                }
            }
            if (!allItemsFullyReturned) break;
        }

        // Increase stock based on returned quantities
        const stockUpdateResult = manageStock(returnRequestData.Items, 'increase');
        Logger.log(`Stock update result: ${stockUpdateResult.updatedItems} items, ${stockUpdateResult.updatedAccessories} accessories`);

        // Update request status
        let newStatus = allItemsFullyReturned ? "Returned" : "Partial_Returned";
        requestSheet.getRange(rowIndex, statusIndex + 1).setValue(newStatus);
        requestSheet.getRange(rowIndex, modifiedByIndex + 1).setValue(Session.getActiveUser().getEmail());
        requestSheet.getRange(rowIndex, modifiedAtIndex + 1).setValue(formatDate(new Date()));

        // Update request items status and returned quantity
        const requestItemData = requestItemSheet.getDataRange().getValues();
        const requestItemHeaders = requestItemData[0];
        const requestItemRequestIdIndex = requestItemHeaders.indexOf("Request_Id");
        const requestItemIdIndex = requestItemHeaders.indexOf("Item_Id");
        const requestItemStatusIndex = requestItemHeaders.indexOf("Status");
        const requestItemReturnedQtyIndex = requestItemHeaders.indexOf("Returned_Qty");

        for (let i = 1; i < requestItemData.length; i++) {
            const dataRequestId = requestItemData[i][requestItemRequestIdIndex];
            const dataItemId = requestItemData[i][requestItemIdIndex];
            
            if (dataRequestId == returnRequestData.Request_Id) {
                // Find matching item in return data
                const returnItem = returnRequestData.Items.find(item => item.Item_Id == dataItemId);
                if (returnItem) {
                    // Get current returned qty and add new return qty
                    const currentReturnedQty = requestItemData[i][requestItemReturnedQtyIndex] || 0;
                    const newReturnedQty = currentReturnedQty + returnItem.Return_Qty;
                    
                    // Update returned quantity (column index 6, which is index 5 in 0-based)
                    requestItemSheet.getRange(i + 1, 6).setValue(newReturnedQty);
                    
                    // Determine status based on total returned vs borrowed
                    let itemStatus = "Distributed";
                    if (newReturnedQty >= returnItem.Borrowed_Qty) {
                        itemStatus = "Returned";
                    } else if (newReturnedQty > 0) {
                        itemStatus = "Partial_Returned";
                    }
                    
                    requestItemSheet.getRange(i + 1, requestItemStatusIndex + 1).setValue(itemStatus);
                }
            }
        }

        // Update request item accessories status and returned quantity
        const requestItemAccessoryData = requestItemAccessorySheet.getDataRange().getValues();
        const requestItemAccessoryHeaders = requestItemAccessoryData[0];
        const requestItemAccessoryRequestIdIndex = requestItemAccessoryHeaders.indexOf("Request_Id");
        const requestItemAccessoryIdIndex = requestItemAccessoryHeaders.indexOf("Accessory_Id");
        const requestItemAccessoryStatusIndex = requestItemAccessoryHeaders.indexOf("Status");
        const requestItemAccessoryReturnedQtyIndex = requestItemAccessoryHeaders.indexOf("Returned_Qty");

        for (let i = 1; i < requestItemAccessoryData.length; i++) {
            const dataRequestId = requestItemAccessoryData[i][requestItemAccessoryRequestIdIndex];
            const dataAccessoryId = requestItemAccessoryData[i][requestItemAccessoryIdIndex];
            
            if (dataRequestId == returnRequestData.Request_Id) {
                // Find matching accessory in return data
                let accessoryStatus = "Distributed";
                let returnedQty = 0;
                
                for (let item of returnRequestData.Items) {
                    if (item.Accessories && item.Accessories.length > 0) {
                        const returnAcc = item.Accessories.find(acc => acc.Accessory_Id == dataAccessoryId);
                        if (returnAcc) {
                            // Get current returned qty and add new return qty
                            const currentReturnedQty = requestItemAccessoryData[i][requestItemAccessoryReturnedQtyIndex] || 0;
                            returnedQty = currentReturnedQty + returnAcc.Return_Qty;
                            
                            // Update returned quantity (column index 7, which is index 6 in 0-based)
                            requestItemAccessorySheet.getRange(i + 1, 7).setValue(returnedQty);
                            
                            // Determine status
                            if (returnedQty >= returnAcc.Borrowed_Qty) {
                                accessoryStatus = "Returned";
                            } else if (returnedQty > 0) {
                                accessoryStatus = "Partial_Returned";
                            }
                            break;
                        }
                    }
                }
                
                requestItemAccessorySheet.getRange(i + 1, requestItemAccessoryStatusIndex + 1).setValue(accessoryStatus);
            }
        }

        // Audit log
        logRequestActivity("Return Request");

        return { 
            success: true, 
            message: `Request ID ${returnRequestData.Request_Id} ${allItemsFullyReturned ? 'returned' : 'partially returned'} successfully`,
            status: newStatus,
            itemsUpdated: stockUpdateResult.updatedItems,
            accessoriesUpdated: stockUpdateResult.updatedAccessories
        };
    }
    catch (error) {
        Logger.log("Error in returnRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}


// ============================================
// Dashboard Functions
// ============================================
/* 
@ Get Dashboard Stats
*/
function getDashboardStats() {
    const ss = getActiveSheet();
    const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
    const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
    const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);

    let totalItems = 0;
    let remainItems = 0;
    let totalAccessories = 0;
    let remainAccessories = 0;

    // itemCount => get total items where Active = TRUE
    const totalItemTypes = itemsSheet ? itemsSheet.getDataRange().getValues().filter((row, index) => {
        if (index === 0) return false;

        if(row[6] === true || row[6] === 'TRUE') {
            totalItems += row[3]; // Assuming 'Total_Qty' is in the 4th column (index 3)
            remainItems += row[4]; // Assuming 'Available_Qty' is in the 5th column (index 4)
        }

        return row[6] === true || row[6] === 'TRUE'; // Assuming 'Active' is in the 7th column (index 6)
    }).length : 0;

    // accessoriesCount => get total accessories where Active = TRUE
    const totalAccessoryTypes = accessorySheet ? accessorySheet.getDataRange().getValues().filter((row, index) => {
        if (index === 0) return false;

        if(row[5] === true || row[5] === 'TRUE') {
            totalAccessories += row[3]; // Total_Qty is in the 4th column (index 3)
            remainAccessories += row[4]; // Available_Qty is in the 5th column (index 4)
        }

        return row[5] === true || row[5] === 'TRUE'; // Active is in the 6th column (index 5)
    }).length : 0;

    const submitRequestCount = requestSheet ? requestSheet.getDataRange().getValues().filter((row, index) => {
        if (index === 0) return false;
        return row[2] === 'Submit'; // Assuming 'Status' is in the 3rd column (index 2)
    }).length : 0;

    const waitingReturnCount = requestSheet ? requestSheet.getDataRange().getValues().filter((row, index) => {
        if (index === 0) return false;
        return row[2] === 'Distributed'; // Assuming 'Status' is in the 3rd column (index 2)
    }).length : 0;

    const result = {
        success: true,
        data: {
            // items
            totalItemTypes: totalItemTypes,
            totalItems: totalItems,
            remainItems: remainItems,
            // accessories
            totalAccessoryTypes: totalAccessoryTypes,
            totalAccessories: totalAccessories,
            remainAccessories: remainAccessories,
            // request
            submitRequestCount: submitRequestCount,
            waitingReturnCount: waitingReturnCount
        }
    };

    console.log('[GAS] Dashboard Stats :', JSON.stringify(result));

    return result;
}

/*
@ Get Request Activity - Shows all request movements (All users)
*/
function getRequestActivity(page = 1, pageSize = 50) {
    try {
        const ss = getActiveSheet();
        const requestActivitySheet = ss.getSheetByName(REQUEST_ACTIVITY_SHEET_NAME);
        
        if (!requestActivitySheet) {
            return { success: true, data: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
        }
        
        const data = requestActivitySheet.getDataRange().getValues();
        const headers = data[0];
        
        // Get all request activities (no filtering needed since sheet is dedicated)
        const requestActivities = [];
        for (let i = 1; i < data.length; i++) {
            if (data[i][0]) { // Check if row has data
                requestActivities.push({
                    Log_Id: data[i][0],
                    Email: data[i][1],
                    Activity: data[i][2],
                    Action_At: formatDate(data[i][3])
                });
            }
        }
        
        // Sort by Action_At descending
        requestActivities.sort((a, b) => {
            const dateA = new Date(a.Action_At);
            const dateB = new Date(b.Action_At);
            return dateB - dateA;
        });
        
        // Pagination
        const total = requestActivities.length;
        const totalPages = Math.ceil(total / pageSize);
        const startIndex = (page - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const paginatedData = requestActivities.slice(startIndex, endIndex);
        
        console.log("Request Activities Paginated Data: " + JSON.stringify(paginatedData));

        return {
            success: true,
            data: paginatedData,
            total: total,
            page: page,
            pageSize: pageSize,
            totalPages: totalPages
        };
    } catch (error) {
        Logger.log("Error in getRequestActivity: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

/*
@ Get System Activity - Shows login/logout activities (Admin only)
*/
function getSystemActivity(page = 1, pageSize = 50) {
    try {
        const userEmail = Session.getActiveUser().getEmail();
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const systemActivitySheet = ss.getSheetByName(SYSTEM_ACTIVITY_SHEET_NAME);
        const sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);
        
        if (!systemActivitySheet) {
            return { success: true, data: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
        }
        
        const data = systemActivitySheet.getDataRange().getValues();
        const sessionData = sessionSheet ? sessionSheet.getDataRange().getValues() : [];
        
        // Get system activities from dedicated sheet
        const systemActivities = [];
        
        // From system activity log
        for (let i = 1; i < data.length; i++) {
            if (data[i][0]) { // Check if row has data
                systemActivities.push({
                    Log_Id: data[i][0],
                    Email: data[i][1],
                    Activity: data[i][2],
                    Action_At: formatDate(data[i][3]),
                    Type: 'System'
                });
            }
        }
        
        // From sessions (active sessions)
        for (let i = 1; i < sessionData.length; i++) {
            systemActivities.push({
                Log_Id: 'SESSION-' + sessionData[i][0], // Use string format to unify with Log_Id
                Email: sessionData[i][1],
                Activity: `Active Session - Permission: ${sessionData[i][2]}`,
                Action_At: formatDate(sessionData[i][4]),
                Type: 'Session'
            });
        }
        
        // Sort by Action_At descending
        systemActivities.sort((a, b) => {
            const dateA = new Date(a.Action_At);
            const dateB = new Date(b.Action_At);
            return dateB - dateA;
        });
        
        // Pagination
        const total = systemActivities.length;
        const totalPages = Math.ceil(total / pageSize);
        const startIndex = (page - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const paginatedData = systemActivities.slice(startIndex, endIndex);
        
        console.log("System Activities Paginated Data: " + JSON.stringify(paginatedData));
        return {
            success: true,
            data: paginatedData,
            total: total,
            page: page,
            pageSize: pageSize,
            totalPages: totalPages
        };
    } catch (error) {
        Logger.log("Error in getSystemActivity: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

/*
@ Get Inventory Activity - Shows item/accessory changes (Admin only)
*/
function getInventoryActivity(page = 1, pageSize = 50) {
    try {
        const userEmail = Session.getActiveUser().getEmail();
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const inventoryActivitySheet = ss.getSheetByName(INVENTORY_ACTIVITY_SHEET_NAME);
        
        if (!inventoryActivitySheet) {
            return { success: true, data: [], total: 0, page: 1, pageSize: pageSize, totalPages: 0 };
        }
        
        const data = inventoryActivitySheet.getDataRange().getValues();
        
        // Get all inventory activities (no filtering needed since sheet is dedicated)
        const inventoryActivities = [];
        for (let i = 1; i < data.length; i++) {
            if (data[i][0]) { // Check if row has data
                inventoryActivities.push({
                    Log_Id: data[i][0],
                    Email: data[i][1],
                    Activity: data[i][2],
                    Action_At: formatDate(data[i][3])
                });
            }
        }
        
        // Sort by Action_At descending
        inventoryActivities.sort((a, b) => {
            const dateA = new Date(a.Action_At);
            const dateB = new Date(b.Action_At);
            return dateB - dateA;
        });
        
        // Pagination
        const total = inventoryActivities.length;
        const totalPages = Math.ceil(total / pageSize);
        const startIndex = (page - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const paginatedData = inventoryActivities.slice(startIndex, endIndex);
        
        console.log("Inventory Activities Paginated Data: " + JSON.stringify(paginatedData));
        return {
            success: true,
            data: paginatedData,
            total: total,
            page: page,
            pageSize: pageSize,
            totalPages: totalPages
        };
    } catch (error) {
        Logger.log("Error in getInventoryActivity: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}


/**
@ Log Request Activity
*/
function logRequestActivity(actionType) {
    try {
        const ss = getActiveSheet();
        const requestActivitySheet = ss.getSheetByName(REQUEST_ACTIVITY_SHEET_NAME);
        if (!requestActivitySheet) {
            throw new Error(`Sheet "${REQUEST_ACTIVITY_SHEET_NAME}" not found`);
        }

        const email = Session.getActiveUser().getEmail();
        const timestamp = formatDate(new Date());
        const logId = getNextId(requestActivitySheet, 0);

        requestActivitySheet.appendRow([logId, email, actionType, timestamp]);

        Logger.log(`Request activity logged: ${email} - ${actionType} at ${timestamp}`);
    } catch (error) {
        Logger.log("Error in logRequestActivity: " + error.toString());
    }
}

/**
@ Log System Activity
*/
function logSystemActivity(actionType, userEmail = null) {
    try {
        const ss = getActiveSheet();
        const systemActivitySheet = ss.getSheetByName(SYSTEM_ACTIVITY_SHEET_NAME);
        if (!systemActivitySheet) {
            throw new Error(`Sheet "${SYSTEM_ACTIVITY_SHEET_NAME}" not found`);
        }

        const email = userEmail || Session.getActiveUser().getEmail();
        const timestamp = formatDate(new Date());
        const logId = getNextId(systemActivitySheet, 0);

        systemActivitySheet.appendRow([logId, email, actionType, timestamp]);

        Logger.log(`System activity logged: ${email} - ${actionType} at ${timestamp}`);
    } catch (error) {
        Logger.log("Error in logSystemActivity: " + error.toString());
    }
}

/**
@ Log Inventory Activity
*/
function logInventoryActivity(actionType) {
    try {
        const ss = getActiveSheet();
        const inventoryActivitySheet = ss.getSheetByName(INVENTORY_ACTIVITY_SHEET_NAME);
        if (!inventoryActivitySheet) {
            throw new Error(`Sheet "${INVENTORY_ACTIVITY_SHEET_NAME}" not found`);
        }

        const email = Session.getActiveUser().getEmail();
        const timestamp = formatDate(new Date());
        const logId = getNextId(inventoryActivitySheet, 0);

        inventoryActivitySheet.appendRow([logId, email, actionType, timestamp]);

        Logger.log(`Inventory activity logged: ${email} - ${actionType} at ${timestamp}`);
    } catch (error) {
        Logger.log("Error in logInventoryActivity: " + error.toString());
    }
}

// ============================================
// Helper Functions - Caching
// ============================================

const setCache = (key, items, durationSec) => {
    try {
        // if no key go to catch block
        if (!key)
            throw new Error("Cache key is required");
            
        const cache = CacheService.getScriptCache();
        const cacheKey = key;
        cache.put(cacheKey, JSON.stringify(items), durationSec);
    } catch (error) {
        Logger.log("Error caching items: " + error.toString());
        throw error;
    }
};

const getCachedItems = (key) => {
    try {
        const cache = CacheService.getScriptCache();
        const cacheKey = key;
        const cached = cache.get(cacheKey);
        return cached ? JSON.parse(cached) : null;
    } catch (error) {
        Logger.log("Error getting cached items: " + error.toString());
        throw error;
    }
};

const clearCache = (key) => {
    try {
        const cache = CacheService.getScriptCache();
        const cacheKey = key;
        cache.remove(cacheKey);
        Logger.log("Cache cleared successfully");
    } catch (error) {
        Logger.log("Error clearing cache: " + error.toString());
        throw error;
    }
};


// ============================================
// Helper Functions - Date Formatting
// ============================================
function formatDate(date, format = 'yyyy-MM-dd HH:mm:ss') {
    try {
        if (date instanceof Date && !isNaN(date)) 
            return Utilities.formatDate(date, Session.getScriptTimeZone(), format);

        return '';
    } catch (error) {
        Logger.log("Error formatting date: " + error.toString());
        throw error;
    }
}


function arrayParser(input) {
    // Normalize parameter (สำคัญมาก)
    if (!Array.isArray(input)) {
        input = Object.values(input || {});
    }

    input = input.map(id => String(id).trim());
    return input;
}

// ============================================
// Initial Data Import Function
// ============================================

/*
@ Initialize Data with Real Data from Sheets
@ This function clears old data and imports real items, accessories, and their mappings
@ Should be run only once for initial setup
*/
function initializeDataFromReference() {
    try {
        const userEmail = Session.getActiveUser().getEmail();
        
        // Check admin permission
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        // Ensure all sheets exist before proceeding
        repairSheets();
        
        const ss = getActiveSheet();
        const itemsSheet = ss.getSheetByName(ITEM_SHEET_NAME);
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        const itemAccessorySheet = ss.getSheetByName(ITEM_ACCESSORY_SHEET_NAME);
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        const stockLedgerSheet = ss.getSheetByName(STOCK_LEDGER_SHEET_NAME);
        const requestActivitySheet = ss.getSheetByName(REQUEST_ACTIVITY_SHEET_NAME);
        const systemActivitySheet = ss.getSheetByName(SYSTEM_ACTIVITY_SHEET_NAME);
        const inventoryActivitySheet = ss.getSheetByName(INVENTORY_ACTIVITY_SHEET_NAME);
        const sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);
        
        // Clear existing data from all sheets (keep headers)
        if (itemsSheet && itemsSheet.getLastRow() > 1) {
            itemsSheet.deleteRows(2, itemsSheet.getLastRow() - 1);
        }
        if (accessorySheet && accessorySheet.getLastRow() > 1) {
            accessorySheet.deleteRows(2, accessorySheet.getLastRow() - 1);
        }
        if (itemAccessorySheet && itemAccessorySheet.getLastRow() > 1) {
            itemAccessorySheet.deleteRows(2, itemAccessorySheet.getLastRow() - 1);
        }
        if (requestSheet && requestSheet.getLastRow() > 1) {
            requestSheet.deleteRows(2, requestSheet.getLastRow() - 1);
        }
        if (requestItemSheet && requestItemSheet.getLastRow() > 1) {
            requestItemSheet.deleteRows(2, requestItemSheet.getLastRow() - 1);
        }
        if (requestItemAccessorySheet && requestItemAccessorySheet.getLastRow() > 1) {
            requestItemAccessorySheet.deleteRows(2, requestItemAccessorySheet.getLastRow() - 1);
        }
        if (userSheet && userSheet.getLastRow() > 1) {
            userSheet.deleteRows(2, userSheet.getLastRow() - 1);
        }
        if (stockLedgerSheet && stockLedgerSheet.getLastRow() > 1) {
            stockLedgerSheet.deleteRows(2, stockLedgerSheet.getLastRow() - 1);
        }
        if (requestActivitySheet && requestActivitySheet.getLastRow() > 1) {
            requestActivitySheet.deleteRows(2, requestActivitySheet.getLastRow() - 1);
        }
        if (systemActivitySheet && systemActivitySheet.getLastRow() > 1) {
            systemActivitySheet.deleteRows(2, systemActivitySheet.getLastRow() - 1);
        }
        if (inventoryActivitySheet && inventoryActivitySheet.getLastRow() > 1) {
            inventoryActivitySheet.deleteRows(2, inventoryActivitySheet.getLastRow() - 1);
        }
        if (sessionSheet && sessionSheet.getLastRow() > 1) {
            sessionSheet.deleteRows(2, sessionSheet.getLastRow() - 1);
        }
        
        const timestamp = new Date();
        let itemsImported = 0;
        let accessoriesImported = 0;
        let mappingsImported = 0;
        let usersImported = 0;
        
        // Real Items Data from InventoryManagement - Items.csv (49 items)
        const realItems = [
            ["841 Titrando", "2.841.0010 (เสีย)", 1, 1, "https://image2url.com/r2/default/images/1770110030781-3b71ec6a-8388-44b3-9596-c990f34d08c8.png"],
            ["907 Titrando", "2.907.0010", 1, 1, "https://image2url.com/r2/default/images/1770111189839-f57eadea-d8f7-4b94-b143-44fabb67fa84.png"],
            ["808 Titrando", "2.808.0010", 1, 1, "https://image2url.com/r2/default/images/1770086412243-2f877e15-cec7-4120-bf1e-b3b7abec07cf.png"],
            ["852 Titrando", "2.852.0010", 1, 1, "https://image2url.com/r2/default/images/1770110600497-eb4fe05b-c636-45c1-8e49-8ebc28558c7c.png"],
            ["905 Titrando", "2.905.0010", 1, 1, "https://image2url.com/r2/default/images/1770111364183-40464f56-82bf-4197-97b8-06e18af778b5.png"],
            ["801 stirrer", "2.801.0010", 1, 1, "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTe5Bn4YioHHchJzvH3cDtShPfcJM_k1U9m-g&s"],
            ["916 Ti-touch", "2.916.0010", 2, 2, "https://image2url.com/r2/default/images/1770170012397-d921dca7-55a3-4d90-b9e5-5a1f158fd98d.png"],
            ["Eco Titrator", "2.1008.0010", 2, 2, "https://metrohm.scene7.com/is/image/metrohm/4840?$xh-544$&bfc=on"],
            ["904 Titrando", "2.904.0010", 1, 1, "https://image2url.com/r2/default/images/1770172036979-d384469b-d3b1-47ab-9ddd-b806e736b998.png"],
            ["900 Touch control", "2.900.0010", 1, 1, "https://static-data2.manualslib.com/product-images/1b6/2342664/metrohm-900-touch-control.jpg"],
            ["Eco Dosimat", "2.1007.0010", 1, 1, "https://image2url.com/r2/default/images/1770169561855-e1213163-8b8d-45b9-9677-f73dbcf151da.png"],
            ["859 Titrothrem", "2.859.0010", 1, 1, "https://image2url.com/r2/default/images/1770170225456-4d9376fd-0842-46c9-b322-ea5c1c2dde34.png"],
            ["800 Dosino", "2.800.0010", 4, 4, "https://metrohm.scene7.com/is/image/metrohm/2127_s?$xh-1280$&bfc=on"],
            ["803 Ti Stand", "2.803.0010", 1, 1, "https://image2url.com/r2/default/images/1770085759392-64db9849-5016-4a77-aca5-99c7e0837be4.jpg"],
            ["805 Dosimat", "2.805.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1378_s?$xh-1280$&bfc=on"],
            ["802 Rod stirrer", "2.802.0010", 6, 6, "https://images.offerup.com/8Zo1b7oifuxDUjKDMQmM8GNXAJI=/1440x1920/ebd7/ebd7f5271c354489a4acff5019a9cb09.jpg"],
            ["728 stirrer without stand", "2.728.0010", 1, 1, "https://image2url.com/r2/default/images/1770174227483-8d86ddd5-d534-4235-9efd-1275d23507dd.png"],
            ["Omnis Dosing Module without stirrer", "2.1003.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/4249?$xh-1280$&bfc=on"],
            ["Omnis Dosing Module with stirrer", "2.1003.0110", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/4250?$xh-1280$&bfc=on"],
            ["OMNIS sample robot", "", 1, 1, "https://image2url.com/r2/default/images/1770170628931-928392be-245f-4902-9680-de6288be283a.png"],
            ["OMNIS Titrator with stirrer", "2.1001.0020", 2, 2, "https://metrohm.scene7.com/is/image/metrohm/4248?$xh-544$&bfc=on"],
            ["OMNIS coulometer with stirrer", "2.1018.0020", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/210180020_FV?$xh-1280$&bfc=on"],
            ["Eco KF Titrator", "2.1027.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1102800010?$xh-544$&bfc=on"],
            ["860 KF Thermoprep", "2.860.0010", 1, 1, "https://image2url.com/r2/default/images/1770111576358-30724e7c-556c-4633-ae0a-559a9c919c1c.png"],
            ["832 KF Thermoprep", "2.832.0010", 1, 1, "https://image2url.com/r2/default/images/1770111749313-ba0cfa3a-5dd1-4e3d-9749-82f21bc68bd3.png"],
            ["899 Coulometer", "2.899.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/3553_s?$xh-1280$&bfc=on"],
            ["756 KF Coulometer", "2.756.0010", 1, 1, "https://www.ntc-tech.com/cdn/shop/products/ukRoN5p-m0Hh60lV7KpDtMODpG0qG3lLXohGGtdK0_BXveWcrQ6OJ21Hu23eLWkYIy7FlFHCNh0s5jcP0NDfn_asIvdDo4k_s4470.jpg"],
            ["831 KF Coulometer", "2.831.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1498_s?$xh-1280$&bfc=on"],
            ["808 Touch control", "2.808.0010", 1, 1, "https://image2url.com/r2/default/images/1770112064025-95265755-5393-4adf-9cd6-bf611c96837f.png"],
            ["885 compact oven sc", "2.885.0010", 1, 1, "https://image2url.com/r2/default/images/1770111926027-af5b978d-fc87-4772-b245-5241839c1949.png"],
            ["915 KF Ti-Touch", "2.915.0010", 1, 1, "https://image2url.com/r2/default/images/1770112334125-53b1d506-d9ac-4d0f-b7e4-c5e13f15bd79.png"],
            ["862 compact Titrosampler", "2.862.0010", 1, 1, "https://image2url.com/r2/default/images/1770112210246-b11ec0de-4491-4454-bbd0-dddc319d1089.png"],
            ["870 KF Titrino plus", "2.870.0010", 1, 1, "https://image2url.com/r2/default/images/1770112548836-996ed91b-afd7-4817-bbf3-251d99a4cd5f.png"],
            ["848 Titrino plus", "2.848.0010", 1, 1, "https://image2url.com/r2/default/images/1770112869227-0c00e97c-21e7-40e0-8622-52c67403ca37.png"],
            ["Eco coulometer", "2.1028.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1102800010?$xh-544$&bfc=on"],
            ["solvent pump", "2.1029.0010", 1, 1, "https://image2url.com/r2/default/images/1770169307529-2ce27fd1-14b2-4ecb-a416-57e0751a1bb7.png"],
            ["914 pH/conductomer", "2.914.0020", 1, 1, "https://image2url.com/r2/default/images/1770113012483-86a1e09d-9c7c-4f36-96bc-7baca9036ebc.png"],
            ["912 Conductometer", "2.912.0010", 1, 1, "https://image2url.com/r2/default/images/1770113078469-c9f491b4-5594-4c9b-940e-c60e4d079b6a.png"],
            ["913 pH meter", "2.913.0010", 1, 1, "https://image2url.com/r2/default/images/1770113144119-c332b957-6231-44d8-80ae-1c3135995589.png"],
            ["867 pH module", "2.867.0010", 1, 1, "https://image2url.com/r2/default/images/1770109167866-45310e93-f7d2-4d57-8685-d18b2e904eaf.jpg"],
            ["826 pH mobile", "2.826.0010", 1, 1, "https://image2url.com/r2/default/images/1770109355002-ac532c7b-6458-48e7-8482-7173c71a13f6.jpg"],
            ["949 pH meter", "2.949.0010", 1, 1, "https://image2url.com/r2/default/images/1770113371736-0201045c-5bf5-4d6e-8012-c8629166b787.png"],
            ["827 pH lab", "2.827.0010", 1, 1, "https://image2url.com/r2/default/images/1770113243494-dfc740ae-32d1-4ed2-b2ae-412c963112be.png"],
            ["781 pH/Ion Meter", "2.781.0010", 1, 1, "https://image2url.com/r2/default/images/1770173212831-2a45cab4-a75b-4db7-aa9f-808db2804ea6.png"],
            ["856 Conductivity module", "2.856.0010", 1, 1, "https://image2url.com/r2/default/images/1770173367571-183417fa-9ccc-410d-9407-183f29ec9a6d.png"],
            ["704 pH meter", "2.704.0010", 1, 1, "https://image2url.com/r2/default/images/1770173654721-4936e03b-6816-40db-8b3b-0fd36e32e58f.png"],
            ["890 Titrando", "2.890.0010", 1, 1, "https://image2url.com/r2/default/images/1770172406088-5e939ffb-2e2b-4218-908e-ca983823dbe1.png"],
            ["892 Professional Rancimat", "2.892.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/3571_s?$xh-1280$&bfc=on"],
            ["743 Rancimat", "2.743.0010", 1, 1, "https://image2url.com/r2/default/images/1770171916935-bac81e4e-8b2a-4ea8-8c17-96d6f7de664d.png"]
        ];
        
        // Import Items
        for (let i = 0; i < realItems.length; i++) {
            const item = realItems[i];
            const itemId = i + 1;
            
            itemsSheet.appendRow([
                itemId,
                item[0],      // Item_Name
                item[1],      // Item_Desc
                item[2],      // Total_Qty
                item[3],      // Available_Qty
                item[4],      // Image
                true,         // Active
                "system",     // Created_By
                timestamp,    // Created_At
                "system",     // Modified_By
                timestamp     // Modified_At
            ]);
            
            itemsImported++;
        }
        
        // Real Accessories Data from InventoryManagement - Accessories.csv (30 accessories)
        // Note: Filtered out empty rows (IDs 10-55, 61, 65, 72-73) that had no name/data
        const realAccessories = [
            ["stirrering proppeller 104mm", "", 1, 1],
            ["stirrering prop. intensive 104mm", "", 1, 1],
            ["stirrering propeller 94mm", "", 0, 0],
            ["Dosing unit 2 mL", "", 0, 0],
            ["Dosing unit 10 mL", "", 0, 0],
            ["Dosing unit 20 mL", "", 0, 0],
            ["Dosing unit 50 mL", "", 0, 0],
            ["OMNIS Dosing 10 mL", "", 0, 0],
            ["OMNIS Dosing 20 mL", "", 0, 0],
            ["Dosing unit 10 mL", "6.1580.210", 5, 3],
            ["Dosing unit 20 mL", "6.1580.220", 2, 2],
            ["Dosing unit 2 mL", "6.1580.120", 2, 2],
            ["Dosing unit 50 mL", "6.1580.250", 2, 2],
            ["Holding clip for bottles", "6.2043.005", 24, 23],
            ["Power cable", "สายPower ต่อกับเครื่อง", 38, 37],
            ["Controller cable", "สาย controller เชื่อม software กับตัวเครื่อง", 4, 3],
            ["SET PC OMNIS", "CPU, Keyboard, mouse, power cable, Screen", 1, 0],
            ["Rod stand + lock camp", "แท่งเหล็ก + ตัวล็อค", 39, 38],
            ["Magnetic bar", "มีน้อยใช้สอยอย่างประหยัด", 14, 13],
            ["Electrode holder", "6.2021.020 Electrode holder for 4 electrodes and 2 buret tips", 6, 5],
            ["SET Tubing for Dosing unit", "1 Tubing, Tip, Microvalve", 10, 9],
            ["Brown glass bottle", "ขวดเปล่า", 15, 14],
            ["Electrod cable /1 m /F", "", 11, 10],
            ["2 Lan cable + Hub box + 1 Hub Power cacle", "สายแลนเหลือง 2 + กล่อง Hub + สายชาร์จ Hub", 1, 1],
            ["OMNIS Dosing unit 10 mL", "6.01508.210", 3, 3],
            ["OMNIS Dosing unit 20 mL", "6.01508.220", 2, 2],
            ["OMNIS Holder", "", 3, 3],
            ["OMNIS Molecular sieve with Cap", "", 3, 3],
            ["Red cable OMNIS", "สายสีแดงใช้กับ Temp senser to plug F", 1, 1],
            ["Blue cable OMNIS", "สายสีน้ำเงินใช้กับ Polarized electrode to plug F (เป็นขั้วโลหะและเป็นแง่ง หรือ Pt sheet)", 1, 1],
            ["Green cable OMNIS", "สายสีเขียวใช้กับ Electrode ทั่วไป to plug F (ทุกประเภท General)", 1, 1]
        ];
        
        // Import Accessories
        for (let i = 0; i < realAccessories.length; i++) {
            const accessory = realAccessories[i];
            const accessoryId = i + 1;
            
            accessorySheet.appendRow([
                accessoryId,
                accessory[0], // Accessory_Name
                accessory[1], // Accessory_Desc
                accessory[2], // Total_Qty
                accessory[3], // Available_Qty
                true,         // Active
                "system",     // Created_By
                timestamp,    // Created_At
                "system",     // Modified_By
                timestamp     // Modified_At
            ]);
            
            accessoriesImported++;
        }
        
        // Real Item-Accessory Mappings from InventoryManagement - Item_Accessory_Mapping.csv (50 mappings)
        const realMappings = [
            [1, 1], [1, 2], [1, 4], [2, 1], [2, 2], [2, 5], [3, 1], [3, 3],
            [4, 1], [4, 2], [5, 1], [5, 2], [6, 1], [6, 3], [7, 1], [7, 2],
            [8, 4], [8, 5], [9, 1], [9, 2], [10, 1], [11, 4], [11, 5], [12, 1],
            [13, 4], [13, 5], [13, 6], [14, 1], [15, 5], [15, 6], [16, 1], [16, 3],
            [17, 1], [17, 3], [18, 8], [18, 9], [19, 8], [19, 9], [19, 1], [20, 8],
            [21, 8], [21, 9], [21, 1], [22, 1], [23, 4], [24, 4], [25, 4], [26, 4],
            [27, 4], [28, 4]
        ];
        
        // Import Mappings
        for (let i = 0; i < realMappings.length; i++) {
            const mapping = realMappings[i];
            const mappingId = i + 1;
            
            itemAccessorySheet.appendRow([
                mappingId,
                mapping[0],   // Item_Id
                mapping[1],   // Accessory_Id
                "system",     // Created_By
                timestamp,    // Created_At
                true          // Active
            ]);
            
            mappingsImported++;
        }
        
        // Initial Users Data (4 columns: Email, Password, Permission, Active)
        const initialUsers = [
            ["modmastei2@gmail.com", "b7766cf93f0fcbcfa13adcc202419a4e5f21816f70360b6e76fd18342b56fdd8", "Admin", true],
            ["admin@admin.com", "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9", "Admin", true]
        ];
        
        // Import Users
        for (let i = 0; i < initialUsers.length; i++) {
            const user = initialUsers[i];
            
            userSheet.appendRow([
                user[0],      // Email
                user[1],      // Password (hashed)
                user[2],      // Permission
                user[3]       // Active
            ]);
            
            usersImported++;
        }
        
        // Log the import
        logSystemActivity(`Initialized real data: ${itemsImported} items, ${accessoriesImported} accessories, ${mappingsImported} mappings, ${usersImported} users`, 'system');
        
        return {
            success: true,
            message: "Real data initialized successfully from sheets folder",
            itemsImported: itemsImported,
            accessoriesImported: accessoriesImported,
            mappingsImported: mappingsImported,
            usersImported: usersImported
        };
    } catch (error) {
        Logger.log("Error in initializeDataFromReference: " + error.toString());
        return { success: false, message: error.toString() };
    }
}