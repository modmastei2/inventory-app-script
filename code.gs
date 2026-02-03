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
const REQUEST_SHEET_NAME = "Requests";
const REQUEST_ITEM_SHEET_NAME = "Request_Item";
const REQUEST_ITEM_ACCESSORY_SHEET_NAME = "Request_Item_Accessory";
const STOCK_LEDGER_SHEET_NAME = "StockLedger";
const USER_SHEET_NAME = "Users";
const AUDIT_LOG_SHEET_NAME = "Audit_Log";
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
  let requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
  let requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
  let requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);
  let stockLedgerSheet = ss.getSheetByName(STOCK_LEDGER_SHEET_NAME);
  let auditLogSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
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
        accessorySheet.getRange(1,1,1,11).setValues([
            ["Accessory_Id", "Item_Id","Accessory_Name","Accessory_Desc","Total_Qty","Available_Qty","Active","Created_By","Created_At","Modified_By","Modified_At"]
        ])

        Logger.log(`Created sheet: ${ACCESSORY_SHEET_NAME}`);
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

    if (!auditLogSheet) {
        auditLogSheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME);
        auditLogSheet.getRange(1,1,1,4).setValues([
            ["Log_Id","Email","Activity","Action_At"]
        ])
        Logger.log(`Created sheet: ${AUDIT_LOG_SHEET_NAME}`);
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
        
        auditLog(`Admin ${userEmail} created item: ${itemData.Item_Name} (ID: ${itemId})`);
        
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
                
                auditLog(`Admin ${userEmail} updated item ID: ${itemId}`);
                
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
                auditLog(`Admin ${userEmail} deleted item ID: ${itemId}`);
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
@ Create Accessory (Admin only)
*/
function createAccessory(accessoryData, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        
        if (!accessorySheet) {
            throw new Error(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
        }
        
        const accessoryId = getNextId(accessorySheet, 0);
        const timestamp = new Date();
        
        accessorySheet.appendRow([
            accessoryId,
            accessoryData.Item_Id,
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
        
        auditLog(`Admin ${userEmail} created accessory: ${accessoryData.Accessory_Name} (ID: ${accessoryId})`);
        
        return { success: true, message: "Accessory created successfully", accessoryId: accessoryId };
    } catch (error) {
        Logger.log("Error in createAccessory: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update Accessory (Admin only)
*/
function updateAccessory(accessoryId, accessoryData, userEmail) {
    try {
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "Unauthorized: Admin permission required" };
        }
        
        const ss = getActiveSheet();
        const accessorySheet = ss.getSheetByName(ACCESSORY_SHEET_NAME);
        
        if (!accessorySheet) {
            throw new Error(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
        }
        
        const data = accessorySheet.getDataRange().getValues();
        const headers = data[0];
        
        // Find accessory row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == accessoryId) {
                const rowNum = i + 1;
                const timestamp = new Date();
                
                // Update fields
                if (accessoryData.Item_Id !== undefined) accessorySheet.getRange(rowNum, 2).setValue(accessoryData.Item_Id);
                if (accessoryData.Accessory_Name !== undefined) accessorySheet.getRange(rowNum, 3).setValue(accessoryData.Accessory_Name);
                if (accessoryData.Accessory_Desc !== undefined) accessorySheet.getRange(rowNum, 4).setValue(accessoryData.Accessory_Desc);
                if (accessoryData.Total_Qty !== undefined) accessorySheet.getRange(rowNum, 5).setValue(accessoryData.Total_Qty);
                if (accessoryData.Available_Qty !== undefined) accessorySheet.getRange(rowNum, 6).setValue(accessoryData.Available_Qty);
                if (accessoryData.Active !== undefined) accessorySheet.getRange(rowNum, 7).setValue(accessoryData.Active);
                
                accessorySheet.getRange(rowNum, 10).setValue(userEmail);
                accessorySheet.getRange(rowNum, 11).setValue(timestamp);
                
                auditLog(`Admin ${userEmail} updated accessory ID: ${accessoryId}`);
                
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
        
        if (!accessorySheet) {
            throw new Error(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
        }
        
        const data = accessorySheet.getDataRange().getValues();
        
        // Find and delete accessory row
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] == accessoryId) {
                accessorySheet.deleteRow(i + 1);
                auditLog(`Admin ${userEmail} deleted accessory ID: ${accessoryId}`);
                return { success: true, message: "Accessory deleted successfully" };
            }
        }
        
        return { success: false, message: "Accessory not found" };
    } catch (error) {
        Logger.log("Error in deleteAccessory: " + error.toString());
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
        
        auditLog(`Admin ${adminEmail} created user: ${normalizedEmail}`);
        
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
                
                auditLog(`Admin ${adminEmail} updated user: ${originalEmail}`);
                
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
                
                auditLog(`Admin ${adminEmail} ${newActiveState ? 'activated' : 'deactivated'} user: ${data[i][0]}`);
                
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
                
                auditLog(`Admin ${adminEmail} changed password for user: ${data[i][0]}`);
                
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

            if (!accessorySheet) {
                Logger.log(`Sheet "${ACCESSORY_SHEET_NAME}" not found`);
                return { data: [] };
            }

            const data = accessorySheet.getDataRange().getValues();
            
            if (data.length === 0) {
                return { data: [] };
            }

            const headers = data.shift(); // Remove header row

            // Convert to objects
            accessories = data
                .filter(row => row.some(cell => cell !== "")) // Filter empty rows
                .map(row => {
                    const accessory = {
                        Accessory_Id: row[0],
                        Item_Id: row[1],	
                        Accessory_Name: row[2],
                        Accessory_Desc: row[3],
                        Total_Qty: row[4],
                        Available_Qty: row[5],
                        Active: row[6],
                        Created_By: row[7],
                        Created_At: formatDate(row[8]),
                        Modified_By: row[9],
                        Modified_At: formatDate(row[10])
                    };

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
                itemIds.includes(String(acc.Item_Id).trim())
            );
        } else {
            // For normal usage - show only active accessories
            accessories = accessories.filter(acc =>
                itemIds.includes(String(acc.Item_Id).trim()) &&
                (acc.Active === true || String(acc.Active).toUpperCase() === "TRUE")
            );
        }

        const result = {
            data: accessories
        }

       console.log("[GAS] Load Accessories :", result)
                
        return result;

    }
    catch (error) {
        Logger.log("Error in loadAccessories: " + error.toString());
        return { data: [] };
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
        // 1. ดึงข้อมูล Request หลัก
        // ============================================
        if (submitRequestData.IsNew) {
            let request = [];
            let requestId = getNextId(requestSheet, 0);
            let status = 'Submit';
            const timestamp = formatDate(new Date());
            const userName = submitRequestData.Requirer_Name;

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
                userName,
                timestamp,
                userName,
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

            auditLog("Submit Request ID " + requestId);
        } else {
            auditLog("Edit Request ID " + submitRequestData.Request_Id);
        }

        Logger.log("Submit Request Data: " + JSON.stringify(submitRequestData));

        return { success: true, message: "Request submitted successfully" };
    }
    catch (error) {
        Logger.log("Error in submitRequest: " + error.toString());
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
@ Distribute Request
*/
function distributeRequest(request) {
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
        requestSheet.getRange(rowIndex, modifiedByIndex + 1).setValue(Session.getActiveUser().getEmail());
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

        auditLog("Distribute Request ID " + request.Request_Id);

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

        if (rowData[statusIndex] !== "Distributed") {
            throw new Error(`Request ID ${requestId} is not in 'Distributed' status`);
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

        auditLog("Cancel Request ID " + requestId);

        return { success: true, message: `Request ID ${requestId} cancelled successfully` };
    }
    catch (error) {
        Logger.log("Error in cancelRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Return Request: Support partial return
*/
function returnRequest(returnRequestData) {
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
        const returnSummary = isPartialReturn ? 
            `Partial return - Request ID ${returnRequestData.Request_Id}` : 
            `Full return - Request ID ${returnRequestData.Request_Id}`;
        auditLog(returnSummary);

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
        throw new Error(error.toString());
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
    const auditLogSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);

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

        if(row[6] === true || row[6] === 'TRUE') {
            totalAccessories += row[4]; // Assuming 'Total_Qty' is in the 5th column (index 4)
            remainAccessories += row[5]; // Assuming 'Available_Qty' is in the 6th column (index 5)
        }

        return row[6] === true || row[6] === 'TRUE'; // Assuming 'Active' is in the 7th column (index 6)
    }).length : 0;

    const submitRequestCount = requestSheet ? requestSheet.getDataRange().getValues().filter((row, index) => {
        if (index === 0) return false;
        return row[2] === 'Submit'; // Assuming 'Status' is in the 3rd column (index 2)
    }).length : 0;

    const waitingReturnCount = requestSheet ? requestSheet.getDataRange().getValues().filter((row, index) => {
        if (index === 0) return false;
        return row[2] === 'Distributed'; // Assuming 'Status' is in the 3rd column (index 2)
    }).length : 0;

    const top50_recentlyActivity = auditLogSheet.getDataRange().getValues().slice(1) // Skip header row
        .sort((a, b) => new Date(b[3]) - new Date(a[3])) // Sort by Action_At descending
        .slice(0, 50)


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
            waitingReturnCount: waitingReturnCount,
            top50_recentlyActivity: top50_recentlyActivity.map((row, index) => ({
                Log_Id: index + 1,
                Email: row[1],
                Activity: row[2],
                Action_At: formatDate(row[3])
            }))
        }
    };

    console.log('[GAS] Dashboard Stats :', JSON.stringify(result));

    return result;
}


/**
@ Audit Log
*/
function auditLog(actionType) {
    try {
        const ss = getActiveSheet();
        const auditLogSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
        if (!auditLogSheet) {
            throw new Error(`Sheet "${AUDIT_LOG_SHEET_NAME}" not found`);
        }

        const email = Session.getActiveUser().getEmail();
        const timestamp = formatDate(new Date());
        const logId = getNextId(auditLogSheet, 0);

        auditLogSheet.appendRow([logId, email, actionType, timestamp]);

        Logger.log(`Audit log recorded: ${email} - ${actionType} at ${timestamp}`);
    } catch (error) {
        Logger.log("Error in auditLog: " + error.toString());
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