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
@ Get Authorized User and Permission by Session
@ Supports shared/group email by using sessionId instead of Session.getActiveUser()
*/
function getAuthorized(sessionId){
    const email = getCurrentUserEmail(sessionId);
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

/*
@ Get Current User Email from Session
@ This function retrieves the email associated with a specific sessionId
@ This supports shared/group emails where multiple users can log in with the same email
*/
function getCurrentUserEmail(sessionId) {
    try {
        const ss = getActiveSheet();
        const sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);

        if (!sessionSheet) {
            throw new Error(`Sheet "${SESSION_SHEET_NAME}" not found`);
        }

        const sessionData = sessionSheet.getDataRange().getValues();
        const sessionHeaders = sessionData[0];
        const sessionIdCol = sessionHeaders.indexOf("Session_Id");
        const emailCol = sessionHeaders.indexOf("Email");

        // Find session by sessionId
        for (let i = 1; i < sessionData.length; i++) {
            if (sessionData[i][sessionIdCol] == sessionId) {
                return sessionData[i][emailCol];
            }
        }

        throw new Error("ไม่พบ session หรือ session หมดอายุ");
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน getCurrentUserEmail: " + error.toString());
        throw error;
    }
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
    
    console.log('Checking admin permission for email:', email);

    const userData = userSheet.getDataRange().getValues();
    const headers = userData[0];
    const emailCol = headers.indexOf("Email");
    const permissionCol = headers.indexOf("Permission");
    const activeCol = headers.indexOf("Active");
    
    for (let i = 1; i < userData.length; i++) {
        if (userData[i][emailCol] === email) {
            const isActive = userData[i][activeCol] === true || userData[i][activeCol] === 'TRUE';
            console.log('Matching with admin email:', userData[i][emailCol], 'Permission:', userData[i][permissionCol], 'Active:', isActive);
            return userData[i][permissionCol] === 'Admin' && isActive;
        }
    }
    
    return false;
}

/*
@ Create Item (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function createItem(itemData, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
            itemData.Total_Qty || 0, //itemData.Available_Qty || 0, // on create item Available_Qty = Total_Qty
            itemData.Image || '',
            itemData.Active !== false ? true : false,
            userEmail,
            timestamp,
            userEmail,
            timestamp
        ]);
        
        logInventoryActivity(`Admin ${userEmail} created item: ${itemData.Item_Name} (ID: ${itemId})`, sessionId);
        
        return { success: true, message: "สร้างรายการ Item สำเร็จ", itemId: itemId };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน createItem: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update Item (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function updateItem(itemId, itemData, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
                
                // find diff between old Total_Qty and itemData.Total_Qty
                // in case 
                // total item 15 
                // borrow 1 
                // available 14
                // and reduce total
                const diffTotalQty = itemData.Total_Qty - data[i][3]; // 14 - 15 = -1
                const updateAvailableQty = data[i][4] + diffTotalQty; // 14 + (-1) = 13

                // Update fields
                if (itemData.Item_Name !== undefined) itemsSheet.getRange(rowNum, 2).setValue(itemData.Item_Name);
                if (itemData.Item_Desc !== undefined) itemsSheet.getRange(rowNum, 3).setValue(itemData.Item_Desc);
                if (itemData.Total_Qty !== undefined) itemsSheet.getRange(rowNum, 4).setValue(itemData.Total_Qty);
                itemsSheet.getRange(rowNum, 5).setValue(updateAvailableQty);
                if (itemData.Image !== undefined) itemsSheet.getRange(rowNum, 6).setValue(itemData.Image);
                if (itemData.Active !== undefined) itemsSheet.getRange(rowNum, 7).setValue(itemData.Active);
                
                itemsSheet.getRange(rowNum, 10).setValue(userEmail);
                itemsSheet.getRange(rowNum, 11).setValue(timestamp);
                
                logInventoryActivity(`Admin ${userEmail} updated item ID: ${itemId}`, sessionId);
                
                return { success: true, message: "แก้ไขรายการ Item สำเร็จ" };
            }
        }
        
        return { success: false, message: "ไม่พบรายการ Item" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน updateItem: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Delete Item (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function deleteItem(itemId, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
                logInventoryActivity(`Admin ${userEmail} deleted item ID: ${itemId}`, sessionId);
                return { success: true, message: "ลบรายการ Item สำเร็จ" };
            }
        }
        
        return { success: false, message: "ไม่พบรายการ Item" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน deleteItem: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Create Accessory (Admin only) - Now supports many-to-many item mapping
@ Now uses sessionId to support shared/group emails
*/
function createAccessory(accessoryData, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
            accessoryData.Total_Qty || 0, // accessoryData.Available_Qty || 0, // on create accessory Available_Qty = Total_Qty
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
        
        logInventoryActivity(`Admin ${userEmail} created accessory: ${accessoryData.Accessory_Name} (ID: ${accessoryId})`, sessionId);
        
        return { success: true, message: "สร้างรายการ Accessory สำเร็จ", accessoryId: accessoryId };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน createAccessory: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update Accessory (Admin only) - Now supports updating item mappings
@ Now uses sessionId to support shared/group emails
*/
function updateAccessory(accessoryId, accessoryData, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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

                // find diff between old Total_Qty and accessoryData.Total_Qty
                // total item 24 
                // borrow 3
                // available 21
                // and reduce total
                const diffTotalQty = accessoryData.Total_Qty - data[i][3]; // 27 - 24 = 3
                const updateAvailableQty = data[i][4] + diffTotalQty; // 21 + 3 = 24
                
                // Update fields (note: column indices changed after removing Item_Id)
                if (accessoryData.Accessory_Name !== undefined) accessorySheet.getRange(rowNum, 2).setValue(accessoryData.Accessory_Name);
                if (accessoryData.Accessory_Desc !== undefined) accessorySheet.getRange(rowNum, 3).setValue(accessoryData.Accessory_Desc);
                if (accessoryData.Total_Qty !== undefined) accessorySheet.getRange(rowNum, 4).setValue(accessoryData.Total_Qty);
                accessorySheet.getRange(rowNum, 5).setValue(updateAvailableQty);
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
                
                logInventoryActivity(`Admin ${userEmail} updated accessory ID: ${accessoryId}`, sessionId);
                
                return { success: true, message: "แก้ไขรายการ Accessory สำเร็จ" };
            }
        }
        
        return { success: false, message: "ไม่พบรายการ Accessory" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน updateAccessory: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Delete Accessory (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function deleteAccessory(accessoryId, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
                
                logInventoryActivity(`Admin ${userEmail} deleted accessory ID: ${accessoryId}`, sessionId);
                return { success: true, message: "ลบรายการ Accessory สำเร็จ" };
            }
        }
        
        return { success: false, message: "ไม่พบรายการ Accessory" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน deleteAccessory: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน getItemAccessoryMappings: " + error.toString());
        return { success: false, message: error.toString(), mappings: [] };
    }
}

/*
@ Link Accessory to Items (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function linkAccessoryToItems(accessoryId, itemIds, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
        
        logInventoryActivity(`Admin ${userEmail} linked accessory ${accessoryId} to ${itemIds.length} items`, sessionId);
        
        return { success: true, message: "เชื่อมโยงอุปกรณ์เสริมสำเร็จ" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน linkAccessoryToItems: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน getUsers: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

/*
@ Create User (Admin only)
*/
function createUser(userData, sessionId) {
    try {
        const adminEmail = getCurrentUserEmail(sessionId);
        
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
        
        return { success: true, message: "สร้างผู้ใช้สำเร็จ" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน createUser: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Update User (Admin only)
*/
function updateUser(originalEmail, userData, sessionId) {
    try {
        const adminEmail = getCurrentUserEmail(sessionId);

        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
                
                return { success: true, message: "อัปเดตผู้ใช้สำเร็จ" };
            }
        }
        
        return { success: false, message: "ไม่พบผู้ใช้" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน updateUser: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Toggle User Active Status (Admin only)
*/
function toggleUserActive(userEmail, newActiveState, sessionId) {
    try {
        const adminEmail = getCurrentUserEmail(sessionId);
        
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
                
                return { success: true, message: `ผู้ใช้${newActiveState ? 'เปิดใช้งาน' : 'ปิดใช้งาน'}สำเร็จ` };
            }
        }
        
        return { success: false, message: "ไม่พบผู้ใช้" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน toggleUserActive: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Change User Password (Admin only)
*/
function changeUserPassword(userEmail, newPassword, sessionId) {
    try {
        const adminEmail = getCurrentUserEmail(sessionId);
        
        if (!checkAdminPermission(adminEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
        }
        
        // Validate new password
        if (!newPassword || newPassword.trim().length < 6) {
            return { success: false, message: "รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร" };
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
                
                return { success: true, message: "เปลี่ยนรหัสผ่านสำเร็จ" };
            }
        }
        
        return { success: false, message: "ไม่พบผู้ใช้" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน changeUserPassword: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน loadItems: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน loadAccessories: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน getRowRequest: " + error.toString());
        throw error;
    }
}

/*
@ Submit Request
@ Now uses sessionId to support shared/group emails
*/
function submitRequest(submitRequestData, sessionId) {
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
            const userEmail = getCurrentUserEmail(sessionId);

            // Check stock availability first (before any writes)
            const stockValidation = validateStockAvailable(submitRequestData.Items);

            if (!stockValidation.success) {
                throw new Error(`รายการไม่เพียงพอ: ${stockValidation.message}`);
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

            logRequestActivity("Create Request", sessionId);
        } else {
            // Edit existing request - only allowed if status is 'Submit' (Pending)
            const result = editRequest(submitRequestData, sessionId);
            if (!result.success) {
                throw new Error(result.message);
            }
        }

        Logger.log("Submit Request Data: " + JSON.stringify(submitRequestData));

        return { success: true, message: "ส่งคำขอสำเร็จ" };
    }
    catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน submitRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Edit Request - Only allowed for requests with status 'Submit' (Pending)
@ Users can edit their own requests, admins can edit any
@ Now uses sessionId to support shared/group emails
*/
function editRequest(editRequestData, sessionId) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
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
            return { success: false, message: "ไม่พบคำขอ" };
        }

        // Check if request is in editable status
        if (rowData[statusIndex] !== "Submit") {
            return { success: false, message: "สามารถแก้ไขได้เฉพาะคำขอที่ Submit เท่านั้น" };
        }

        // Check permission: User can edit own requests, Admin can edit any
        const isAdmin = checkAdminPermission(userEmail);
        const createdBy = rowData[createdByIndex];
        
        if (!isAdmin && createdBy !== userEmail) {
            return { success: false, message: "คุณสามารถแก้ไขคำขอของคุณเองได้เท่านั้น" };
        }

        // Check stock availability
        const stockValidation = validateStockAvailable(editRequestData.Items);
        if (!stockValidation.success) {
            return { success: false, message: `รายการไม่เพียงพอ: ${stockValidation.message}` };
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

        logRequestActivity("Edit Request", sessionId);

        return { success: true, message: "แก้ไขคำขอสำเร็จ" };
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน editRequest: " + error.toString());
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
            return { success: false, message: `รายการไม่เพียงพอสำหรับ Item ${item.Item_Id}: ${item.Item_Name}` };
        }

        // check each accessory
        for (let accessory of item.Accessories) {
            const accAvailableQty = accessoryMap[accessory.Accessory_Id] || 0;
            if(accAvailableQty < accessory.Qty) {
                return { success: false, message: `รายการไม่เพียงพอสำหรับ Accessory ${accessory.Accessory_Id}: ${accessory.Accessory_Name}` };
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

    Logger.log(`Managing stock with action: ${action} with items: ${JSON.stringify(items)}`);

    // check each item
    for (let item of items) {
        const itemData = itemsMap[item.Item_Id];
        if (!itemData) {
            Logger.log(`Warning: Item ID ${item.Item_Id} not found in stock`);
            continue;
        }

        // Use Return_Qty for returns, Qty for borrows
        const qtyToUpdate = action === 'increase' ? (item.Return_Qty !== undefined ? item.Return_Qty : item.Qty) : item.Qty;
        
        Logger.log(`Item ID ${item.Item_Id} - Qty to update: ${qtyToUpdate}`);

        // // Skip if quantity is 0 or undefined (nothing to update)
        // if (qtyToUpdate === undefined || qtyToUpdate === 0) {
        //     continue;
        // }
        // Note: Skip for support partial return with not return any of the item but return accessories only
        
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
                const accQtyToUpdate = action === 'increase' ? (accessory.Return_Qty !== undefined ? accessory.Return_Qty : accessory.Qty) : accessory.Qty;
                
                Logger.log(`Accessory ID ${accessory.Accessory_Id} - Qty to update: ${accQtyToUpdate}`);

                // // Skip if quantity is 0 or undefined (nothing to update)
                // if (accQtyToUpdate === undefined || accQtyToUpdate === 0) {
                //     continue;
                // }
                // Note: Skip for support partial return with not return any of the item but return accessories only
                
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

    Logger.log(`Items to update: ${JSON.stringify(updateItems)}`);
    // Update items stock
    updateItems.forEach(item => {
        Logger.log(`Updating Item ID ${item.Item_Id}: ${item.Available_Qty} ${action === 'decrease' ? '-' : '+'} ${item.Update_Qty} = ${item.New_Available_Qty}`);
        ss.getSheetByName(ITEM_SHEET_NAME).getRange(item.Item_Row_Index, itemHeaders.indexOf("Available_Qty") + 1).setValue(item.New_Available_Qty);
        ss.getSheetByName(ITEM_SHEET_NAME).getRange(item.Item_Row_Index, itemHeaders.indexOf("Modified_At") + 1).setValue(formatDate(new Date()));
    });

    Logger.log(`Accessories to update: ${JSON.stringify(updateAccessories)}`);

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
@ Now uses sessionId to support shared/group emails
*/
function distributeRequest(request, sessionId) {
    try {
        // Check admin permission
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
            throw new Error(`ไม่พบคำขอหมายเลข ${request.Request_Id}`);
        }

        if (rowData[headers.indexOf("Status")] !== "Submit") {
            throw new Error(`คำขอหมายเลข ${request.Request_Id} ไม่ได้อยู่ในสถานะ 'Submit'`);
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

        logRequestActivity("Distribute Request", sessionId);

        return { success: true, message: `คำขอหมายเลข ${request.Request_Id} ถูกแจกจ่ายเรียบร้อยแล้ว` };
    }
    catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน distributeRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Cancel Request
@ Now uses sessionId to support shared/group emails
*/
function cancelRequest(requestId, sessionId) {
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
            throw new Error(`ไม่พบคำขอหมายเลข ${requestId}`);
        }

        const statusIndex = headers.indexOf("Status");

        if (rowData[statusIndex] !== "Submit") {
            throw new Error(`คำขอหมายเลข ${requestId} ไม่ได้อยู่ในสถานะ 'Submit'`);
        }

        const modifiedByIndex = headers.indexOf("Modified_By");
        const modifiedAtIndex = headers.indexOf("Modified_At");
        requestSheet.getRange(rowIndex, statusIndex + 1).setValue("Cancelled");
        requestSheet.getRange(rowIndex, modifiedByIndex + 1).setValue(getCurrentUserEmail(sessionId));
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

        logRequestActivity("Cancel Request", sessionId);

        return { success: true, message: `คำขอหมายเลข ${requestId} ถูกยกเลิกเรียบร้อยแล้ว` };
    }
    catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน cancelRequest: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Return Request: Support partial return (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function returnRequest(returnRequestData, sessionId) {
    try {
        // Check admin permission
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
            throw new Error(`ไม่พบคำขอหมายเลข ${returnRequestData.Request_Id}`);
        }

        if (rowData[statusIndex] !== "Distributed" && rowData[statusIndex] !== "Partial_Returned") {
            throw new Error(`คำขอหมายเลข ${returnRequestData.Request_Id} ไม่ได้อยู่ในสถานะ 'Distributed' หรือ 'Partial_Returned' สถานะปัจจุบัน: ${rowData[statusIndex]}`);
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

        Logger.log(`All Items Fully Returned: ${allItemsFullyReturned}`);

        // Increase stock based on returned quantities
        const stockUpdateResult = manageStock(returnRequestData.Items, 'increase');
        Logger.log(`Stock update result: ${stockUpdateResult.updatedItems} items, ${stockUpdateResult.updatedAccessories} accessories`);

        // Update request status
        let newStatus = allItemsFullyReturned ? "Returned" : "Partial_Returned";
        requestSheet.getRange(rowIndex, statusIndex + 1).setValue(newStatus);
        requestSheet.getRange(rowIndex, modifiedByIndex + 1).setValue(getCurrentUserEmail(sessionId));
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
        logRequestActivity("Return Request", sessionId);

        return { 
            success: true, 
            message: `ดำเนินการคืนคำขอหมายเลข ${returnRequestData.Request_Id} ${allItemsFullyReturned ? 'เรียบร้อยแล้ว' : 'บางส่วนเรียบร้อยแล้ว'}`,
            status: newStatus,
            itemsUpdated: stockUpdateResult.updatedItems,
            accessoriesUpdated: stockUpdateResult.updatedAccessories
        };
    }
    catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน returnRequest: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน getRequestActivity: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

/*
@ Get System Activity - Shows login/logout activities (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function getSystemActivity(sessionId, page = 1, pageSize = 50) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
        Logger.log("เกิดข้อผิดพลาดใน getSystemActivity: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}

/*
@ Get Inventory Activity - Shows item/accessory changes (Admin only)
@ Now uses sessionId to support shared/group emails
*/
function getInventoryActivity(sessionId, page = 1, pageSize = 50) {
    try {
        const userEmail = getCurrentUserEmail(sessionId);
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
        Logger.log("เกิดข้อผิดพลาดใน getInventoryActivity: " + error.toString());
        return { success: false, message: error.toString(), data: [] };
    }
}


/**
@ Log Request Activity
@ Now supports sessionId for shared/group emails
*/
function logRequestActivity(actionType, sessionId = null) {
    try {
        const ss = getActiveSheet();
        const requestActivitySheet = ss.getSheetByName(REQUEST_ACTIVITY_SHEET_NAME);
        if (!requestActivitySheet) {
            throw new Error(`Sheet "${REQUEST_ACTIVITY_SHEET_NAME}" not found`);
        }

        const email = sessionId ? getCurrentUserEmail(sessionId) : Session.getActiveUser().getEmail();
        const timestamp = formatDate(new Date());
        const logId = getNextId(requestActivitySheet, 0);

        requestActivitySheet.appendRow([logId, email, actionType, timestamp]);

        Logger.log(`Request activity logged: ${email} - ${actionType} at ${timestamp}`);
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน logRequestActivity: " + error.toString());
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
        Logger.log("เกิดข้อผิดพลาดใน logSystemActivity: " + error.toString());
    }
}

/**
@ Log Inventory Activity
@ Now supports sessionId for shared/group emails
*/
function logInventoryActivity(actionType, sessionId = null) {
    try {
        const ss = getActiveSheet();
        const inventoryActivitySheet = ss.getSheetByName(INVENTORY_ACTIVITY_SHEET_NAME);
        if (!inventoryActivitySheet) {
            throw new Error(`Sheet "${INVENTORY_ACTIVITY_SHEET_NAME}" not found`);
        }

        const email = sessionId ? getCurrentUserEmail(sessionId) : Session.getActiveUser().getEmail();
        const timestamp = formatDate(new Date());
        const logId = getNextId(inventoryActivitySheet, 0);

        inventoryActivitySheet.appendRow([logId, email, actionType, timestamp]);

        Logger.log(`Inventory activity logged: ${email} - ${actionType} at ${timestamp}`);
    } catch (error) {
        Logger.log("เกิดข้อผิดพลาดใน logInventoryActivity: " + error.toString());
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
@ Now supports sessionId parameter (optional for manual execution)
*/
function initializeDataFromReference(sessionId = null) {
    try {
        const userEmail = sessionId ? getCurrentUserEmail(sessionId) : Session.getActiveUser().getEmail();
        
        // Check admin permission
        if (!checkAdminPermission(userEmail)) {
            return { success: false, message: "ไม่ได้รับอนุญาต: ต้องมีสิทธิ์ผู้ดูแลระบบ" };
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
        
        // Real Items Data from InventoryManagement - Items.csv
        // Format: [Item_Id, Item_Name, Item_Desc, Total_Qty, Available_Qty, Image]
        const realItems = [
            [1, "841 Titrando", "2.841.0010 (เสีย)", 1, 0, "https://image2url.com/r2/default/images/1770110030781-3b71ec6a-8388-44b3-9596-c990f34d08c8.png"],
            [2, "907 Titrando", "2.907.0010", 1, 1, "https://image2url.com/r2/default/images/1770111189839-f57eadea-d8f7-4b94-b143-44fabb67fa84.png"],
            [3, "808 Titrando", "2.808.0010", 1, 1, "https://image2url.com/r2/default/images/1770086412243-2f877e15-cec7-4120-bf1e-b3b7abec07cf.png"],
            [4, "852 Titrando", "2.852.0010", 1, 1, "https://image2url.com/r2/default/images/1770110600497-eb4fe05b-c636-45c1-8e49-8ebc28558c7c.png"],
            [5, "905 Titrando", "2.905.0010", 1, 1, "https://image2url.com/r2/default/images/1770111364183-40464f56-82bf-4197-97b8-06e18af778b5.png"],
            [6, "801 stirrer", "2.801.0010", 5, 4, "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTe5Bn4YioHHchJzvH3cDtShPfcJM_k1U9m-g&s"],
            [7, "916 Ti-touch", "2.916.0010", 2, 2, "https://image2url.com/r2/default/images/1770170012397-d921dca7-55a3-4d90-b9e5-5a1f158fd98d.png"],
            [8, "Eco Titrator", "2.1008.0010", 2, 2, "https://metrohm.scene7.com/is/image/metrohm/4840?$xh-544$&bfc=on"],
            [9, "904 Titrando", "2.904.0010", 1, 1, "https://image2url.com/r2/default/images/1770172036979-d384469b-d3b1-47ab-9ddd-b806e736b998.png"],
            [10, "900 Touch control", "2.900.0010", 3, 3, "https://static-data2.manualslib.com/product-images/1b6/2342664/metrohm-900-touch-control.jpg"],
            [11, "Eco Dosimat", "2.1007.0010", 1, 0, "https://image2url.com/r2/default/images/1770169561855-e1213163-8b8d-45b9-9677-f73dbcf151da.png"],
            [12, "859 Titrothrem", "2.859.0010", 1, 1, "https://image2url.com/r2/default/images/1770170225456-4d9376fd-0842-46c9-b322-ea5c1c2dde34.png"],
            [13, "800 Dosino", "2.800.0010", 4, 4, "https://metrohm.scene7.com/is/image/metrohm/2127_s?$xh-1280$&bfc=on"],
            [14, "803 Ti Stand", "2.803.0010", 4, 4, "https://image2url.com/r2/default/images/1770085759392-64db9849-5016-4a77-aca5-99c7e0837be4.jpg"],
            [15, "805 Dosimat", "2.805.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1378_s?$xh-1280$&bfc=on"],
            [16, "802 Rod stirrer", "2.802.0010", 6, 6, "https://images.offerup.com/8Zo1b7oifuxDUjKDMQmM8GNXAJI=/1440x1920/ebd7/ebd7f5271c354489a4acff5019a9cb09.jpg"],
            [17, "728 stirrer without stand", "2.728.0010", 5, 5, "https://image2url.com/r2/default/images/1770174227483-8d86ddd5-d534-4235-9efd-1275d23507dd.png"],
            [18, "Omnis Dosing Module without stirrer", "2.1003.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/4249?$xh-1280$&bfc=on"],
            [19, "Omnis Dosing Module with stirrer", "2.1003.0110", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/4250?$xh-1280$&bfc=on"],
            [20, "OMNIS sample robot", "", 1, 1, "https://image2url.com/r2/default/images/1770170628931-928392be-245f-4902-9680-de6288be283a.png"],
            [21, "OMNIS Titrator with stirrer", "2.1001.0020", 2, 2, "https://metrohm.scene7.com/is/image/metrohm/4248?$xh-544$&bfc=on"],
            [22, "OMNIS coulometer with stirrer", "2.1018.0020", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/210180020_FV?$xh-1280$&bfc=on"],
            [23, "Eco KF Titrator", "2.1027.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1102800010?$xh-544$&bfc=on"],
            [24, "860 KF Thermoprep", "2.860.0010", 1, 1, "https://image2url.com/r2/default/images/1770111576358-30724e7c-556c-4633-ae0a-559a9c919c1c.png"],
            [25, "832 KF Thermoprep", "2.832.0010", 1, 1, "https://image2url.com/r2/default/images/1770111749313-ba0cfa3a-5dd1-4e3d-9749-82f21bc68bd3.png"],
            [26, "899 Coulometer", "2.899.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/3553_s?$xh-1280$&bfc=on"],
            [27, "756 KF Coulometer", "2.756.0010", 1, 1, "https://www.ntc-tech.com/cdn/shop/products/ukRoN5p-m0Hh60lV7KpDtMODpG0qG3lLXohGGtdK0_BXveWcrQ6OJ21Hu23eLWkYIy7FlFHCNh0s5jcP0NDfn_asIvdDo4k_s4470.jpg"],
            [28, "831 KF Coulometer", "2.831.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1498_s?$xh-1280$&bfc=on"],
            [29, "808 Touch control", "2.808.0010", 1, 1, "https://image2url.com/r2/default/images/1770112064025-95265755-5393-4adf-9cd6-bf611c96837f.png"],
            [30, "885 compact oven sc", "2.885.0010", 1, 1, "https://image2url.com/r2/default/images/1770111926027-af5b978d-fc87-4772-b245-5241839c1949.png"],
            [31, "915 KF Ti-Touch", "2.915.0010", 1, 1, "https://image2url.com/r2/default/images/1770112334125-53b1d506-d9ac-4d0f-b7e4-c5e13f15bd79.png"],
            [32, "862 compact Titrosampler", "2.862.0010", 1, 1, "https://image2url.com/r2/default/images/1770112210246-b11ec0de-4491-4454-bbd0-dddc319d1089.png"],
            [33, "870 KF Titrino plus", "2.870.0010", 1, 1, "https://image2url.com/r2/default/images/1770112548836-996ed91b-afd7-4817-bbf3-251d99a4cd5f.png"],
            [34, "848 Titrino plus", "2.848.0010", 1, 1, "https://image2url.com/r2/default/images/1770112869227-0c00e97c-21e7-40e0-8622-52c67403ca37.png"],
            [35, "Eco coulometer", "2.1028.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/1102800010?$xh-544$&bfc=on"],
            [36, "solvent pump", "2.1029.0010", 1, 1, "https://image2url.com/r2/default/images/1770169307529-2ce27fd1-14b2-4ecb-a416-57e0751a1bb7.png"],
            [37, "914 pH/conductomer", "2.914.0020", 1, 1, "https://image2url.com/r2/default/images/1770113012483-86a1e09d-9c7c-4f36-96bc-7baca9036ebc.png"],
            [38, "912 Conductometer", "2.912.0010", 1, 1, "https://image2url.com/r2/default/images/1770113078469-c9f491b4-5594-4c9b-940e-c60e4d079b6a.png"],
            [39, "913 pH meter", "2.913.0010", 1, 1, "https://image2url.com/r2/default/images/1770113144119-c332b957-6231-44d8-80ae-1c3135995589.png"],
            [40, "867 pH module", "2.867.0010", 1, 1, "https://image2url.com/r2/default/images/1770109167866-45310e93-f7d2-4d57-8685-d18b2e904eaf.jpg"],
            [41, "826 pH mobile", "2.826.0010", 1, 1, "https://image2url.com/r2/default/images/1770109355002-ac532c7b-6458-48e7-8482-7173c71a13f6.jpg"],
            [42, "949 pH meter", "2.949.0010", 1, 1, "https://image2url.com/r2/default/images/1770113371736-0201045c-5bf5-4d6e-8012-c8629166b787.png"],
            [43, "827 pH lab", "2.827.0010", 1, 1, "https://image2url.com/r2/default/images/1770113243494-dfc740ae-32d1-4ed2-b2ae-412c963112be.png"],
            [44, "781 pH/Ion Meter", "2.781.0010", 1, 1, "https://image2url.com/r2/default/images/1770173212831-2a45cab4-a75b-4db7-aa9f-808db2804ea6.png"],
            [45, "856 Conductivity module", "2.856.0010", 1, 1, "https://image2url.com/r2/default/images/1770173367571-183417fa-9ccc-410d-9407-183f29ec9a6d.png"],
            [46, "704 pH meter", "2.704.0010", 1, 1, "https://image2url.com/r2/default/images/1770173654721-4936e03b-6816-40db-8b3b-0fd36e32e58f.png"],
            [47, "890 Titrando", "2.890.0010", 1, 1, "https://image2url.com/r2/default/images/1770172406088-5e939ffb-2e2b-4218-908e-ca983823dbe1.png"],
            [48, "892 Professional Rancimat", "2.892.0010", 1, 1, "https://metrohm.scene7.com/is/image/metrohm/3571_s?$xh-1280$&bfc=on"],
            [49, "743 Rancimat", "2.743.0010", 1, 1, "https://image2url.com/r2/default/images/1770171916935-bac81e4e-8b2a-4ea8-8c17-96d6f7de664d.png"],
            [50, "Pt Titrode", "6.0431.100", 3, 3, "https://s7e5a.scene7.com/is/image/metrohm/60431100?$xh-1280$&bfc=on"],
            [51, "Ag Titrode", "6.0430.100", 7, 6, "https://image2url.com/r2/default/images/1770260948756-7f5b7e92-88c3-47bd-af5b-4c730cbe1182.png"],
            [52, "Ag/s Titrode", "6.00430.100 s", 1, 1, "https://image2url.com/r2/default/images/1770261033594-39c6edd8-041b-4c62-9472-b626678a145c.png"],
            [53, "Ag brominate Titrode", "6.0430.100", 2, 2, ""],
            [54, "LL Profitrode length 17.8cm", "6.0255.110", 1, 1, "https://image2url.com/r2/default/images/1770261381594-7970edad-10bb-4aa1-9311-ce402ba62579.png"],
            [55, "Separate Pt wire electrode", "6.0301.100", 4, 4, "https://image2url.com/r2/default/images/1770261529597-fc96a714-1bb6-491d-978c-7496caf7900c.png"],
            [56, "Double Pt sheet electrode", "6.0334.000", 4, 4, "https://s7e5a.scene7.com/is/image/metrohm/60340000?$xh-1280$&bfc=on"],
            [57, "Ionic surfactant", "6.0507.120", 3, 3, "https://image2url.com/r2/default/images/1770262012582-6016a9be-56e1-450d-913c-5b802c5b0ecc.png"],
            [58, "Surfactrode Resistant", "6.0507.130", 4, 4, "https://image2url.com/r2/default/images/1770262150716-0fb9d583-7c3c-47b4-b278-a52ab908c320.png"],
            [59, "Surfactrode Refill", "6.0507.140", 1, 1, "https://image2url.com/r2/default/images/1770262320347-411a3c42-0b01-4bd0-9ecd-5ae4701b874d.png"],
            [60, "Cation surfacte", "6.0507.150", 2, 2, "https://image2url.com/r2/default/images/1770262464543-6ce13f97-ac7b-40d3-9d80-75f0e96d65fb.png"],
            [61, "Double Pt sheet electrode", "6.0309.100", 1, 1, "https://image2url.com/r2/default/images/1770261912733-790b79ab-47f0-435a-bef8-7a245fe0141a.png"],
            [62, "Combined Pt wire electrode", "6.0401.100", 2, 2, ""],
            [63, "Thermoprobe", "6.9011.020", 1, 1, "https://image2url.com/r2/default/images/1770262645839-edc33ff2-d8a4-4f84-8b1d-b7299b6f8477.png"],
            [64, "Thermoprobe HF", "6.9011.040", 1, 1, "https://image2url.com/r2/default/images/1770262711044-4b50c776-66bc-4aa8-a174-1dd505ba4abc.png"],
            [65, "iSolvotrode", "6.0279.300", 2, 2, "https://image2url.com/r2/default/images/1770262782223-7bea597b-ce2a-4a73-be98-ff6a301e6325.png"],
            [66, "iEcotrode plus", "6.0280.300", 2, 2, "https://image2url.com/r2/default/images/1770263368301-2614814d-8c11-4121-9287-2aadefcbabd4.png"],
            [67, "Double Au Ringelectrode", "6.00353.100", 1, 1, "https://image2url.com/r2/default/images/1770263489729-91392e8e-6117-4e50-ab3a-5ba92ac94902.png"],
            [68, "dAg-Titrode", "6.00400.300", 1, 1, ""],
            [69, "iAg Titrode", "6.0470.300", 3, 3, "https://image2url.com/r2/default/images/1770263677370-c25cb769-9af6-4b27-b9fd-66b6a12c7ead.png"],
            [70, "iUnitrode pt 1000", "6.0278.300", 1, 1, "https://image2url.com/r2/default/images/1770264048168-a95b979b-1cd9-4d7e-8497-d748177ff666.png"],
            [71, "Comb. Pt Ring WOC", "6.0451.100", 2, 2, "https://image2url.com/r2/default/images/1770264284994-11d64ccb-343b-48ed-b20d-9b64d2c7f2e4.png"],
            [74, "conductometric cell pt1000 c=0.11", "6.0918.040", 1, 1, "https://image2url.com/r2/default/images/1770264460183-7f041fe8-0907-459a-a6cb-d9a1cef75704.png"],
            [75, "Ag/AgCl ref", "6.0729.100", 1, 1, "https://image2url.com/r2/default/images/1770264525643-13a2026f-1be6-43fd-aa0b-b96370ad2729.png"],
            [76, "Electrode shaft", "6.124105", 3, 3, "https://image2url.com/r2/default/images/1770264284994-11d64ccb-343b-48ed-b20d-9b64d2c7f2e4.png"],
            [77, "Fluoride-ISE", "6.0502.150", 1, 1, "https://image2url.com/r2/default/images/1770266669254-fbe9ff5b-41c7-486f-8907-0ced4113d37a.png"],
            [78, "Cu ISE WOC SGJ", "6.0502.140", 1, 1, "https://image2url.com/r2/default/images/1770266788097-b9bc7c94-325b-498a-b5ed-100fdae18b85.png"],
            [79, "Shaft F.Plug In Electr. SCH", "6.1241.040", 2, 2, "https://image2url.com/r2/default/images/1770267064294-60bee5f9-6393-46de-8bd3-fbf2956c9ac6.png"],
            [81, "Na glass electrode", "6.0501.100", 1, 1, "https://image2url.com/r2/default/images/1770267247529-db5c245b-2c6b-4995-8eb8-c7323bb69b4b.png"],
            [82, "Separate glass", "6.0150.100", 1, 1, "https://image2url.com/r2/default/images/1770267436667-9eb96d12-c5a2-4ca6-9985-a16241d9dbb7.png"],
            [83, "Lead-ISE", "6.0502.170", 1, 1, "https://image2url.com/r2/default/images/1770270939292-04cabc91-2868-4937-8aea-5d279f1cd9d9.png"],
            [84, "Cl-ISE", "6.0502.120", 1, 1, "https://image2url.com/r2/default/images/1770271058467-d2ff7ca4-83a4-4832-a3cf-bebb7198e2e3.png"],
            [85, "Insertion pt1000 WOC", "6.00226.600", 1, 1, "https://image2url.com/r2/default/images/1770271142207-7ce896d4-5aeb-4c53-84e2-49eb2b87a532.png"],
            [86, "Polymer membrane Na WOC", "6.0508.100", 1, 1, "https://image2url.com/r2/default/images/1770264284994-11d64ccb-343b-48ed-b20d-9b64d2c7f2e4.png"],
            [87, "4 conduct pt1000 c=0.5", "6.0917.080", 2, 2, "https://image2url.com/r2/default/images/1770264284994-11d64ccb-343b-48ed-b20d-9b64d2c7f2e4.png"],
            [89, "Iodine ISE", "6.0502.160", 1, 1, "https://image2url.com/r2/default/images/1770276886629-ad2d0c21-a509-4fbe-9914-84fc728b022d.png"],
            [90, "Sodium ISE", "6.0508.100", 1, 1, "https://image2url.com/r2/default/images/1770277013507-7332a125-1ca6-4bac-841d-b935e01199df.png"],
            [91, "NH3 selective electrode", "6.0506.010", 1, 1, "https://image2url.com/r2/default/images/1770277120462-b1afa445-ae37-41d5-a50f-17c48a27f267.png"],
            [92, "conductometric cell pt1000 c=1.59", "6.0919.140", 1, 1, "https://image2url.com/r2/default/images/1770277223843-fab0081d-debe-4127-8927-19b3441cee42.png"],
            [93, "NH3 module kit", "6.1255.000", 2, 2, "https://image2url.com/r2/default/images/1770277317481-4fcc39ff-2c52-46e2-bd54-9a94b1a4d1c9.png"],
            [95, "Cyanide-ISE", "6.0502.130", 1, 1, "https://image2url.com/r2/default/images/1770277474061-4b0e9581-4dfc-4b28-8a7e-037a03eeb10d.png"],
            [96, "NO3 ISE liquid membrane", "6.0504.120", 1, 1, ""],
            [97, "Cd2+ ISE", "6.0502.110", 1, 1, "https://image2url.com/r2/default/images/1770260793268-d7d104d3-614e-4b86-9eee-d34e7b4a48b6.png"],
            [98, "Br- ISE", "6.0502.110", 1, 1, ""],
            [99, "electrode shaft woc SGJ", "6.1241.050", 2, 2, "https://image2url.com/r2/default/images/1770277958661-fcd015f9-d1c7-435b-a3c1-2b063dd1e9b4.png."],
            [103, "LL Primatrode NTC", "6.0228.020", 3, 3, "https://image2url.com/r2/default/images/1770278573849-20fb9870-810a-4b91-bbae-cd13373c2389.png"],
            [105, "LL ISE ref", "6.0750.100", 1, 1, "https://image2url.com/r2/default/images/1770278713189-ac353731-f718-41e0-8e43-b11d67ff3f72.png"],
            [106, "Temp. sensor", "6.1110.100", 2, 2, "https://image2url.com/r2/default/images/1770278888101-16213bf0-4233-4d11-82ef-270916a8dcf1.png"],
            [108, "Ecotrode Gel NTC", "6.0221.600", 1, 1, "https://image2url.com/r2/default/images/1770278888101-16213bf0-4233-4d11-82ef-270916a8dcf1.png"],
            [109, "Comb. Flat membrane", "6.0256.100", 2, 2, "https://image2url.com/r2/default/images/1770279247160-7672d6a1-b63b-41d6-83c7-92a12d1d324f.png"],
            [111, "Comb. pH glass", "6.0219.110", 2, 2, ""],
            [113, "Comb. Glass", "6.0210.100", 2, 2, ""],
            [115, "LL Micro glass", "6.0234.100", 3, 3, "https://image2url.com/r2/default/images/1770346101100-5fb64d7b-20d4-4b8e-a6dc-638fd9702293.png"],
            [118, "separate pH glass", "6.0123.100", 3, 3, "https://image2url.com/r2/default/images/1770264284994-11d64ccb-343b-48ed-b20d-9b64d2c7f2e4.png"],
            [120, "LL aquatrode plus Pt1000", "6.0257.600", 1, 1, "https://image2url.com/r2/default/images/1770363611101-feeb11bc-2794-4469-868d-9b0805fa1876.png"],
            [121, "LL Porotrode WOC", "6.0235.200", 2, 2, "https://image2url.com/r2/default/images/1770363715884-25a3856c-9519-408f-822a-f84aaef16d58.png"],
            [124, "LL Primatrode NTC", "6.022802", 1, 1, "https://image2url.com/r2/default/images/1770364100582-015bc2e9-607e-4313-a9d0-5103ec16c9af.png"],
            [125, "Primatrode", "6.0228.010", 2, 2, "https://image2url.com/r2/default/images/1770364230160-5891c0c5-bcc0-4f51-92b1-f91f64dd3d5d.png"],
            [127, "EtOH trode", "6.0269.100", 1, 1, "https://image2url.com/r2/default/images/1770364484269-e4a0233b-6dfc-44a5-9316-ab1f9b5d8924.png"],
            [128, "Generator with diaphragm", "6.0344.100", 1, 1, "https://image2url.com/r2/default/images/1770364558239-0bd36f65-9f80-47fa-ba8c-3359a2f984f7.png"],
            [129, "Generator without diaphragm", "6.0342.110", 5, 5, "https://image2url.com/r2/default/images/1770364670854-3a9b2200-7e42-47bc-a0fb-ae5e9515402a.png"],
            [134, "Double pt wire for coulometry", "6.0341.100", 6, 6, "https://image2url.com/r2/default/images/1770364812752-5001dba4-b6cf-4104-8ab6-935e7ea1746c.png"],
            [140, "Double pt wire for Volumetry", "6.0338.100", 2, 2, "https://image2url.com/r2/default/images/1770365063714-456d6141-7a09-4ee7-b2af-c8ff89f55836.png"],
            [142, "LL Ecotrode plus WOC", "6.0262.100", 1, 1, "https://image2url.com/r2/default/images/1770365153954-2b6fac70-2186-4eea-8011-a2a93cf51dc5.png"],
            [143, "LL Unitrode WOC", "6.0259.100", 2, 2, "https://image2url.com/r2/default/images/1770365406743-6ebc77ec-cf09-4ee6-83c3-43d6e203438d.png"],
            [147, "NIO surfactant electrode", "6.0507.010", 1, 1, "https://image2url.com/r2/default/images/1770365601032-7e9f1693-41a6-475e-b3a8-1265615a222d.png"],
            [148, "Ecotrode Gel", "6.0221.100", 1, 1, "https://image2url.com/r2/default/images/1770365810866-ef371b2f-14f3-4f30-8ac1-fe3ab747097b.png"],
            [149, "Ion Sel El Pb WOC SGJ", "6.0502.170", 1, 1, "https://image2url.com/r2/default/images/1770365939310-1702bfd0-7571-495c-a734-15d1f74b95dc.png"],
            [150, "DJ Ag/AgCl Reference", "6.0726.100", 1, 1, "https://image2url.com/r2/default/images/1770366030893-a8dbda37-0c08-429c-b77a-909e9fac2462.png"],
            [151, "NH3 selective gas", "6.0506.100", 1, 1, "https://image2url.com/r2/default/images/1770366129736-45e1838d-a5fd-44ec-82db-20076ec83719.png"],
            [153, "Double Pt-wire electrod for volumetric", "6.0338.100", 2, 2, "https://image2url.com/r2/default/images/1770264284994-11d64ccb-343b-48ed-b20d-9b64d2c7f2e4.png"],
            [173, "Ag brominate Titrode", "6.0430.100", 2, 2, "https://s7e5a.scene7.com/is/image/metrohm/60430100_s?$xh-1280$&bfc=on"],
            [174, "804 Ti stand", "2.804.0010", 2, 2, "https://metrohm.scene7.com/is/image/metrohm/1556_s?$xh-1280$&bfc=on"]
        ];
        
        // Import Items with real Item_Id from CSV
        for (let i = 0; i < realItems.length; i++) {
            const item = realItems[i];
            
            itemsSheet.appendRow([
                item[0],      // Item_Id (real ID from CSV)
                item[1],      // Item_Name
                item[2],      // Item_Desc
                item[3],      // Total_Qty
                item[4],      // Available_Qty
                item[5],      // Image
                true,         // Active
                "system",     // Created_By
                timestamp,    // Created_At
                "system",     // Modified_By
                timestamp     // Modified_At
            ]);
            
            itemsImported++;
        }
        
        // Real Accessories Data from InventoryManagement - Accessories.csv
        // Format: [Accessory_Id, Accessory_Name, Accessory_Desc, Total_Qty, Available_Qty]
        const realAccessories = [
            [1, "stirrering proppeller 104mm", "", 1, 1],
            [3, "stirrering propeller 96mm", "", 2, 2],
            [4, "Dosing unit 2 mL for Titrate", "6.1580.120", 2, 2],
            [5, "Dosing unit 10 mL for Titrate", "6.1580.210", 5, 5],
            [6, "Dosing unit 20 mL for Titrate", "6.1580.220", 2, 2],
            [7, "Dosing unit 50 mL for Titrate", "6.1580.250", 2, 2],
            [8, "OMNIS Dosing 10 mL", "6.01503.210", 2, 2],
            [9, "OMNIS Dosing 20 mL", "6.01503.220", 2, 2],
            [14, "Holding clip for bottles", "6.2043.005", 24, 24],
            [15, "Power cable", "สายPower อย่างเดียวต่อกับเครื่อง", 38, 38],
            [16, "Controller cable", "สาย controller เชื่อม software กับตัวเครื่อง", 4, 4],
            [17, "SET PC OMNIS", "CPU, Keyboard, mouse, power cable, Screen", 1, 1],
            [18, "Rod stand + lock camp", "แท่งเหล็ก + ตัวล็อค", 39, 39],
            [20, "Electrode holder", "6.2021.020 Electrode holder for 4 electrodes and 2 buret tips", 6, 6],
            [21, "SET Tubing for Dosing unit", "1 Tubing เขียว, Tip, Microvalve\n1 Tubing ใส+จุกใส, Molecular sieve", 10, 10],
            [22, "Brown glass bottle", "ขวดเปล่า", 15, 15],
            [23, "Electrod cable /1 m /F", "6.2104.020", 11, 11],
            [24, "2 Lan cable + Hub box + 1 Hub Power cacle", "สายแลนเหลือง 2 + กล่อง Hub + สายชาร์จ Hub", 1, 1],
            [25, "OMNIS Dosing unit 10 mL", "6.01508.210", 3, 3],
            [26, "OMNIS Dosing unit 20 mL", "6.01508.220", 2, 2],
            [27, "OMNIS Holder", "", 3, 3],
            [28, "OMNIS Molecular sieve with Cap", "", 4, 4],
            [29, "Red Electrode cable plug-in head G (temp.) / plug P, 0.55 m", "6.02104.020\nสายแดงใช้กับ temp senser", 1, 1],
            [30, "Blue Electrode cable plug-in head G (pol.) / plug P, 0.55 m", "6.02104.040\nสายสีน้ำเงินใช้กับ Polarized electrode to plug F (เป็นขั้วโลหะและเป็นแง่ง หรือ Pt sheet)", 2, 2],
            [31, "Green Electrode cable plug-in head G / plug P, 0.55 m", "6.02104.000\nสายสีเขียวใช้กับ Electrode ทั่วไป to plug F (ทุกประเภท General)", 1, 1],
            [32, "Magnetic bar", "", 14, 14],
            [33, "Power cable with adapter", "สาย Power ที่มี Adapter สำหรับบางเครื่อง เช่น Eco", 6, 5],
            [34, "SET Tubing for Cylinder Unit", "3 Tubing เขียว,+Tip+Microvalve\n1 Tubing ใส+ Cap SET", 10, 10],
            [35, "Exchange unit 5 mL", "6.1576.150", 1, 1],
            [36, "Exchange unit 10 mL", "6.1576.210", 2, 2],
            [37, "Exchange unit 20 mL", "6.1576.220", 5, 5],
            [38, "Exchange unit 50 mL", "6.1576.250", 1, 1],
            [39, "Cylinder unit 5 mL", "6.1518.150", 7, 5],
            [40, "Cylinder unit 10 mL", "6.1518.210", 13, 7],
            [41, "Cylinder unit 1 mL", "6.1518.110", 1, 0],
            [42, "Cylinder unit 20 mL", "6.1518.220", 20, 6],
            [43, "Cylinder unit 50 mL", "6.1518.250", 1, 1],
            [44, "Electrode cable /1 m /H", "6.2104.120", 15, 15],
            [45, "Electrode cable 1 m, 2xB", "6.2104.080", 3, 3],
            [46, "Strand / 1 m / 2 x B (banana)", "6.2106.020", 7, 7],
            [47, "854 iConnect", "2.854.0010", 2, 2],
            [48, "Electrode cable 2 m, 2 x 2 mm (temp senser)", "6.2104.150", 3, 3],
            [49, "Connect cable for 703 stirrer 0.5 m", "6.2108.100", 1, 1],
            [50, "Electrode cable / 1 m / RM", "6.2104.130\nWith RM plug. For connecting electrodes with plug-in head G - Radiometer instruments", 1, 1],
            [51, "Adaptor USB Mini (OTG) - USB A", "6.2151.100", 1, 1],
            [52, "Cable MDL PL/SO 0.5 m", "6.02102.010", 1, 1],
            [53, "Cable MDL PL/SO 1 m", "6.02102.020\n\n", 1, 1]
        ];
        
        // Import Accessories with real Accessory_Id from CSV
        for (let i = 0; i < realAccessories.length; i++) {
            const accessory = realAccessories[i];
            
            accessorySheet.appendRow([
                accessory[0], // Accessory_Id (real ID from CSV)
                accessory[1], // Accessory_Name
                accessory[2], // Accessory_Desc
                accessory[3], // Total_Qty
                accessory[4], // Available_Qty
                true,         // Active
                "system",     // Created_By
                timestamp,    // Created_At
                "system",     // Modified_By
                timestamp     // Modified_At
            ]);
            
            accessoriesImported++;
        }
        
        // Real Item-Accessory Mappings from InventoryManagement - Item_Accessory_Mapping.csv
        // Format: [Mapping_Id, Item_Id, Accessory_Id]
        // Note: Using real IDs from CSV - ALL 321 mappings included
        const realMappings = [
            [51, 6, 32], [52, 16, 3], [53, 16, 1],
            [66, 13, 5], [67, 13, 4], [68, 13, 6], [69, 13, 7],
            [70, 18, 8], [71, 19, 8], [72, 21, 8], [73, 22, 8],
            [74, 18, 9], [75, 19, 9], [76, 21, 9], [77, 22, 9],
            [78, 1, 14], [79, 2, 14], [80, 3, 14], [81, 4, 14], [82, 5, 14], [83, 7, 14], [84, 9, 14], [85, 31, 14],
            [86, 1, 15], [87, 2, 15], [88, 3, 15], [89, 4, 15], [90, 5, 15], [91, 9, 15], [92, 12, 15],
            [93, 18, 15], [94, 19, 15], [95, 20, 15], [96, 21, 15], [97, 22, 15],
            [98, 24, 15], [99, 25, 15], [100, 26, 15], [101, 27, 15], [102, 28, 15], [103, 30, 15],
            [104, 32, 15], [105, 33, 15], [106, 34, 15], [107, 40, 15], [108, 45, 15], [109, 47, 15],
            [110, 48, 15], [111, 49, 15],
            [112, 7, 33], [113, 8, 33], [114, 10, 33], [115, 11, 33], [116, 15, 33], [117, 23, 33],
            [118, 31, 33], [119, 35, 33], [120, 36, 33],
            [121, 1, 16], [122, 2, 16], [123, 3, 16], [124, 4, 16], [125, 5, 16], [126, 9, 16], [127, 47, 16],
            [128, 1, 17], [129, 2, 17], [130, 3, 17], [131, 4, 17], [132, 5, 17], [133, 9, 17],
            [134, 18, 17], [135, 19, 17], [136, 20, 17], [137, 21, 17], [138, 22, 17], [139, 47, 17],
            [140, 6, 18], [141, 7, 18], [142, 8, 18], [143, 11, 18], [144, 14, 18],
            [145, 18, 18], [146, 19, 18], [147, 21, 18], [148, 22, 18], [149, 23, 18], [150, 31, 18],
            [151, 35, 18], [152, 37, 18], [153, 38, 18], [154, 39, 18], [155, 40, 18], [156, 42, 18], [157, 45, 18],
            [158, 1, 20], [159, 2, 20], [160, 3, 20], [161, 4, 20], [162, 5, 20], [163, 7, 20], [164, 8, 20],
            [165, 9, 20], [166, 11, 20], [167, 12, 20], [168, 15, 20], [169, 23, 20], [170, 31, 20],
            [171, 35, 20], [172, 37, 20], [173, 38, 20], [174, 39, 20], [175, 40, 20], [176, 41, 20],
            [177, 42, 20], [178, 43, 20], [179, 44, 20], [180, 45, 20], [181, 46, 20], [182, 47, 20],
            [198, 1, 22], [199, 2, 22], [200, 3, 22], [201, 4, 22], [202, 5, 22], [203, 7, 22],
            [204, 8, 22], [205, 9, 22], [206, 11, 22], [207, 12, 22], [208, 15, 22],
            [209, 18, 22], [210, 19, 22], [211, 21, 22], [212, 22, 22], [213, 23, 22],
            [214, 25, 22], [215, 26, 22], [216, 27, 22], [217, 28, 22], [218, 31, 22],
            [219, 33, 22], [220, 34, 22], [221, 35, 22], [222, 36, 22], [223, 47, 22],
            [225, 3, 34], [226, 8, 34], [227, 9, 34], [228, 11, 34], [229, 15, 34],
            [230, 23, 34], [231, 31, 34], [232, 47, 34],
            [233, 13, 21],
            [234, 18, 24], [235, 19, 24], [236, 20, 24], [237, 21, 24], [238, 22, 24],
            [239, 18, 25], [240, 19, 25], [241, 21, 25], [242, 22, 25],
            [243, 18, 26], [244, 19, 26], [245, 21, 26], [246, 22, 26],
            [247, 18, 27], [248, 19, 27], [249, 21, 27], [250, 22, 27],
            [609, 3, 35], [610, 9, 35], [611, 15, 35], [612, 33, 35], [613, 34, 35], [614, 47, 35],
            [621, 3, 37], [622, 9, 37], [623, 15, 37], [624, 33, 37], [625, 34, 37], [626, 47, 37],
            [627, 3, 36], [628, 9, 36], [629, 15, 36], [630, 33, 36], [631, 34, 36], [632, 47, 36],
            [633, 3, 38], [634, 9, 38], [635, 15, 38], [636, 33, 38], [637, 34, 38], [638, 47, 38],
            [639, 3, 39], [640, 9, 39], [641, 15, 39], [642, 33, 39], [643, 34, 39], [644, 47, 39],
            [645, 3, 40], [646, 9, 40], [647, 15, 40], [648, 33, 40], [649, 34, 40], [650, 47, 40],
            [651, 3, 41], [652, 9, 41], [653, 15, 41], [654, 33, 41], [655, 34, 41], [656, 47, 41],
            [657, 3, 42], [658, 9, 42], [659, 15, 42], [660, 33, 42], [661, 34, 42], [662, 47, 42],
            [663, 3, 43], [664, 9, 43], [665, 15, 43], [666, 33, 43], [667, 34, 43], [668, 47, 43],
            [669, 50, 23], [670, 51, 23], [671, 52, 23], [672, 53, 23], [673, 54, 23], [674, 55, 23],
            [675, 56, 23], [676, 57, 23], [677, 58, 23], [678, 59, 23], [679, 60, 23], [680, 61, 23],
            [681, 62, 23], [682, 63, 23], [683, 64, 23], [684, 65, 23], [685, 66, 23], [686, 67, 23],
            [687, 68, 23], [688, 69, 23], [689, 70, 23], [690, 71, 23], [691, 74, 23], [692, 75, 23],
            [693, 76, 23], [694, 77, 23], [695, 78, 23], [696, 79, 23], [697, 81, 23], [698, 82, 23],
            [699, 83, 23], [700, 84, 23], [701, 85, 23], [702, 86, 23], [703, 87, 23], [704, 89, 23],
            [705, 90, 23], [706, 91, 23], [707, 92, 23], [708, 93, 23], [709, 95, 23], [710, 96, 23],
            [711, 97, 23], [712, 98, 23], [713, 99, 23], [714, 103, 23], [715, 105, 23], [716, 106, 23],
            [717, 108, 23], [718, 109, 23], [719, 111, 23], [720, 113, 23], [721, 115, 23], [722, 118, 23],
            [723, 120, 23], [724, 121, 23], [725, 124, 23], [726, 125, 23], [727, 127, 23], [728, 128, 23],
            [729, 129, 23], [730, 134, 23], [731, 140, 23], [732, 142, 23], [733, 143, 23], [734, 147, 23],
            [735, 148, 23], [736, 149, 23], [737, 150, 23], [738, 151, 23], [739, 152, 23], [740, 155, 23],
            [741, 156, 23], [742, 157, 23], [743, 158, 23], [744, 159, 23], [745, 160, 23], [746, 163, 23],
            [747, 164, 23], [748, 165, 23], [749, 166, 23], [750, 167, 23], [751, 168, 23], [752, 169, 23],
            [753, 170, 23], [754, 173, 23],
            [755, 18, 52], [756, 19, 52], [757, 20, 52], [758, 21, 52], [759, 22, 52],
            [760, 18, 28], [761, 19, 28], [762, 21, 28], [763, 22, 28]
        ];
        
        // Import Mappings with real Mapping_Id, Item_Id, and Accessory_Id from CSV
        for (let i = 0; i < realMappings.length; i++) {
            const mapping = realMappings[i];
            
            itemAccessorySheet.appendRow([
                mapping[0],   // Mapping_Id (real ID from CSV)
                mapping[1],   // Item_Id (real ID from CSV)
                mapping[2],   // Accessory_Id (real ID from CSV)
                "system",     // Created_By
                timestamp,    // Created_At
                true          // Active
            ]);
            
            mappingsImported++;
        }
        
        // Initial Users Data from InventoryManagement - Users.csv
        // Format: [Email, Password (hashed), Permission, Active]
        // Note: All 4 users from CSV included
        const initialUsers = [
            ["modmastei2@gmail.com", "b7766cf93f0fcbcfa13adcc202419a4e5f21816f70360b6e76fd18342b56fdd8", "Admin", true],
            ["admin@admin.com", "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9", "Admin", true],
            ["warangkana.jenny13@gmail.com", "8d969eef6ecad3c29a3a629280e686cf0c3f5d5a86aff3ca12020c923adc6c92", "User", true],
            ["sales@sales.com", "8d969eef6ecad3c29a3a629280e686cf0c3f5d5a86aff3ca12020c923adc6c92", "User", true]
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
        Logger.log("เกิดข้อผิดพลาดใน initializeDataFromReference: " + error.toString());
        return { success: false, message: error.toString() };
    }
}