// ============================================
// Archive Functions for Historical Data
// ============================================

/*
@ Archive old requests (Cancelled or Returned) older than 7 days
@ Moves data from Requests, Request_Item, Request_Item_Accessory
@ to Historical_Requests, Historical_Request_Item, Historical_Request_Item_Accessory
@ This function should be run daily via time-based trigger
*/
function archiveOldRequests() {
    try {
        const ss = getActiveSheet();
        const requestSheet = ss.getSheetByName(REQUEST_SHEET_NAME);
        const requestItemSheet = ss.getSheetByName(REQUEST_ITEM_SHEET_NAME);
        const requestItemAccessorySheet = ss.getSheetByName(REQUEST_ITEM_ACCESSORY_SHEET_NAME);
        
        const historicalRequestsSheet = ss.getSheetByName(HISTORICAL_REQUESTS_SHEET_NAME);
        const historicalRequestItemSheet = ss.getSheetByName(HISTORICAL_REQUEST_ITEM_SHEET_NAME);
        const historicalRequestItemAccessorySheet = ss.getSheetByName(HISTORICAL_REQUEST_ITEM_ACCESSORY_SHEET_NAME);
        
        if (!requestSheet || !historicalRequestsSheet) {
            Logger.log("Required sheets not found for archiving");
            return;
        }
        
        const requestData = requestSheet.getDataRange().getValues();
        const headers = requestData[0];
        
        // Find column indices
        const requestIdCol = headers.indexOf("Request_Id");
        const statusCol = headers.indexOf("Status");
        const modifiedAtCol = headers.indexOf("Modified_At");
        
        // Calculate cutoff date (7 days ago)
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - 7);
        cutoffDate.setHours(0, 0, 0, 0); // Set to start of day
        
        let archivedCount = 0;
        const rowsToDelete = [];
        
        // Process from bottom to top to avoid index shifting issues
        for (let i = requestData.length - 1; i >= 1; i--) {
            const row = requestData[i];
            const requestId = row[requestIdCol];
            const status = row[statusCol];
            const modifiedAt = new Date(row[modifiedAtCol]);
            
            // Check if request should be archived
            // Status must be "Cancelled" or "Returned" AND modified date > 7 days ago
            if ((status === "Cancelled" || status === "Returned") && modifiedAt < cutoffDate) {
                // Archive the request
                historicalRequestsSheet.appendRow(row);
                
                // Archive related Request_Item records
                if (requestItemSheet) {
                    const itemData = requestItemSheet.getDataRange().getValues();
                    const itemHeaders = itemData[0];
                    const itemRequestIdCol = itemHeaders.indexOf("Request_Id");
                    
                    for (let j = itemData.length - 1; j >= 1; j--) {
                        if (itemData[j][itemRequestIdCol] == requestId) {
                            if (historicalRequestItemSheet) {
                                historicalRequestItemSheet.appendRow(itemData[j]);
                            }
                            requestItemSheet.deleteRow(j + 1);
                        }
                    }
                }
                
                // Archive related Request_Item_Accessory records
                if (requestItemAccessorySheet) {
                    const accessoryData = requestItemAccessorySheet.getDataRange().getValues();
                    const accessoryHeaders = accessoryData[0];
                    const accessoryRequestIdCol = accessoryHeaders.indexOf("Request_Id");
                    
                    for (let k = accessoryData.length - 1; k >= 1; k--) {
                        if (accessoryData[k][accessoryRequestIdCol] == requestId) {
                            if (historicalRequestItemAccessorySheet) {
                                historicalRequestItemAccessorySheet.appendRow(accessoryData[k]);
                            }
                            requestItemAccessorySheet.deleteRow(k + 1);
                        }
                    }
                }
                
                // Mark row for deletion
                rowsToDelete.push(i + 1);
                archivedCount++;
            }
        }
        
        // Delete archived request rows
        rowsToDelete.forEach(rowNum => {
            requestSheet.deleteRow(rowNum);
        });
        
        Logger.log(`Archived ${archivedCount} old requests to Historical sheets`);
        return { success: true, archivedCount: archivedCount };
        
    } catch (error) {
        Logger.log("Error in archiveOldRequests: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Set up daily trigger for archiveOldRequests
@ Run this function once to install the daily trigger
@ The trigger will run archiveOldRequests every day at 2 AM
*/
function setupDailyArchiveTrigger() {
    try {
        // Delete existing triggers for archiveOldRequests to avoid duplicates
        const triggers = ScriptApp.getProjectTriggers();
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'archiveOldRequests') {
                ScriptApp.deleteTrigger(trigger);
            }
        });
        
        // Create new daily trigger at 2 AM
        ScriptApp.newTrigger('archiveOldRequests')
            .timeBased()
            .atHour(2)
            .everyDays(1)
            .create();
        
        Logger.log("Daily archive trigger set up successfully");
        return { success: true, message: "Daily archive trigger created at 2:00 AM" };
        
    } catch (error) {
        Logger.log("Error setting up trigger: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Remove daily archive trigger
@ Run this function to remove the daily trigger
*/
function removeDailyArchiveTrigger() {
    try {
        const triggers = ScriptApp.getProjectTriggers();
        let removed = 0;
        
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'archiveOldRequests') {
                ScriptApp.deleteTrigger(trigger);
                removed++;
            }
        });
        
        Logger.log(`Removed ${removed} archive trigger(s)`);
        return { success: true, message: `Removed ${removed} trigger(s)` };
        
    } catch (error) {
        Logger.log("Error removing trigger: " + error.toString());
        return { success: false, message: error.toString() };
    }
}
