// ============================================
// Trigger Management Functions
// ============================================

/*
@ Simple test function to verify script is working
@ Run this first to authorize the script
*/
function testBasicFunction() {
    try {
        Logger.log("Test function started");
        Logger.log("Current time: " + new Date());
        
        // Test accessing spreadsheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        Logger.log("Spreadsheet name: " + ss.getName());
        
        return { success: true, message: "Basic test passed", timestamp: new Date() };
    } catch (error) {
        Logger.log("Test failed: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Test if getActiveSheet and constants are accessible
*/
function testDependencies() {
    try {
        Logger.log("Testing dependencies...");
        
        // Test getActiveSheet from code.gs
        const ss = getActiveSheet();
        Logger.log("✓ getActiveSheet() works");
        
        // Test constants
        Logger.log("REQUEST_SHEET_NAME: " + REQUEST_SHEET_NAME);
        Logger.log("✓ Constants accessible");
        
        // Test if required functions exist
        const functionsExist = {
            archiveOldRequests: typeof archiveOldRequests === 'function',
            sendOverdueItemsEmail: typeof sendOverdueItemsEmail === 'function',
            getActiveSheet: typeof getActiveSheet === 'function'
        };
        
        Logger.log("Functions check: " + JSON.stringify(functionsExist, null, 2));
        
        const allExist = functionsExist.archiveOldRequests && 
                        functionsExist.sendOverdueItemsEmail && 
                        functionsExist.getActiveSheet;
        
        if (allExist) {
            return { success: true, message: "All dependencies OK", details: functionsExist };
        } else {
            return { success: false, message: "Some functions missing", details: functionsExist };
        }
        
    } catch (error) {
        Logger.log("Dependency test failed: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Setup daily trigger for archiving old requests
@ Archives Cancelled/Returned requests older than 7 days
@ Runs daily at 2:00 AM
*/
function setupArchiveTrigger() {
    try {
        Logger.log("Setting up archive trigger...");
        
        // Delete existing triggers for archiveOldRequests to avoid duplicates
        const triggers = ScriptApp.getProjectTriggers();
        let deletedCount = 0;
        
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'archiveOldRequests') {
                try {
                    ScriptApp.deleteTrigger(trigger);
                    deletedCount++;
                } catch (e) {
                    Logger.log("Failed to delete existing trigger: " + e.toString());
                }
            }
        });
        
        if (deletedCount > 0) {
            Logger.log(`Deleted ${deletedCount} existing archive trigger(s)`);
        }
        
        // Create new daily trigger at 2 AM
        const newTrigger = ScriptApp.newTrigger('archiveOldRequests')
            .timeBased()
            .atHour(2)
            .everyDays(1)
            .create();
        
        Logger.log("Archive trigger created successfully (2:00 AM daily) - Trigger ID: " + newTrigger.getUniqueId());
        return { success: true, message: "Archive trigger created at 2:00 AM", triggerId: newTrigger.getUniqueId() };
        
    } catch (error) {
        const errorMsg = "Error setting up archive trigger: " + error.toString() + " | Stack: " + (error.stack || "N/A");
        Logger.log(errorMsg);
        return { success: false, message: error.toString(), error: error };
    }
}

/*
@ Remove daily archive trigger
*/
function removeArchiveTrigger() {
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
        Logger.log("Error removing archive trigger: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Setup daily trigger for overdue items email notification
@ Sends email for items past their return date
@ Runs daily at 8:30 AM
*/
function setupOverdueEmailTrigger() {
    try {
        Logger.log("Setting up overdue email trigger...");
        
        // Delete existing triggers for this function
        const triggers = ScriptApp.getProjectTriggers();
        let deletedCount = 0;
        
        for (let trigger of triggers) {
            if (trigger.getHandlerFunction() === 'sendOverdueItemsEmail') {
                try {
                    ScriptApp.deleteTrigger(trigger);
                    deletedCount++;
                } catch (e) {
                    Logger.log("Failed to delete existing trigger: " + e.toString());
                }
            }
        }
        
        if (deletedCount > 0) {
            Logger.log(`Deleted ${deletedCount} existing overdue email trigger(s)`);
        }
        
        // Create new trigger at 08:30 daily
        const newTrigger = ScriptApp.newTrigger('sendOverdueItemsEmail')
            .timeBased()
            .atHour(8)
            .nearMinute(30)
            .everyDays(1)
            .create();
        
        Logger.log("Overdue email trigger created successfully (8:30 AM daily) - Trigger ID: " + newTrigger.getUniqueId());
        return { success: true, message: "Overdue email trigger created at 8:30 AM", triggerId: newTrigger.getUniqueId() };
        
    } catch (error) {
        const errorMsg = "Error setting up overdue email trigger: " + error.toString() + " | Stack: " + (error.stack || "N/A");
        Logger.log(errorMsg);
        return { success: false, message: error.toString(), error: error };
    }
}

/*
@ Remove overdue email trigger
*/
function removeOverdueEmailTrigger() {
    try {
        const triggers = ScriptApp.getProjectTriggers();
        let removed = 0;
        
        for (let trigger of triggers) {
            if (trigger.getHandlerFunction() === 'sendOverdueItemsEmail') {
                ScriptApp.deleteTrigger(trigger);
                removed++;
            }
        }
        
        Logger.log(`Removed ${removed} overdue email trigger(s)`);
        return { success: true, message: `Removed ${removed} trigger(s)` };
        
    } catch (error) {
        Logger.log("Error removing overdue email trigger: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Setup all triggers at once
@ Call this function to initialize all system triggers
*/
function setupAllTriggers() {
    try {
        Logger.log("========================================");
        Logger.log("Setting up all triggers...");
        Logger.log("========================================");
        
        const results = {};
        
        // Setup archive trigger
        Logger.log("\n[1/2] Setting up Archive trigger...");
        try {
            results.archive = setupArchiveTrigger();
            Logger.log("Archive trigger result: " + JSON.stringify(results.archive));
        } catch (e) {
            Logger.log("Archive trigger failed: " + e.toString());
            results.archive = { success: false, message: e.toString() };
        }
        
        // Add a small delay between triggers
        Utilities.sleep(1000);
        
        // Setup overdue email trigger
        Logger.log("\n[2/2] Setting up Overdue Email trigger...");
        try {
            results.overdueEmail = setupOverdueEmailTrigger();
            Logger.log("Overdue email trigger result: " + JSON.stringify(results.overdueEmail));
        } catch (e) {
            Logger.log("Overdue email trigger failed: " + e.toString());
            results.overdueEmail = { success: false, message: e.toString() };
        }
        
        const allSuccess = results.archive.success && results.overdueEmail.success;
        
        Logger.log("\n========================================");
        if (allSuccess) {
            Logger.log("✓ All triggers set up successfully");
            Logger.log("========================================");
            return { 
                success: true, 
                message: "All triggers created:\n- Archive (2:00 AM)\n- Overdue Email (8:30 AM)",
                details: results
            };
        } else {
            Logger.log("✗ Some triggers failed to set up");
            Logger.log("========================================");
            
            let failureMsg = "Failed triggers:\n";
            if (!results.archive.success) failureMsg += "- Archive: " + results.archive.message + "\n";
            if (!results.overdueEmail.success) failureMsg += "- Overdue Email: " + results.overdueEmail.message + "\n";
            
            return { 
                success: false, 
                message: failureMsg,
                details: results
            };
        }
        
    } catch (error) {
        const errorMsg = "Error setting up all triggers: " + error.toString() + " | Stack: " + (error.stack || "N/A");
        Logger.log(errorMsg);
        return { success: false, message: error.toString(), error: error };
    }
}

/*
@ Remove all triggers
@ Use this to clean up all system triggers
*/
function removeAllTriggers() {
    try {
        Logger.log("Removing all triggers...");
        
        const results = {
            archive: removeArchiveTrigger(),
            overdueEmail: removeOverdueEmailTrigger()
        };
        
        Logger.log("All triggers removed");
        return { 
            success: true, 
            message: "All triggers removed successfully",
            details: results
        };
        
    } catch (error) {
        Logger.log("Error removing all triggers: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ List all active triggers
@ Useful for debugging and monitoring
*/
function listAllTriggers() {
    try {
        const triggers = ScriptApp.getProjectTriggers();
        const triggerList = [];
        
        for (let trigger of triggers) {
            const info = {
                handlerFunction: trigger.getHandlerFunction(),
                triggerSource: trigger.getTriggerSource().toString(),
                eventType: trigger.getEventType().toString(),
                uniqueId: trigger.getUniqueId()
            };
            
            // Add time-based trigger details if applicable
            if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
                try {
                    info.hour = "Daily schedule";
                } catch (e) {
                    info.hour = "N/A";
                }
            }
            
            triggerList.push(info);
        }
        
        Logger.log("Active triggers: " + JSON.stringify(triggerList, null, 2));
        return { 
            success: true, 
            triggers: triggerList,
            count: triggerList.length
        };
        
    } catch (error) {
        Logger.log("Error listing triggers: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Test if required functions exist
@ Run this before setting up triggers to verify all dependencies
*/
function testTriggerDependencies() {
    try {
        Logger.log("Testing trigger dependencies...");
        
        const results = {
            archiveOldRequests: false,
            sendOverdueItemsEmail: false,
            getActiveSheet: false
        };
        
        // Test if archiveOldRequests exists
        try {
            if (typeof archiveOldRequests === 'function') {
                results.archiveOldRequests = true;
                Logger.log("✓ archiveOldRequests function found");
            }
        } catch (e) {
            Logger.log("✗ archiveOldRequests function NOT found: " + e.toString());
        }
        
        // Test if sendOverdueItemsEmail exists
        try {
            if (typeof sendOverdueItemsEmail === 'function') {
                results.sendOverdueItemsEmail = true;
                Logger.log("✓ sendOverdueItemsEmail function found");
            }
        } catch (e) {
            Logger.log("✗ sendOverdueItemsEmail function NOT found: " + e.toString());
        }
        
        // Test if getActiveSheet exists
        try {
            if (typeof getActiveSheet === 'function') {
                results.getActiveSheet = true;
                Logger.log("✓ getActiveSheet function found");
            }
        } catch (e) {
            Logger.log("✗ getActiveSheet function NOT found: " + e.toString());
        }
        
        const allFound = results.archiveOldRequests && 
                        results.sendOverdueItemsEmail && 
                        results.getActiveSheet;
        
        if (allFound) {
            Logger.log("\n✓ All required functions are available");
            return { success: true, message: "All dependencies found", details: results };
        } else {
            Logger.log("\n✗ Some required functions are missing");
            return { success: false, message: "Missing dependencies", details: results };
        }
        
    } catch (error) {
        Logger.log("Error testing dependencies: " + error.toString());
        return { success: false, message: error.toString() };
    }
}
