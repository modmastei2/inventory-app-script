// ============================================
// Authentication Functions
// ============================================

/*
@ Hash Password using SHA-256
*/
function hashPassword(password) {
    const rawHash = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        password,
        Utilities.Charset.UTF_8
    );
    
    // Convert byte array to hex string
    let hash = '';
    for (let i = 0; i < rawHash.length; i++) {
        const byte = rawHash[i];
        if (byte < 0) {
            hash += ('0' + (byte + 256).toString(16)).slice(-2);
        } else {
            hash += ('0' + byte.toString(16)).slice(-2);
        }
    }
    return hash;
}

/*
@ Verify Password
*/
function verifyPassword(password, hash) {
    return hashPassword(password) === hash;
}

/*
@ Create Password Hash (Helper function for testing)
*/
function createPasswordHash(password) {
    const hash = hashPassword(password);
    Logger.log(`Password: ${password}`);
    Logger.log(`Hash: ${hash}`);
    return hash;
}

/*
@ Generate Session Token
*/
function generateSessionToken() {
    const timestamp = new Date().getTime();
    const random = Math.random().toString(36).substring(2, 15);
    return `${timestamp}_${random}`;
}

/*
@ Login Function
*/
function login(email, password) {
    try {
        const ss = getActiveSheet();
        const userSheet = ss.getSheetByName(USER_SHEET_NAME);
        const sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);

        if (!userSheet) {
            throw new Error(`Sheet "${USER_SHEET_NAME}" not found`);
        }

        if (!sessionSheet) {
            throw new Error(`Sheet "${SESSION_SHEET_NAME}" not found`);
        }

        const userData = userSheet.getDataRange().getValues();
        const userHeaders = userData[0];
        const emailCol = userHeaders.indexOf("Email");
        const passwordCol = userHeaders.indexOf("Password");
        const permissionCol = userHeaders.indexOf("Permission");
        const activeCol = userHeaders.indexOf("Active");

        // Normalize email for case-insensitive comparison
        const normalizedEmail = email.trim().toLowerCase();

        // Find user by email (case-insensitive)
        let userFound = null;
        for (let i = 1; i < userData.length; i++) {
            if (userData[i][emailCol].toLowerCase() === normalizedEmail) {
                userFound = {
                    Email: userData[i][emailCol],
                    PasswordHash: userData[i][passwordCol],
                    Permission: userData[i][permissionCol] || 'Guest',
                    Active: userData[i][activeCol]
                };
                break;
            }
        }

        if (!userFound) {
            return { success: false, message: "Email not found" };
        }

        // Check if user is active
        if (userFound.Active === false || userFound.Active === 'FALSE') {
            return { success: false, message: "Account is inactive. Please contact administrator." };
        }

        // Verify password
        if (!verifyPassword(password, userFound.PasswordHash)) {
            return { success: false, message: "Invalid password" };
        }
        
        // Generate session token
        const sessionToken = generateSessionToken();
        const now = new Date();
        const sessionId = getNextId(sessionSheet, 0);

        // Save session to Sessions sheet
        sessionSheet.appendRow([
            sessionId,
            userFound.Email,
            userFound.Permission,
            now,
            now
        ]);

        logSystemActivity(`User ${email} logged in`, email);

        return {
            success: true,
            sessionToken: sessionToken,
            sessionId: sessionId,
            email: userFound.Email,
            permission: userFound.Permission
        };
    } catch (error) {
        Logger.log("Error in login: " + error.toString());
        return { success: false, message: error.toString() };
    }
}

/*
@ Validate Session
*/
function validateSession(sessionId, sessionToken) {
    try {
        const ss = getActiveSheet();
        const sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);

        if (!sessionSheet) {
            return { valid: false, message: "Session sheet not found" };
        }

        const sessionData = sessionSheet.getDataRange().getValues();
        const sessionHeaders = sessionData[0];

        // Find session
        for (let i = 1; i < sessionData.length; i++) {
            if (sessionData[i][0] == sessionId) {
                const lastActivity = new Date(sessionData[i][4]);
                const now = new Date();
                const hoursSinceActivity = (now - lastActivity) / (1000 * 60 * 60);

                // Session expires after 24 hours
                if (hoursSinceActivity > 24) {
                    return { valid: false, message: "Session expired" };
                }

                const email = sessionData[i][1];

                // Get latest permission from Users sheet
                const userSheet = ss.getSheetByName(USER_SHEET_NAME);
                if (!userSheet) {
                    return { valid: false, message: "User sheet not found" };
                }

                const userData = userSheet.getDataRange().getValues();
                const userHeaders = userData[0];
                const emailCol = userHeaders.indexOf("Email");
                const permissionCol = userHeaders.indexOf("Permission");

                let currentPermission = 'Guest';
                const normalizedEmail = email.toLowerCase();
                for (let j = 1; j < userData.length; j++) {
                    if (userData[j][emailCol].toLowerCase() === normalizedEmail) {
                        currentPermission = userData[j][permissionCol] || 'Guest';
                        break;
                    }
                }

                // Update last activity and permission in session
                sessionSheet.getRange(i + 1, 3).setValue(currentPermission);
                sessionSheet.getRange(i + 1, 5).setValue(now);

                return {
                    valid: true,
                    email: email,
                    permission: currentPermission
                };
            }
        }

        return { valid: false, message: "Invalid session" };
    } catch (error) {
        Logger.log("Error in validateSession: " + error.toString());
        return { valid: false, message: error.toString() };
    }
}

/*
@ Logout Function
*/
function logout(sessionId) {
    try {
        const ss = getActiveSheet();
        const sessionSheet = ss.getSheetByName(SESSION_SHEET_NAME);

        if (!sessionSheet) {
            return { success: false, message: "Session sheet not found" };
        }

        const sessionData = sessionSheet.getDataRange().getValues();

        // Find and delete session
        for (let i = 1; i < sessionData.length; i++) {
            if (sessionData[i][0] == sessionId) {
                const userEmail = sessionData[i][1];
                sessionSheet.deleteRow(i + 1);
                logSystemActivity(`User ${userEmail} logged out`, userEmail);
                return { success: true, message: "Logged out successfully" };
            }
        }

        return { success: false, message: "Session not found" };
    } catch (error) {
        Logger.log("Error in logout: " + error.toString());
        return { success: false, message: error.toString() };
    }
}
