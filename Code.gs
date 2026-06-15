// ============================================================================
// GOOGLE APPS SCRIPT - UPDATED WITH EMAIL & PASSWORD VALIDATION
// Deploy as Web App with "Execute as" = Your email, "Access" = Anyone
// ============================================================================

const SPREADSHEET_ID = '148za9Kw8c6Ao0opmbcUGuHcLcSIE0048yZCvutCCOgU'; // Replace with your Sheets ID

function doPost(e) {
    try {
        const action = e.parameter.action;

        if (action === 'login') {
            return handleLogin(e);
        } else if (action === 'register') {
            return handleRegister(e);
        } else if (action === 'submitScore') {
            return handleSubmitScore(e);
        } else if (action === 'getLeaderboard') {
            return handleGetLeaderboard(e);
        }

        return respond({ success: false, message: 'Invalid action' });
    } catch (error) {
        console.error('Error:', error);
        return respond({ success: false, message: error.toString() });
    }
}

function respond(data) {
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// SHEET INITIALIZATION
// ============================================================================

function initializeSheets() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Create Users sheet if not exists
    if (!ss.getSheetByName('Users')) {
        const usersSheet = ss.insertSheet('Users');
        usersSheet.appendRow(['Username', 'Email', 'Password', 'TotalScore', 'GamesPlayed', 'CreatedDate']);
    }

    // Create Scores sheet if not exists
    if (!ss.getSheetByName('Scores')) {
        const scoresSheet = ss.insertSheet('Scores');
        scoresSheet.appendRow(['Username', 'Level', 'Score', 'Time', 'Moves', 'Date']);
    }
}

// ============================================================================
// AUTHENTICATION
// ============================================================================

function hashPassword(password) {
    // Simple hash - in production, use more secure method
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
        .map(function(char) {
            var v = (char < 0) ? 256 + char : char;
            return ("0" + v.toString(16)).slice(-2);
        }).join('');
}

function validateEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validateUsername(username) {
    return username && username.length >= 3 && /^[a-zA-Z0-9_]+$/.test(username);
}

function validatePassword(password) {
    return password && password.length >= 6;
}

function handleRegister(e) {
    initializeSheets();
    
    const username = e.parameter.username ? e.parameter.username.trim() : '';
    const email = e.parameter.email ? e.parameter.email.trim() : '';
    const password = e.parameter.password || '';

    // Validate inputs
    if (!validateUsername(username)) {
        return respond({ success: false, message: 'Username must be 3+ chars, alphanumeric only' });
    }

    if (!validateEmail(email)) {
        return respond({ success: false, message: 'Please enter valid email' });
    }

    if (!validatePassword(password)) {
        return respond({ success: false, message: 'Password must be 6+ chars' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName('Users');
    const values = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 1).getValues();

    // Check if username exists
    for (let i = 0; i < values.length; i++) {
        if (values[i][0] === username) {
            return respond({ success: false, message: 'Username already exists' });
        }
    }

    // Check if email exists
    const emailValues = usersSheet.getRange(2, 2, usersSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < emailValues.length; i++) {
        if (emailValues[i][0] === email) {
            return respond({ success: false, message: 'Email already registered' });
        }
    }

    // Add new user
    const hashedPassword = hashPassword(password);
    usersSheet.appendRow([username, email, hashedPassword, 0, 0, new Date()]);

    return respond({ 
        success: true, 
        message: 'User registered successfully',
        username: username,
        email: email,
        totalScore: 0,
        gamesPlayed: 0,
        stats: {
            easy: { bestTime: null, bestMoves: null, highScore: 0 },
            medium: { bestTime: null, bestMoves: null, highScore: 0 },
            hard: { bestTime: null, bestMoves: null, highScore: 0 }
        }
    });
}

function handleLogin(e) {
    initializeSheets();
    
    const username = e.parameter.username ? e.parameter.username.trim() : '';
    const password = e.parameter.password || '';

    if (!username || !password) {
        return respond({ success: false, message: 'Username and password required' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName('Users');
    const data = usersSheet.getDataRange().getValues();

    const hashedPassword = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === username && data[i][2] === hashedPassword) {
            // Found user - load their stats
            const stats = loadUserStats(username);
            
            return respond({
                success: true,
                message: 'Login successful',
                username: username,
                email: data[i][1],
                totalScore: data[i][3] || 0,
                gamesPlayed: data[i][4] || 0,
                stats: stats
            });
        }
    }

    return respond({ success: false, message: 'Invalid username or password' });
}

function loadUserStats(username) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const scoresSheet = ss.getSheetByName('Scores');
    const data = scoresSheet.getDataRange().getValues();

    const stats = {
        easy: { bestTime: null, bestMoves: null, highScore: 0 },
        medium: { bestTime: null, bestMoves: null, highScore: 0 },
        hard: { bestTime: null, bestMoves: null, highScore: 0 }
    };

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === username) {
            const level = data[i][1];
            const score = data[i][2];
            const time = data[i][3];
            const moves = data[i][4];

            if (!stats[level]) {
                stats[level] = { bestTime: null, bestMoves: null, highScore: 0 };
            }

            // Update best time
            if (!stats[level].bestTime || time < stats[level].bestTime) {
                stats[level].bestTime = time;
            }

            // Update best moves
            if (!stats[level].bestMoves || moves < stats[level].bestMoves) {
                stats[level].bestMoves = moves;
            }

            // Update high score
            if (score > stats[level].highScore) {
                stats[level].highScore = score;
            }
        }
    }

    return stats;
}

// ============================================================================
// SCORE SUBMISSION
// ============================================================================

function handleSubmitScore(e) {
    initializeSheets();
    
    const username = e.parameter.username || '';
    const level = e.parameter.level || '';
    const score = parseInt(e.parameter.score) || 0;
    const time = parseInt(e.parameter.time) || 0;
    const moves = parseInt(e.parameter.moves) || 0;

    if (!username || !level || score < 0) {
        return respond({ success: false, message: 'Invalid score data' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const scoresSheet = ss.getSheetByName('Scores');

    // Add score
    scoresSheet.appendRow([username, level, score, time, moves, new Date()]);

    // Update user's total score
    const usersSheet = ss.getSheetByName('Users');
    const data = usersSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === username) {
            const currentTotal = (data[i][3] || 0) + score;
            const gamesPlayed = (data[i][4] || 0) + 1;
            usersSheet.getRange(i + 1, 4).setValue(currentTotal);
            usersSheet.getRange(i + 1, 5).setValue(gamesPlayed);
            break;
        }
    }

    return respond({ success: true, message: 'Score submitted successfully' });
}

// ============================================================================
// LEADERBOARD
// ============================================================================

function handleGetLeaderboard(e) {
    initializeSheets();
    
    const level = e.parameter.level || 'easy';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const scoresSheet = ss.getSheetByName('Scores');
    const data = scoresSheet.getDataRange().getValues();

    // Collect scores for level
    const scores = {};
    for (let i = 1; i < data.length; i++) {
        if (data[i][1] === level) {
            const username = data[i][0];
            const score = data[i][2];
            const time = data[i][3];

            if (!scores[username] || score > scores[username].score) {
                scores[username] = { score: score, time: time };
            }
        }
    }

    // Convert to array and sort by score descending
    const leaderboard = Object.keys(scores).map(username => ({
        username: username,
        score: scores[username].score,
        time: scores[username].time
    })).sort((a, b) => b.score - a.score);

    return respond({
        success: true,
        data: leaderboard.slice(0, 50) // Top 50
    });
}

// ============================================================================
// SETUP FUNCTION - RUN ONCE
// ============================================================================

function setup() {
    initializeSheets();
    Logger.log('Sheets initialized');
}
