// ==========================================
// SMART HRMS - Enhanced Version with All Features
// Code.gs - Server Side Script
// ==========================================

// Configuration
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
const COMPANY_NAME = 'Smart HRMS';
const GRACE_MINUTES = 15;

// Shift Times Configuration
const SHIFT_TIMES = {
  'A': { start: '06:00', end: '14:00', name: 'Shift A (06:00 - 14:00)' },
  'B': { start: '14:00', end: '22:00', name: 'Shift B (14:00 - 22:00)' },
  'C': { start: '22:00', end: '06:00', name: 'Shift C (22:00 - 06:00)' }
};

// Office Location for Geofencing (Configure as needed)
const OFFICE_LOCATION = {
  latitude: 28.6139,  // Default: New Delhi
  longitude: 77.2090,
  radiusMeters: 500   // 500 meters radius
};

// Initialize Spreadsheet
function getSpreadsheet() {
  if (!SPREADSHEET_ID) {
    throw new Error('Spreadsheet ID not configured. Please set SPREADSHEET_ID in script properties.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Get or Create Sheet
function getOrCreateSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheetHeaders(sheet, name);
  }
  return sheet;
}

// Initialize Sheet Headers
function initializeSheetHeaders(sheet, name) {
  const headers = {
    'Employees': ['EmployeeID', 'Name', 'Email', 'Phone', 'Department', 'Designation', 'JoinDate', 'Password', 'Role', 'Status', 'Shift', 'CreatedAt', 'WeekOffDay', 'BankName', 'AccountNumber', 'IFSCCode'],
    'Attendance': ['ID', 'EmployeeID', 'Date', 'PunchIn', 'PunchOut', 'Status', 'Shift', 'WorkHours', 'PunchInLocation', 'PunchOutLocation', 'PunchInSelfie', 'PunchOutSelfie', 'Notes'],
    'Leaves': ['LeaveID', 'EmployeeID', 'LeaveType', 'StartDate', 'EndDate', 'Days', 'Reason', 'Status', 'AppliedDate', 'ApprovedBy', 'ApprovedDate', 'Comments'],
    'LeaveBalance': ['EmployeeID', 'Casual', 'Sick', 'Earned', 'CompOff', 'Total', 'Year'],
    'PunchCorrections': ['RequestID', 'EmployeeID', 'Date', 'CurrentPunchIn', 'CurrentPunchOut', 'RequestedPunchIn', 'RequestedPunchOut', 'Reason', 'Status', 'RequestDate', 'ApprovedBy', 'ApprovedDate', 'Comments'],
    'Holidays': ['HolidayID', 'Name', 'Date', 'Type', 'Optional', 'CreatedAt'],
    'Roster': ['RosterID', 'EmployeeID', 'Name', 'Date', 'DayOfWeek', 'Shift', 'IsWeekOff', 'IsBuffer', 'CreatedAt'],
    'RosterConfig': ['ConfigID', 'EmployeeID', 'WeekOffDay', 'RotationSequence', 'CurrentShift', 'LastUpdated'],
    'Warnings': ['WarningID', 'EmployeeID', 'IssueDate', 'Reason', 'Description', 'IssuedBy', 'Status', 'CreatedAt'],
    'OTP': ['Email', 'OTP', 'CreatedAt', 'ExpiresAt', 'Used'],
    'AuditLog': ['LogID', 'Timestamp', 'UserID', 'Action', 'Details', 'IPAddress'],
    'Settings': ['SettingKey', 'SettingValue', 'UpdatedAt']
  };
  
  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    sheet.setFrozenRows(1);
  }
}

// ==========================================
// AUTHENTICATION FUNCTIONS
// ==========================================

function login(email, password) {
  try {
    email = String(email || '').toLowerCase().trim();
    password = String(password || '').trim();
    
    console.log('Login attempt for email: ' + email);
    
    const sheet = getOrCreateSheet('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][2] || '').toLowerCase().trim();
      const rowPassword = String(data[i][7] || '').trim();
      const rowStatus = String(data[i][9] || '').trim().toLowerCase();
      
      if (rowEmail === email && rowPassword === password) {
        if (rowStatus !== 'active') {
          return { success: false, message: 'Account is inactive. Please contact admin.' };
        }
        
        const sessionId = Utilities.getUuid();
        const user = {
          id: String(data[i][0] || ''),
          name: String(data[i][1] || ''),
          email: String(data[i][2] || ''),
          phone: String(data[i][3] || ''),
          department: String(data[i][4] || ''),
          designation: String(data[i][5] || ''),
          joinDate: String(data[i][6] || ''),
          role: String(data[i][8] || 'employee'),
          shift: String(data[i][10] || 'A'),
          weekOffDay: String(data[i][12] || 'Sunday'),
          sessionId: sessionId
        };
        
        // Store session
        PropertiesService.getUserProperties().setProperty('session_' + sessionId, JSON.stringify(user));
        
        // Audit log
        addAuditLog(user.id, 'LOGIN', 'User logged in successfully');
        
        return { success: true, user: user };
      }
    }
    
    return { success: false, message: 'Invalid email or password.' };
  } catch (error) {
    console.error('Login error: ' + error.toString());
    return { success: false, message: 'Login failed: ' + error.message };
  }
}

function logout(sessionId) {
  try {
    sessionId = String(sessionId || '').trim();
    PropertiesService.getUserProperties().deleteProperty('session_' + sessionId);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function validateSession(sessionId) {
  try {
    sessionId = String(sessionId || '').trim();
    const sessionData = PropertiesService.getUserProperties().getProperty('session_' + sessionId);
    if (sessionData) {
      return { valid: true, user: JSON.parse(sessionData) };
    }
    return { valid: false };
  } catch (error) {
    return { valid: false };
  }
}

// ==========================================
// PASSWORD RESET FUNCTIONS
// ==========================================

function sendOTP(email) {
  try {
    email = String(email || '').toLowerCase().trim();
    console.log('Sending OTP to: ' + email);
    
    const sheet = getOrCreateSheet('Employees');
    const data = sheet.getDataRange().getValues();
    
    let employeeExists = false;
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][2] || '').toLowerCase().trim();
      if (rowEmail === email) {
        employeeExists = true;
        break;
      }
    }
    
    if (!employeeExists) {
      return { success: false, message: 'Email not registered.' };
    }
    
    const otp = String(Math.floor(100000 + Math.random() * 900000));
    const now = new Date();
    const expiresAt = new Date(now.getTime() + (10 * 60 * 1000));
    
    const otpSheet = getOrCreateSheet('OTP');
    const otpData = otpSheet.getDataRange().getValues();
    
    for (let i = otpData.length - 1; i >= 1; i--) {
      const rowEmail = String(otpData[i][0] || '').toLowerCase().trim();
      if (rowEmail === email) {
        otpSheet.deleteRow(i + 1);
      }
    }
    
    SpreadsheetApp.flush();
    otpSheet.appendRow([email, otp, now.toISOString(), expiresAt.toISOString(), false]);
    SpreadsheetApp.flush();
    
    const subject = 'Password Reset OTP - ' + COMPANY_NAME;
    const body = `
      <html>
        <body style="font-family: Arial, sans-serif;">
          <h2>Password Reset Request</h2>
          <p>Your OTP for password reset is: <strong style="font-size: 24px; color: #4285f4;">${otp}</strong></p>
          <p>This OTP will expire in 10 minutes.</p>
          <p>If you did not request this, please ignore this email.</p>
          <br>
          <p>Regards,<br>${COMPANY_NAME} Team</p>
        </body>
      </html>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
    
    console.log('OTP sent successfully: ' + otp);
    return { success: true, message: 'OTP sent to your email.' };
  } catch (error) {
    console.error('Send OTP error: ' + error.toString());
    return { success: false, message: 'Failed to send OTP: ' + error.message };
  }
}

function verifyOTP(email, otp) {
  try {
    email = String(email || '').toLowerCase().trim();
    otp = String(otp || '').trim();
    
    console.log('Verifying OTP - Email: ' + email + ', OTP: ' + otp);
    
    const otpSheet = getOrCreateSheet('OTP');
    SpreadsheetApp.flush();
    const otpData = otpSheet.getDataRange().getValues();
    
    for (let i = 1; i < otpData.length; i++) {
      const rowEmail = String(otpData[i][0] || '').toLowerCase().trim();
      const rowOTP = String(otpData[i][1] || '').trim();
      const expiresAt = new Date(otpData[i][3]);
      const used = String(otpData[i][4] || 'false').toLowerCase() === 'true';
      
      if (rowEmail === email && rowOTP === otp) {
        const now = new Date();
        
        if (used) {
          return { success: false, message: 'OTP already used.' };
        }
        
        if (now > expiresAt) {
          return { success: false, message: 'OTP has expired.' };
        }
        
        otpSheet.getRange(i + 1, 5).setValue(true);
        SpreadsheetApp.flush();
        
        return { success: true, message: 'OTP verified successfully.' };
      }
    }
    
    return { success: false, message: 'Invalid OTP.' };
  } catch (error) {
    console.error('Verify OTP error: ' + error.toString());
    return { success: false, message: 'OTP verification failed: ' + error.message };
  }
}

function resetPassword(email, newPassword) {
  try {
    email = String(email || '').toLowerCase().trim();
    newPassword = String(newPassword || '').trim();
    
    if (newPassword.length < 6) {
      return { success: false, message: 'Password must be at least 6 characters.' };
    }
    
    const sheet = getOrCreateSheet('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][2] || '').toLowerCase().trim();
      if (rowEmail === email) {
        sheet.getRange(i + 1, 8).setValue(newPassword);
        SpreadsheetApp.flush();
        
        const otpSheet = getOrCreateSheet('OTP');
        const otpData = otpSheet.getDataRange().getValues();
        for (let j = otpData.length - 1; j >= 1; j--) {
          const otpEmail = String(otpData[j][0] || '').toLowerCase().trim();
          if (otpEmail === email) {
            otpSheet.deleteRow(j + 1);
          }
        }
        
        return { success: true, message: 'Password reset successfully.' };
      }
    }
    
    return { success: false, message: 'Email not found.' };
  } catch (error) {
    return { success: false, message: 'Password reset failed: ' + error.message };
  }
}

function changePassword(sessionId, currentPassword, newPassword) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][2] || '').toLowerCase().trim();
      const rowPassword = String(data[i][7] || '').trim();
      
      if (rowEmail === user.email && rowPassword === currentPassword) {
        sheet.getRange(i + 1, 8).setValue(newPassword);
        SpreadsheetApp.flush();
        return { success: true, message: 'Password changed successfully.' };
      }
    }
    
    return { success: false, message: 'Current password is incorrect.' };
  } catch (error) {
    return { success: false, message: 'Password change failed: ' + error.message };
  }
}

// ==========================================
// EMPLOYEE MANAGEMENT
// ==========================================

function addEmployee(sessionId, employeeData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Employees');
    const employeeId = 'EMP' + String(Date.now()).slice(-6);
    
    sheet.appendRow([
      employeeId,
      employeeData.name,
      employeeData.email.toLowerCase().trim(),
      employeeData.phone,
      employeeData.department,
      employeeData.designation,
      employeeData.joinDate,
      employeeData.password || 'password123',
      employeeData.role || 'employee',
      'active',
      employeeData.shift || 'A',
      new Date().toISOString(),
      employeeData.weekOffDay || 'Sunday',
      employeeData.bankName || '',
      employeeData.accountNumber || '',
      employeeData.ifscCode || ''
    ]);
    SpreadsheetApp.flush();
    
    // Initialize leave balance
    const leaveSheet = getOrCreateSheet('LeaveBalance');
    leaveSheet.appendRow([employeeId, 12, 6, 15, 3, 36, new Date().getFullYear()]);
    SpreadsheetApp.flush();
    
    // Initialize roster config
    const rosterConfigSheet = getOrCreateSheet('RosterConfig');
    rosterConfigSheet.appendRow([
      'CFG' + String(Date.now()).slice(-6),
      employeeId,
      employeeData.weekOffDay || 'Sunday',
      'A',
      employeeData.shift || 'A',
      new Date().toISOString()
    ]);
    SpreadsheetApp.flush();
    
    addAuditLog(session.user.id, 'ADD_EMPLOYEE', 'Added employee: ' + employeeId);
    
    return { success: true, message: 'Employee added successfully.', employeeId: employeeId };
  } catch (error) {
    return { success: false, message: 'Failed to add employee: ' + error.message };
  }
}

function updateEmployee(sessionId, employeeId, employeeData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(employeeId).trim()) {
        if (employeeData.name) sheet.getRange(i + 1, 2).setValue(employeeData.name);
        if (employeeData.email) sheet.getRange(i + 1, 3).setValue(employeeData.email.toLowerCase().trim());
        if (employeeData.phone) sheet.getRange(i + 1, 4).setValue(employeeData.phone);
        if (employeeData.department) sheet.getRange(i + 1, 5).setValue(employeeData.department);
        if (employeeData.designation) sheet.getRange(i + 1, 6).setValue(employeeData.designation);
        if (employeeData.joinDate) sheet.getRange(i + 1, 7).setValue(employeeData.joinDate);
        if (employeeData.password) sheet.getRange(i + 1, 8).setValue(employeeData.password);
        if (employeeData.role) sheet.getRange(i + 1, 9).setValue(employeeData.role);
        if (employeeData.status) sheet.getRange(i + 1, 10).setValue(employeeData.status);
        if (employeeData.shift) sheet.getRange(i + 1, 11).setValue(employeeData.shift);
        if (employeeData.weekOffDay) sheet.getRange(i + 1, 13).setValue(employeeData.weekOffDay);
        if (employeeData.bankName !== undefined) sheet.getRange(i + 1, 14).setValue(employeeData.bankName);
        if (employeeData.accountNumber !== undefined) sheet.getRange(i + 1, 15).setValue(employeeData.accountNumber);
        if (employeeData.ifscCode !== undefined) sheet.getRange(i + 1, 16).setValue(employeeData.ifscCode);
        SpreadsheetApp.flush();
        
        addAuditLog(session.user.id, 'UPDATE_EMPLOYEE', 'Updated employee: ' + employeeId);
        return { success: true, message: 'Employee updated successfully.' };
      }
    }
    
    return { success: false, message: 'Employee not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to update employee: ' + error.message };
  }
}

function updateOwnProfile(sessionId, profileData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === user.id) {
        if (profileData.phone) sheet.getRange(i + 1, 4).setValue(profileData.phone);
        if (profileData.bankName !== undefined) sheet.getRange(i + 1, 14).setValue(profileData.bankName);
        if (profileData.accountNumber !== undefined) sheet.getRange(i + 1, 15).setValue(profileData.accountNumber);
        if (profileData.ifscCode !== undefined) sheet.getRange(i + 1, 16).setValue(profileData.ifscCode);
        SpreadsheetApp.flush();
        
        addAuditLog(user.id, 'UPDATE_PROFILE', 'Updated own profile');
        return { success: true, message: 'Profile updated successfully.' };
      }
    }
    
    return { success: false, message: 'Employee not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to update profile: ' + error.message };
  }
}

function getEmployees(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const sheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const employees = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        employees.push({
          id: String(data[i][0] || ''),
          name: String(data[i][1] || ''),
          email: String(data[i][2] || ''),
          phone: String(data[i][3] || ''),
          department: String(data[i][4] || ''),
          designation: String(data[i][5] || ''),
          joinDate: String(data[i][6] || ''),
          role: String(data[i][8] || 'employee'),
          status: String(data[i][9] || 'active'),
          shift: String(data[i][10] || 'A'),
          createdAt: String(data[i][11] || ''),
          weekOffDay: String(data[i][12] || 'Sunday'),
          bankName: String(data[i][13] || ''),
          accountNumber: String(data[i][14] || ''),
          ifscCode: String(data[i][15] || '')
        });
      }
    }
    
    return { success: true, employees: employees };
  } catch (error) {
    return { success: false, message: 'Failed to get employees: ' + error.message };
  }
}

function getEmployeeById(sessionId, employeeId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const sheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(employeeId).trim()) {
        return {
          success: true,
          employee: {
            id: String(data[i][0] || ''),
            name: String(data[i][1] || ''),
            email: String(data[i][2] || ''),
            phone: String(data[i][3] || ''),
            department: String(data[i][4] || ''),
            designation: String(data[i][5] || ''),
            joinDate: String(data[i][6] || ''),
            role: String(data[i][8] || 'employee'),
            status: String(data[i][9] || 'active'),
            shift: String(data[i][10] || 'A'),
            createdAt: String(data[i][11] || ''),
            weekOffDay: String(data[i][12] || 'Sunday'),
            bankName: String(data[i][13] || ''),
            accountNumber: String(data[i][14] || ''),
            ifscCode: String(data[i][15] || '')
          }
        };
      }
    }
    
    return { success: false, message: 'Employee not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to get employee: ' + error.message };
  }
}

// ==========================================
// ATTENDANCE MANAGEMENT WITH GEOFENCING & SELFIE
// ==========================================

function verifyGeofence(latitude, longitude) {
  try {
    const lat = parseFloat(latitude);
    const lng = parseFloat(longitude);
    
    // Calculate distance using Haversine formula
    const R = 6371000; // Earth's radius in meters
    const dLat = (lat - OFFICE_LOCATION.latitude) * Math.PI / 180;
    const dLng = (lng - OFFICE_LOCATION.longitude) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(OFFICE_LOCATION.latitude * Math.PI / 180) * Math.cos(lat * Math.PI / 180) *
              Math.sin(dLng/2) * Math.sin(dLng/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    const distance = R * c;
    
    return {
      valid: distance <= OFFICE_LOCATION.radiusMeters,
      distance: Math.round(distance),
      message: distance <= OFFICE_LOCATION.radiusMeters ? 
               'Location verified' : 
               'You are ' + Math.round(distance - OFFICE_LOCATION.radiusMeters) + 'm outside office area'
    };
  } catch (error) {
    return { valid: false, distance: -1, message: 'Location verification failed' };
  }
}

function punchIn(sessionId, location, selfieData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const now = new Date();
    
    // Normalize user ID for consistent comparison
    const normalizedUserId = String(user.id || '').trim();
    console.log('Punch In attempt - User ID: ' + normalizedUserId + ', Date: ' + today);
    
    // Verify geofence
    let locationObj = {};
    try {
      locationObj = typeof location === 'string' ? JSON.parse(location) : location;
    } catch(e) {
      locationObj = { latitude: 0, longitude: 0 };
    }
    
    const geofenceResult = verifyGeofence(locationObj.latitude, locationObj.longitude);
    console.log('Geofence result: ' + JSON.stringify(geofenceResult));
    
    const sheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    // Check for existing punch-in today (check ALL records)
    let hasActivePunchIn = false;
    let activePunchInTime = '';
    let completedRecords = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      const rowDate = String(data[i][2] || '').trim();
      const rowPunchIn = String(data[i][3] || '').trim();
      const rowPunchOut = String(data[i][4] || '').trim();
      
      // Compare with normalized user ID
      if (rowEmpId === normalizedUserId && rowDate === today) {
        console.log('Found record - PunchIn: ' + rowPunchIn + ', PunchOut: ' + rowPunchOut);
        
        if (rowPunchIn && !rowPunchOut) {
          // Already punched in, not punched out yet - this is an ACTIVE punch-in
          hasActivePunchIn = true;
          activePunchInTime = rowPunchIn;
        } else if (rowPunchIn && rowPunchOut) {
          // This record is completed (both punch in and out exist)
          completedRecords++;
        }
      }
    }
    
    // BLOCK if there's an active punch-in (not punched out yet)
    if (hasActivePunchIn) {
      return { 
        success: false, 
        message: 'You have already punched in today at ' + activePunchInTime + '. Please punch out first before punching in again.' 
      };
    }
    
    // If there are completed records, this is an OT (Overtime) punch-in
    let isOT = completedRecords > 0;
    console.log('Completed records today: ' + completedRecords + ', Is OT: ' + isOT);
    
    const attendanceId = 'ATT' + String(Date.now()).slice(-6);
    const punchInTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    // Determine shift and status from roster
    const rosterResult = getEmployeeRosterForDate(user.id, today);
    let shift = rosterResult.shift || user.shift;
    let isWeekOff = rosterResult.isWeekOff;
    
    let status = 'Present';
    if (isOT) {
      status = 'Overtime';
    } else if (isWeekOff) {
      status = 'Week Off Working';
    } else {
      const shiftStart = SHIFT_TIMES[shift] ? SHIFT_TIMES[shift].start : '06:00';
      const [shiftHour, shiftMin] = shiftStart.split(':').map(Number);
      const punchInHour = now.getHours();
      const punchInMinute = now.getMinutes();
      
      if (punchInHour > shiftHour || (punchInHour === shiftHour && punchInMinute > GRACE_MINUTES)) {
        status = 'Late';
      }
    }
    
    const locationStr = JSON.stringify({
      latitude: locationObj.latitude || 0,
      longitude: locationObj.longitude || 0,
      distance: geofenceResult.distance,
      verified: geofenceResult.valid
    });
    
    sheet.appendRow([
      attendanceId, user.id, today, punchInTime, '', status, shift, 0,
      locationStr, '', selfieData || '', '', ''
    ]);
    SpreadsheetApp.flush();
    
    addAuditLog(user.id, 'PUNCH_IN', 'Punched in at ' + punchInTime + ' (Geofence: ' + (geofenceResult.valid ? 'Valid' : 'Invalid') + ')');
    
    return { 
      success: true, 
      message: 'Punched in successfully at ' + punchInTime, 
      attendanceId: attendanceId,
      geofence: geofenceResult
    };
  } catch (error) {
    console.error('Punch in error: ' + error.toString());
    return { success: false, message: 'Punch in failed: ' + error.message };
  }
}

function punchOut(sessionId, location, selfieData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const now = new Date();
    const punchOutTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    // Normalize user ID for consistent comparison
    const normalizedUserId = String(user.id || '').trim();
    console.log('Punch Out attempt - User ID: ' + normalizedUserId + ', Date: ' + today);
    
    // Verify geofence
    let locationObj = {};
    try {
      locationObj = typeof location === 'string' ? JSON.parse(location) : location;
    } catch(e) {
      locationObj = { latitude: 0, longitude: 0 };
    }
    
    const geofenceResult = verifyGeofence(locationObj.latitude, locationObj.longitude);
    
    const sheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    let punchInRowIndex = -1;
    let punchInTime = '';
    
    // Find today's ACTIVE punch-in record (punched in but not out)
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      const rowDate = String(data[i][2] || '').trim();
      const rowPunchIn = String(data[i][3] || '').trim();
      const rowPunchOut = String(data[i][4] || '').trim();
      
      if (rowEmpId === normalizedUserId && rowDate === today) {
        console.log('Found record - PunchIn: ' + rowPunchIn + ', PunchOut: ' + rowPunchOut);
        if (rowPunchIn && !rowPunchOut) {
          punchInRowIndex = i;
          punchInTime = rowPunchIn;
          // Don't break - find the LATEST active punch-in
        }
      }
    }
    
    if (punchInRowIndex === -1) {
      return { success: false, message: 'No active punch-in found for today. Please punch in first.' };
    }
    
    // Calculate work hours
    const punchInDate = new Date('1970-01-01T' + punchInTime);
    const punchOutDate = new Date('1970-01-01T' + punchOutTime);
    let workHours = (punchOutDate - punchInDate) / (1000 * 60 * 60);
    if (workHours < 0) workHours += 24; // Handle overnight shifts
    
    // Get current status and update if needed
    let currentStatus = String(data[punchInRowIndex][5] || 'Present');
    let status = currentStatus;
    
    // Only update status for non-OT records
    if (currentStatus !== 'Overtime') {
      if (workHours < 4) {
        status = 'Half Day';
      } else if (workHours < 8 && currentStatus !== 'Late') {
        status = 'Early Out';
      }
    }
    
    const locationStr = JSON.stringify({
      latitude: locationObj.latitude || 0,
      longitude: locationObj.longitude || 0,
      distance: geofenceResult.distance,
      verified: geofenceResult.valid
    });
    
    sheet.getRange(punchInRowIndex + 1, 5).setValue(punchOutTime);
    sheet.getRange(punchInRowIndex + 1, 6).setValue(status);
    sheet.getRange(punchInRowIndex + 1, 8).setValue(workHours.toFixed(2));
    sheet.getRange(punchInRowIndex + 1, 10).setValue(locationStr);
    sheet.getRange(punchInRowIndex + 1, 12).setValue(selfieData || '');
    SpreadsheetApp.flush();
    
    addAuditLog(normalizedUserId, 'PUNCH_OUT', 'Punched out at ' + punchOutTime + ' (Hours: ' + workHours.toFixed(2) + ')');
    
    return { 
      success: true, 
      message: 'Punched out successfully at ' + punchOutTime, 
      workHours: workHours.toFixed(2),
      geofence: geofenceResult
    };
  } catch (error) {
    console.error('Punch out error: ' + error.toString());
    return { success: false, message: 'Punch out failed: ' + error.message };
  }
}

function getEmployeeRosterForDate(employeeId, dateStr) {
  try {
    const sheet = getOrCreateSheet('Roster');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() === String(employeeId).trim() && 
          String(data[i][3] || '').trim() === String(dateStr).trim()) {
        return {
          shift: String(data[i][5] || 'A'),
          isWeekOff: String(data[i][6] || '').toLowerCase() === 'true',
          isBuffer: String(data[i][7] || '').toLowerCase() === 'true'
        };
      }
    }
    
    return { shift: 'A', isWeekOff: false, isBuffer: false };
  } catch (error) {
    return { shift: 'A', isWeekOff: false, isBuffer: false };
  }
}

function getTodayAttendance(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const normalizedUserId = String(user.id || '').trim();
    
    console.log('getTodayAttendance - User ID: ' + normalizedUserId + ', Date: ' + today);
    
    const sheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    // Collect ALL attendance records for today
    let allRecords = [];
    let activeRecord = null; // Record with punch-in but no punch-out
    let latestCompletedRecord = null; // Most recent completed record
    
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      const rowDate = String(data[i][2] || '').trim();
      
      if (rowEmpId === normalizedUserId && rowDate === today) {
        const record = {
          id: String(data[i][0] || ''),
          employeeId: String(data[i][1] || ''),
          date: String(data[i][2] || ''),
          punchIn: String(data[i][3] || ''),
          punchOut: String(data[i][4] || ''),
          status: String(data[i][5] || ''),
          shift: String(data[i][6] || ''),
          workHours: parseFloat(data[i][7]) || 0,
          punchInLocation: String(data[i][8] || ''),
          punchOutLocation: String(data[i][9] || ''),
          punchInSelfie: String(data[i][10] || ''),
          punchOutSelfie: String(data[i][11] || '')
        };
        
        allRecords.push(record);
        
        // Check if this is an active punch-in (punched in but not out)
        if (record.punchIn && !record.punchOut) {
          activeRecord = record;
        }
        // Check if this is a completed record
        else if (record.punchIn && record.punchOut) {
          if (!latestCompletedRecord || record.id > latestCompletedRecord.id) {
            latestCompletedRecord = record;
          }
        }
      }
    }
    
    console.log('Total records today: ' + allRecords.length + ', Active: ' + (activeRecord ? 'Yes' : 'No'));
    
    // Determine the attendance state
    // 1. If there's an active punch-in (not punched out), return that record with state 'punched_in'
    // 2. If all records are completed, return the latest one with state 'completed' (or 'ot' if multiple)
    // 3. If no records, return null with state 'not_punched'
    
    if (activeRecord) {
      // Employee is currently punched in
      return {
        success: true,
        attendance: activeRecord,
        state: 'punched_in',
        allRecords: allRecords,
        totalWorkHours: allRecords.reduce((sum, r) => sum + (r.workHours || 0), 0)
      };
    } else if (latestCompletedRecord) {
      // All records are completed
      let state = allRecords.length > 1 ? 'ot_completed' : 'completed';
      return {
        success: true,
        attendance: latestCompletedRecord,
        state: state,
        allRecords: allRecords,
        totalWorkHours: allRecords.reduce((sum, r) => sum + (r.workHours || 0), 0)
      };
    } else {
      // No attendance records for today
      return { 
        success: true, 
        attendance: null, 
        state: 'not_punched',
        allRecords: [],
        totalWorkHours: 0
      };
    }
  } catch (error) {
    console.error('getTodayAttendance error: ' + error.toString());
    return { success: false, message: 'Failed to get attendance: ' + error.message };
  }
}

function getAttendanceHistory(sessionId, startDate, endDate) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const history = [];
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      const rowDate = String(data[i][2] || '').trim();
      
      if (rowEmpId === user.id) {
        if ((!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate)) {
          history.push({
            id: String(data[i][0] || ''),
            date: rowDate,
            punchIn: String(data[i][3] || ''),
            punchOut: String(data[i][4] || ''),
            status: String(data[i][5] || ''),
            shift: String(data[i][6] || ''),
            workHours: parseFloat(data[i][7]) || 0
          });
        }
      }
    }
    
    history.sort((a, b) => new Date(b.date) - new Date(a.date));
    return { success: true, history: history };
  } catch (error) {
    return { success: false, message: 'Failed to get attendance history: ' + error.message };
  }
}

function getAttendanceLogs(sessionId, startDate, endDate) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const logs = [];
    for (let i = 1; i < data.length; i++) {
      const rowDate = String(data[i][2] || '').trim();
      
      if ((!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate)) {
        logs.push({
          id: String(data[i][0] || ''),
          employeeId: String(data[i][1] || ''),
          date: rowDate,
          punchIn: String(data[i][3] || ''),
          punchOut: String(data[i][4] || ''),
          status: String(data[i][5] || ''),
          shift: String(data[i][6] || ''),
          workHours: parseFloat(data[i][7]) || 0
        });
      }
    }
    
    logs.sort((a, b) => new Date(b.date) - new Date(a.date));
    return { success: true, logs: logs };
  } catch (error) {
    return { success: false, message: 'Failed to get attendance logs: ' + error.message };
  }
}

// ==========================================
// PUNCH CORRECTION REQUESTS
// ==========================================

function requestPunchCorrection(sessionId, correctionData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const requestId = 'PCR' + String(Date.now()).slice(-6);
    
    const sheet = getOrCreateSheet('PunchCorrections');
    sheet.appendRow([
      requestId,
      user.id,
      correctionData.date,
      correctionData.currentPunchIn || '',
      correctionData.currentPunchOut || '',
      correctionData.requestedPunchIn || '',
      correctionData.requestedPunchOut || '',
      correctionData.reason || '',
      'Pending',
      new Date().toISOString(),
      '',
      '',
      ''
    ]);
    SpreadsheetApp.flush();
    
    addAuditLog(user.id, 'PUNCH_CORRECTION_REQUEST', 'Requested correction for ' + correctionData.date);
    
    return { success: true, message: 'Punch correction request submitted.', requestId: requestId };
  } catch (error) {
    return { success: false, message: 'Failed to submit request: ' + error.message };
  }
}

function getPunchCorrections(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('PunchCorrections');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const corrections = [];
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      const rowStatus = String(data[i][8] || '').trim();
      
      if (user.role === 'admin' || rowEmpId === user.id) {
        corrections.push({
          requestId: String(data[i][0] || ''),
          employeeId: rowEmpId,
          date: String(data[i][2] || ''),
          currentPunchIn: String(data[i][3] || ''),
          currentPunchOut: String(data[i][4] || ''),
          requestedPunchIn: String(data[i][5] || ''),
          requestedPunchOut: String(data[i][6] || ''),
          reason: String(data[i][7] || ''),
          status: rowStatus,
          requestDate: String(data[i][9] || ''),
          approvedBy: String(data[i][10] || ''),
          approvedDate: String(data[i][11] || ''),
          comments: String(data[i][12] || '')
        });
      }
    }
    
    corrections.sort((a, b) => new Date(b.requestDate) - new Date(a.requestDate));
    return { success: true, corrections: corrections };
  } catch (error) {
    return { success: false, message: 'Failed to get punch corrections: ' + error.message };
  }
}

function approvePunchCorrection(sessionId, requestId, comments) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('PunchCorrections');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(requestId).trim()) {
        sheet.getRange(i + 1, 9).setValue('Approved');
        sheet.getRange(i + 1, 11).setValue(session.user.id);
        sheet.getRange(i + 1, 12).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 13).setValue(comments || '');
        
        const employeeId = String(data[i][1] || '').trim();
        const date = String(data[i][2] || '').trim();
        const newPunchIn = String(data[i][5] || '').trim();
        const newPunchOut = String(data[i][6] || '').trim();
        
        const attSheet = getOrCreateSheet('Attendance');
        SpreadsheetApp.flush();
        const attData = attSheet.getDataRange().getValues();
        
        for (let j = 1; j < attData.length; j++) {
          if (String(attData[j][1] || '').trim() === employeeId && String(attData[j][2] || '').trim() === date) {
            if (newPunchIn) attSheet.getRange(j + 1, 4).setValue(newPunchIn);
            if (newPunchOut) attSheet.getRange(j + 1, 5).setValue(newPunchOut);
            
            if (newPunchIn && newPunchOut) {
              const punchInTime = new Date('1970-01-01T' + newPunchIn);
              const punchOutDate = new Date('1970-01-01T' + newPunchOut);
              const workHours = (punchOutDate - punchInTime) / (1000 * 60 * 60);
              attSheet.getRange(j + 1, 8).setValue(workHours.toFixed(2));
            }
            break;
          }
        }
        
        SpreadsheetApp.flush();
        addAuditLog(session.user.id, 'APPROVE_CORRECTION', 'Approved correction ' + requestId);
        
        return { success: true, message: 'Punch correction approved.' };
      }
    }
    
    return { success: false, message: 'Request not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to approve request: ' + error.message };
  }
}

function rejectPunchCorrection(sessionId, requestId, comments) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('PunchCorrections');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(requestId).trim()) {
        sheet.getRange(i + 1, 9).setValue('Rejected');
        sheet.getRange(i + 1, 11).setValue(session.user.id);
        sheet.getRange(i + 1, 12).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 13).setValue(comments || '');
        SpreadsheetApp.flush();
        
        addAuditLog(session.user.id, 'REJECT_CORRECTION', 'Rejected correction ' + requestId);
        
        return { success: true, message: 'Punch correction rejected.' };
      }
    }
    
    return { success: false, message: 'Request not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to reject request: ' + error.message };
  }
}

// ==========================================
// LEAVE MANAGEMENT
// ==========================================

function getLeaveBalance(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('LeaveBalance');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    const currentYear = new Date().getFullYear();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === user.id && parseInt(data[i][6]) === currentYear) {
        return {
          success: true,
          balance: {
            casual: parseInt(data[i][1]) || 0,
            sick: parseInt(data[i][2]) || 0,
            earned: parseInt(data[i][3]) || 0,
            compOff: parseInt(data[i][4]) || 0,
            total: parseInt(data[i][5]) || 0
          }
        };
      }
    }
    
    sheet.appendRow([user.id, 12, 6, 15, 3, 36, currentYear]);
    SpreadsheetApp.flush();
    
    return {
      success: true,
      balance: { casual: 12, sick: 6, earned: 15, compOff: 3, total: 36 }
    };
  } catch (error) {
    return { success: false, message: 'Failed to get leave balance: ' + error.message };
  }
}

function applyLeave(sessionId, leaveData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    
    const startDate = new Date(leaveData.startDate);
    const endDate = new Date(leaveData.endDate);
    const days = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    
    const balanceResult = getLeaveBalance(sessionId);
    if (!balanceResult.success) {
      return balanceResult;
    }
    
    const balance = balanceResult.balance;
    const leaveType = String(leaveData.leaveType || '').toLowerCase();
    
    if (leaveType === 'casual' && balance.casual < days) {
      return { success: false, message: 'Insufficient casual leave balance.' };
    }
    if (leaveType === 'sick' && balance.sick < days) {
      return { success: false, message: 'Insufficient sick leave balance.' };
    }
    
    const leaveId = 'LV' + String(Date.now()).slice(-6);
    
    const sheet = getOrCreateSheet('Leaves');
    sheet.appendRow([
      leaveId,
      user.id,
      leaveData.leaveType,
      leaveData.startDate,
      leaveData.endDate,
      days,
      leaveData.reason || '',
      'Pending',
      new Date().toISOString(),
      '',
      '',
      ''
    ]);
    SpreadsheetApp.flush();
    
    addAuditLog(user.id, 'APPLY_LEAVE', 'Applied leave ' + leaveId);
    
    return { success: true, message: 'Leave application submitted.', leaveId: leaveId };
  } catch (error) {
    return { success: false, message: 'Failed to apply leave: ' + error.message };
  }
}

function getLeaveApplications(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('Leaves');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const leaves = [];
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      
      if (user.role === 'admin' || rowEmpId === user.id) {
        leaves.push({
          leaveId: String(data[i][0] || ''),
          employeeId: rowEmpId,
          leaveType: String(data[i][2] || ''),
          startDate: String(data[i][3] || ''),
          endDate: String(data[i][4] || ''),
          days: parseInt(data[i][5]) || 0,
          reason: String(data[i][6] || ''),
          status: String(data[i][7] || ''),
          appliedDate: String(data[i][8] || ''),
          approvedBy: String(data[i][9] || ''),
          approvedDate: String(data[i][10] || ''),
          comments: String(data[i][11] || '')
        });
      }
    }
    
    leaves.sort((a, b) => new Date(b.appliedDate) - new Date(a.appliedDate));
    return { success: true, leaves: leaves };
  } catch (error) {
    return { success: false, message: 'Failed to get leave applications: ' + error.message };
  }
}

function approveLeave(sessionId, leaveId, comments) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Leaves');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(leaveId).trim()) {
        const employeeId = String(data[i][1] || '').trim();
        const days = parseInt(data[i][5]) || 0;
        const leaveType = String(data[i][2] || '').toLowerCase();
        
        sheet.getRange(i + 1, 8).setValue('Approved');
        sheet.getRange(i + 1, 10).setValue(session.user.id);
        sheet.getRange(i + 1, 11).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 12).setValue(comments || '');
        
        const balanceSheet = getOrCreateSheet('LeaveBalance');
        SpreadsheetApp.flush();
        const balanceData = balanceSheet.getDataRange().getValues();
        const currentYear = new Date().getFullYear();
        
        for (let j = 1; j < balanceData.length; j++) {
          if (String(balanceData[j][0] || '').trim() === employeeId && parseInt(balanceData[j][6]) === currentYear) {
            let currentBalance = parseInt(balanceData[j][leaveType === 'casual' ? 1 : 2]) || 0;
            balanceSheet.getRange(j + 1, leaveType === 'casual' ? 2 : 3).setValue(Math.max(0, currentBalance - days));
            
            let total = parseInt(balanceData[j][5]) || 0;
            balanceSheet.getRange(j + 1, 6).setValue(Math.max(0, total - days));
            break;
          }
        }
        
        SpreadsheetApp.flush();
        addAuditLog(session.user.id, 'APPROVE_LEAVE', 'Approved leave ' + leaveId);
        
        return { success: true, message: 'Leave approved.' };
      }
    }
    
    return { success: false, message: 'Leave not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to approve leave: ' + error.message };
  }
}

function rejectLeave(sessionId, leaveId, comments) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Leaves');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(leaveId).trim()) {
        sheet.getRange(i + 1, 8).setValue('Rejected');
        sheet.getRange(i + 1, 10).setValue(session.user.id);
        sheet.getRange(i + 1, 11).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 12).setValue(comments || '');
        SpreadsheetApp.flush();
        
        addAuditLog(session.user.id, 'REJECT_LEAVE', 'Rejected leave ' + leaveId);
        
        return { success: true, message: 'Leave rejected.' };
      }
    }
    
    return { success: false, message: 'Leave not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to reject leave: ' + error.message };
  }
}

// ==========================================
// HOLIDAY MANAGEMENT
// ==========================================

function getHolidays(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const sheet = getOrCreateSheet('Holidays');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const holidays = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        holidays.push({
          holidayId: String(data[i][0] || ''),
          name: String(data[i][1] || ''),
          date: String(data[i][2] || ''),
          type: String(data[i][3] || ''),
          optional: String(data[i][4] || '').toLowerCase() === 'true',
          createdAt: String(data[i][5] || '')
        });
      }
    }
    
    holidays.sort((a, b) => new Date(a.date) - new Date(b.date));
    return { success: true, holidays: holidays };
  } catch (error) {
    return { success: false, message: 'Failed to get holidays: ' + error.message };
  }
}

function addHoliday(sessionId, holidayData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const holidayId = 'HD' + String(Date.now()).slice(-6);
    
    const sheet = getOrCreateSheet('Holidays');
    sheet.appendRow([
      holidayId,
      holidayData.name,
      holidayData.date,
      holidayData.type || 'National',
      holidayData.optional || false,
      new Date().toISOString()
    ]);
    SpreadsheetApp.flush();
    
    addAuditLog(session.user.id, 'ADD_HOLIDAY', 'Added holiday: ' + holidayData.name);
    
    return { success: true, message: 'Holiday added successfully.' };
  } catch (error) {
    return { success: false, message: 'Failed to add holiday: ' + error.message };
  }
}

function deleteHoliday(sessionId, holidayId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Holidays');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || '').trim() === String(holidayId).trim()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        addAuditLog(session.user.id, 'DELETE_HOLIDAY', 'Deleted holiday: ' + holidayId);
        return { success: true, message: 'Holiday deleted.' };
      }
    }
    
    return { success: false, message: 'Holiday not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to delete holiday: ' + error.message };
  }
}

// ==========================================
// ROSTER MANAGEMENT WITH FIXED POLICY
// ==========================================

function getRosterConfig(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const sheet = getOrCreateSheet('RosterConfig');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const configs = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) {
        configs.push({
          configId: String(data[i][0] || ''),
          employeeId: String(data[i][1] || ''),
          weekOffDay: String(data[i][2] || 'Sunday'),
          rotationSequence: String(data[i][3] || 'A'),
          currentShift: String(data[i][4] || 'A'),
          lastUpdated: String(data[i][5] || '')
        });
      }
    }
    
    return { success: true, configs: configs };
  } catch (error) {
    return { success: false, message: 'Failed to get roster config: ' + error.message };
  }
}

function updateRosterConfig(sessionId, employeeId, configData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('RosterConfig');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() === String(employeeId).trim()) {
        if (configData.weekOffDay) sheet.getRange(i + 1, 3).setValue(configData.weekOffDay);
        if (configData.currentShift) sheet.getRange(i + 1, 5).setValue(configData.currentShift);
        sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
        SpreadsheetApp.flush();
        return { success: true, message: 'Roster config updated.' };
      }
    }
    
    // Create new config if not exists
    sheet.appendRow([
      'CFG' + String(Date.now()).slice(-6),
      employeeId,
      configData.weekOffDay || 'Sunday',
      'A',
      configData.currentShift || 'A',
      new Date().toISOString()
    ]);
    SpreadsheetApp.flush();
    
    return { success: true, message: 'Roster config created.' };
  } catch (error) {
    return { success: false, message: 'Failed to update roster config: ' + error.message };
  }
}

function getRoster(sessionId, startDate, endDate) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const sheet = getOrCreateSheet('Roster');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const roster = [];
    for (let i = 1; i < data.length; i++) {
      const rowDate = String(data[i][3] || '').trim();
      
      if ((!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate)) {
        roster.push({
          rosterId: String(data[i][0] || ''),
          employeeId: String(data[i][1] || ''),
          name: String(data[i][2] || ''),
          date: rowDate,
          dayOfWeek: String(data[i][4] || ''),
          shift: String(data[i][5] || ''),
          isWeekOff: String(data[i][6] || '').toLowerCase() === 'true',
          isBuffer: String(data[i][7] || '').toLowerCase() === 'true',
          createdAt: String(data[i][8] || '')
        });
      }
    }
    
    roster.sort((a, b) => new Date(a.date) - new Date(b.date));
    return { success: true, roster: roster };
  } catch (error) {
    return { success: false, message: 'Failed to get roster: ' + error.message };
  }
}

function getMyShiftDetails(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const today = new Date();
    const shifts = [];
    
    // Generate 14 days roster
    for (let i = 0; i < 14; i++) {
      const date = new Date(today);
      date.setDate(today.getDate() + i);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      const dayName = dayNames[date.getDay()];
      
      // Get roster for this date
      const rosterSheet = getOrCreateSheet('Roster');
      SpreadsheetApp.flush();
      const rosterData = rosterSheet.getDataRange().getValues();
      
      let shift = user.shift;
      let isWeekOff = false;
      let isBuffer = false;
      
      for (let j = 1; j < rosterData.length; j++) {
        if (String(rosterData[j][1] || '').trim() === user.id && String(rosterData[j][3] || '').trim() === dateStr) {
          shift = String(rosterData[j][5] || shift);
          isWeekOff = String(rosterData[j][6] || '').toLowerCase() === 'true';
          isBuffer = String(rosterData[j][7] || '').toLowerCase() === 'true';
          break;
        }
      }
      
      const shiftTime = SHIFT_TIMES[shift] || SHIFT_TIMES['A'];
      
      shifts.push({
        date: dateStr,
        day: dayName,
        shift: isWeekOff ? 'OFF' : (isBuffer ? 'Buffer' : shift),
        time: isWeekOff ? 'Week Off' : (isBuffer ? 'Standby' : shiftTime.start + ' - ' + shiftTime.end),
        isWeekOff: isWeekOff,
        isBuffer: isBuffer,
        isToday: i === 0
      });
    }
    
    return { success: true, shifts: shifts };
  } catch (error) {
    return { success: false, message: 'Failed to get shift details: ' + error.message };
  }
}

function generateAutoRoster(sessionId, startDate, endDate, weekOffCount) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    console.log('generateAutoRoster called with startDate: ' + startDate + ', endDate: ' + endDate + ', weekOffCount: ' + weekOffCount);
    
    // Get all active employees with their configs
    const empSheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const empData = empSheet.getDataRange().getValues();
    
    const employees = [];
    for (let i = 1; i < empData.length; i++) {
      const status = String(empData[i][9] || '').toLowerCase().trim();
      if (status === 'active') {
        employees.push({
          id: String(empData[i][0] || '').trim(),
          name: String(empData[i][1] || '').trim(),
          shift: String(empData[i][10] || 'A').trim(),
          weekOffDay: String(empData[i][12] || 'Sunday').trim()
        });
      }
    }
    
    console.log('Found ' + employees.length + ' active employees');
    
    if (employees.length === 0) {
      return { success: false, message: 'No active employees found.' };
    }
    
    // Parse dates
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      return { success: false, message: 'Invalid date format.' };
    }
    
    if (start > end) {
      return { success: false, message: 'Start date must be before end date.' };
    }
    
    // Clear existing roster for date range
    const rosterSheet = getOrCreateSheet('Roster');
    SpreadsheetApp.flush();
    const existingData = rosterSheet.getDataRange().getValues();
    
    for (let i = existingData.length - 1; i >= 1; i--) {
      const rowDate = String(existingData[i][3] || '').trim();
      if (rowDate >= startDate && rowDate <= endDate) {
        rosterSheet.deleteRow(i + 1);
      }
    }
    SpreadsheetApp.flush();
    
    // Roster Policy:
    // - Shifts: A (06:00-14:00), B (14:00-22:00), C (22:00-06:00)
    // - Min 3 employees per shift
    // - Week off configurable (1/2/3/none)
    // - 1 buffer employee
    // - Rotation: A → OFF → B → OFF → C → OFF → A
    
    const totalDays = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
    const shifts = ['A', 'B', 'C'];
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    
    // Initialize employee rotation states
    const employeeStates = {};
    const shiftOrder = ['A', 'OFF', 'B', 'OFF', 'C', 'OFF'];
    
    employees.forEach((emp, index) => {
      // Distribute employees across rotation starting points
      const startIndex = index % shiftOrder.length;
      employeeStates[emp.id] = {
        currentIndex: startIndex,
        daysInState: 0,
        weekOffDay: emp.weekOffDay
      };
    });
    
    const newRosterEntries = [];
    
    // Generate roster for each day
    for (let d = 0; d < totalDays; d++) {
      const currentDate = new Date(start);
      currentDate.setDate(start.getDate() + d);
      const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dayOfWeek = dayNames[currentDate.getDay()];
      
      // Count employees per shift
      const shiftCounts = { 'A': 0, 'B': 0, 'C': 0 };
      let bufferCount = 0;
      let weekOffCount_today = 0;
      
      const dayAssignments = [];
      
      // First pass: assign based on rotation and week off day
      employees.forEach(emp => {
        const state = employeeStates[emp.id];
        let assignedShift = null;
        let isWeekOff = false;
        let isBuffer = false;
        
        // Check if it's employee's designated week off day
        if (emp.weekOffDay === dayOfWeek) {
          isWeekOff = true;
          weekOffCount_today++;
        } else {
          // Get current rotation state
          const currentState = shiftOrder[state.currentIndex];
          
          if (currentState === 'OFF') {
            isWeekOff = true;
            weekOffCount_today++;
            state.daysInState++;
            
            // Move to next state after 1 day off
            if (state.daysInState >= 1) {
              state.currentIndex = (state.currentIndex + 1) % shiftOrder.length;
              state.daysInState = 0;
            }
          } else {
            assignedShift = currentState;
            
            // Check if shift has capacity
            if (shiftCounts[assignedShift] < 3) {
              shiftCounts[assignedShift]++;
              state.daysInState++;
              
              // Move to next state after working enough days
              // (we'll check this after the shift)
            } else if (bufferCount < 1) {
              isBuffer = true;
              bufferCount++;
            } else {
              // Force to another shift with capacity
              for (let s of shifts) {
                if (shiftCounts[s] < 4) { // Allow extra
                  assignedShift = s;
                  shiftCounts[s]++;
                  break;
                }
              }
            }
          }
        }
        
        dayAssignments.push({
          employeeId: emp.id,
          name: emp.name,
          date: dateStr,
          dayOfWeek: dayOfWeek,
          shift: assignedShift || 'A',
          isWeekOff: isWeekOff,
          isBuffer: isBuffer
        });
      });
      
      // Second pass: ensure minimum 3 per shift
      // Reassign buffer/excess to shifts needing more
      shifts.forEach(shift => {
        if (shiftCounts[shift] < 3) {
          const needed = 3 - shiftCounts[shift];
          // Find buffer or extra employees to reassign
          for (let assignment of dayAssignments) {
            if (needed <= 0) break;
            if (assignment.isBuffer && !assignment.isWeekOff) {
              assignment.shift = shift;
              assignment.isBuffer = false;
              shiftCounts[shift]++;
            }
          }
        }
      });
      
      // Add entries
      dayAssignments.forEach(a => {
        const rosterId = 'RST' + String(Date.now()).slice(-6) + '_' + a.employeeId + '_' + d;
        newRosterEntries.push([
          rosterId,
          a.employeeId,
          a.name,
          a.date,
          a.dayOfWeek,
          a.shift,
          a.isWeekOff ? 'true' : 'false',
          a.isBuffer ? 'true' : 'false',
          new Date().toISOString()
        ]);
      });
      
      // Update rotation states for next day
      employees.forEach(emp => {
        const state = employeeStates[emp.id];
        const assignment = dayAssignments.find(a => a.employeeId === emp.id);
        
        if (!assignment.isWeekOff && !assignment.isBuffer) {
          // After working a shift, check if should rotate
          // Simple rotation: move to next state each day
          state.currentIndex = (state.currentIndex + 1) % shiftOrder.length;
          state.daysInState = 0;
        }
      });
    }
    
    // Write all entries to sheet
    if (newRosterEntries.length > 0) {
      rosterSheet.getRange(rosterSheet.getLastRow() + 1, 1, newRosterEntries.length, 9).setValues(newRosterEntries);
    }
    SpreadsheetApp.flush();
    
    console.log('Generated ' + newRosterEntries.length + ' roster entries');
    
    addAuditLog(session.user.id, 'GENERATE_ROSTER', 'Generated roster from ' + startDate + ' to ' + endDate);
    
    return { 
      success: true, 
      message: 'Roster generated successfully. ' + newRosterEntries.length + ' entries created for ' + totalDays + ' days.',
      entriesCount: newRosterEntries.length
    };
  } catch (error) {
    console.error('generateAutoRoster error: ' + error.toString());
    return { success: false, message: 'Failed to generate roster: ' + error.message };
  }
}

function updateRosterEntry(sessionId, rosterData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const sheet = getOrCreateSheet('Roster');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() === String(rosterData.employeeId).trim() && 
          String(data[i][3] || '').trim() === String(rosterData.date).trim()) {
        sheet.getRange(i + 1, 6).setValue(rosterData.shift);
        sheet.getRange(i + 1, 7).setValue(rosterData.isWeekOff ? 'true' : 'false');
        sheet.getRange(i + 1, 8).setValue(rosterData.isBuffer ? 'true' : 'false');
        SpreadsheetApp.flush();
        return { success: true, message: 'Roster updated.' };
      }
    }
    
    // Create new entry
    const empSheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const empData = empSheet.getDataRange().getValues();
    let empName = '';
    
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][0] || '').trim() === String(rosterData.employeeId).trim()) {
        empName = String(empData[i][1] || '');
        break;
      }
    }
    
    const date = new Date(rosterData.date);
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    
    sheet.appendRow([
      'RST' + String(Date.now()).slice(-6),
      rosterData.employeeId,
      empName,
      rosterData.date,
      dayNames[date.getDay()],
      rosterData.shift,
      rosterData.isWeekOff ? 'true' : 'false',
      rosterData.isBuffer ? 'true' : 'false',
      new Date().toISOString()
    ]);
    SpreadsheetApp.flush();
    
    return { success: true, message: 'Roster entry added.' };
  } catch (error) {
    return { success: false, message: 'Failed to update roster: ' + error.message };
  }
}

// ==========================================
// WARNING LETTERS MANAGEMENT
// ==========================================

function issueWarning(sessionId, warningData) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    const warningId = 'WRN' + String(Date.now()).slice(-6);
    const sheet = getOrCreateSheet('Warnings');
    
    sheet.appendRow([
      warningId,
      warningData.employeeId,
      warningData.issueDate || new Date().toISOString().split('T')[0],
      warningData.reason,
      warningData.description || '',
      session.user.id,
      'Active',
      new Date().toISOString()
    ]);
    SpreadsheetApp.flush();
    
    addAuditLog(session.user.id, 'ISSUE_WARNING', 'Issued warning ' + warningId + ' to ' + warningData.employeeId);
    
    return { success: true, message: 'Warning issued successfully.', warningId: warningId };
  } catch (error) {
    return { success: false, message: 'Failed to issue warning: ' + error.message };
  }
}

function getWarnings(sessionId, employeeId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('Warnings');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    const warnings = [];
    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][1] || '').trim();
      
      // Admin sees all, employee sees only their own
      if (user.role === 'admin' || rowEmpId === user.id) {
        // If employeeId filter is provided, filter by it
        if (!employeeId || rowEmpId === employeeId) {
          warnings.push({
            warningId: String(data[i][0] || ''),
            employeeId: rowEmpId,
            issueDate: String(data[i][2] || ''),
            reason: String(data[i][3] || ''),
            description: String(data[i][4] || ''),
            issuedBy: String(data[i][5] || ''),
            status: String(data[i][6] || 'Active'),
            createdAt: String(data[i][7] || '')
          });
        }
      }
    }
    
    warnings.sort((a, b) => new Date(b.issueDate) - new Date(a.issueDate));
    return { success: true, warnings: warnings };
  } catch (error) {
    return { success: false, message: 'Failed to get warnings: ' + error.message };
  }
}

function generateWarningPDF(sessionId, warningId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const sheet = getOrCreateSheet('Warnings');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(warningId).trim()) {
        const warning = {
          warningId: String(data[i][0] || ''),
          employeeId: String(data[i][1] || ''),
          issueDate: String(data[i][2] || ''),
          reason: String(data[i][3] || ''),
          description: String(data[i][4] || ''),
          issuedBy: String(data[i][5] || ''),
          status: String(data[i][6] || 'Active')
        };
        
        // Check permission
        if (user.role !== 'admin' && warning.employeeId !== user.id) {
          return { success: false, message: 'Unauthorized access.' };
        }
        
        // Get employee name
        const empSheet = getOrCreateSheet('Employees');
        SpreadsheetApp.flush();
        const empData = empSheet.getDataRange().getValues();
        let employeeName = warning.employeeId;
        let department = '';
        
        for (let j = 1; j < empData.length; j++) {
          if (String(empData[j][0] || '').trim() === warning.employeeId) {
            employeeName = String(empData[j][1] || '');
            department = String(empData[j][4] || '');
            break;
          }
        }
        
        // Get issuer name
        let issuerName = 'Administrator';
        for (let j = 1; j < empData.length; j++) {
          if (String(empData[j][0] || '').trim() === warning.issuedBy) {
            issuerName = String(empData[j][1] || '');
            break;
          }
        }
        
        // Generate HTML content for PDF
        const htmlContent = `
          <html>
            <head>
              <style>
                body { font-family: Arial, sans-serif; padding: 40px; }
                .header { text-align: center; margin-bottom: 40px; }
                .title { font-size: 24px; font-weight: bold; color: #333; }
                .subtitle { font-size: 16px; color: #666; margin-top: 10px; }
                .content { margin-top: 30px; }
                .field { margin-bottom: 20px; }
                .label { font-weight: bold; color: #333; }
                .value { color: #666; margin-top: 5px; }
                .warning-box { background: #fff3cd; border-left: 4px solid #fbbc04; padding: 15px; margin: 20px 0; }
                .footer { margin-top: 50px; text-align: center; color: #999; font-size: 12px; }
              </style>
            </head>
            <body>
              <div class="header">
                <div class="title">${COMPANY_NAME}</div>
                <div class="subtitle">WARNING LETTER</div>
              </div>
              
              <div class="content">
                <div class="field">
                  <div class="label">Warning ID:</div>
                  <div class="value">${warning.warningId}</div>
                </div>
                
                <div class="field">
                  <div class="label">Date:</div>
                  <div class="value">${warning.issueDate}</div>
                </div>
                
                <div class="field">
                  <div class="label">To:</div>
                  <div class="value">${employeeName} (${warning.employeeId})</div>
                </div>
                
                <div class="field">
                  <div class="label">Department:</div>
                  <div class="value">${department}</div>
                </div>
                
                <div class="warning-box">
                  <div class="label">Reason for Warning:</div>
                  <div class="value" style="font-size: 18px; margin-top: 10px;">${warning.reason}</div>
                </div>
                
                <div class="field">
                  <div class="label">Description:</div>
                  <div class="value">${warning.description || 'No additional details provided.'}</div>
                </div>
                
                <div class="field">
                  <div class="label">Issued By:</div>
                  <div class="value">${issuerName}</div>
                </div>
                
                <div class="field">
                  <div class="label">Status:</div>
                  <div class="value">${warning.status}</div>
                </div>
              </div>
              
              <div class="footer">
                <p>This is an official warning letter from ${COMPANY_NAME}.</p>
                <p>Generated on: ${new Date().toLocaleString()}</p>
              </div>
            </body>
          </html>
        `;
        
        return { success: true, htmlContent: htmlContent, warning: warning };
      }
    }
    
    return { success: false, message: 'Warning not found.' };
  } catch (error) {
    return { success: false, message: 'Failed to generate PDF: ' + error.message };
  }
}

// ==========================================
// DASHBOARD DATA
// ==========================================

function getEmployeeDashboard(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid) {
      return { success: false, message: 'Invalid session.' };
    }
    
    const user = session.user;
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    
    const attSheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const attData = attSheet.getDataRange().getValues();
    
    let todayAttendance = null;
    let presentDays = 0, lateDays = 0, absentDays = 0, halfDays = 0;
    
    for (let i = 1; i < attData.length; i++) {
      const rowEmpId = String(attData[i][1] || '').trim();
      const rowDate = String(attData[i][2] || '').trim();
      
      if (rowEmpId === user.id) {
        if (rowDate === today) {
          todayAttendance = {
            date: rowDate,
            punchIn: String(attData[i][3] || ''),
            punchOut: String(attData[i][4] || ''),
            status: String(attData[i][5] || ''),
            shift: String(attData[i][6] || ''),
            workHours: parseFloat(attData[i][7]) || 0
          };
        }
        
        if (rowDate.startsWith(currentMonth)) {
          const status = String(attData[i][5] || '').toLowerCase();
          if (status === 'present') presentDays++;
          else if (status === 'late') lateDays++;
          else if (status === 'half day') halfDays++;
          else if (status === 'absent') absentDays++;
        }
      }
    }
    
    const leaveBalance = getLeaveBalance(sessionId);
    const corrections = getPunchCorrections(sessionId);
    const leaves = getLeaveApplications(sessionId);
    const warnings = getWarnings(sessionId, user.id);
    
    let pendingCorrections = 0, pendingLeaves = 0, activeWarnings = 0;
    if (corrections.success) {
      pendingCorrections = corrections.corrections.filter(c => c.status.toLowerCase() === 'pending').length;
    }
    if (leaves.success) {
      pendingLeaves = leaves.leaves.filter(l => l.status.toLowerCase() === 'pending').length;
    }
    if (warnings.success) {
      activeWarnings = warnings.warnings.filter(w => w.status === 'Active').length;
    }
    
    return {
      success: true,
      dashboard: {
        todayAttendance: todayAttendance,
        monthlyStats: {
          present: presentDays,
          late: lateDays,
          absent: absentDays,
          halfDay: halfDays
        },
        leaveBalance: leaveBalance.success ? leaveBalance.balance : null,
        pendingRequests: {
          corrections: pendingCorrections,
          leaves: pendingLeaves
        },
        warnings: activeWarnings
      }
    };
  } catch (error) {
    return { success: false, message: 'Failed to get dashboard: ' + error.message };
  }
}

function getAdminDashboard(sessionId) {
  try {
    const session = validateSession(sessionId);
    if (!session.valid || session.user.role !== 'admin') {
      return { success: false, message: 'Unauthorized access.' };
    }
    
    console.log('getAdminDashboard called');
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Get all employees
    const empSheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const empData = empSheet.getDataRange().getValues();
    
    let totalEmployees = 0, activeEmployees = 0;
    
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][0]) {
        totalEmployees++;
        const status = String(empData[i][9] || '').toLowerCase().trim();
        if (status === 'active') {
          activeEmployees++;
        }
      }
    }
    
    // Get today's attendance
    const attSheet = getOrCreateSheet('Attendance');
    SpreadsheetApp.flush();
    const attData = attSheet.getDataRange().getValues();
    
    let todayPresent = 0, todayLate = 0, todayAbsent = 0;
    const todayAttendanceList = [];
    
    for (let i = 1; i < attData.length; i++) {
      const rowDate = String(attData[i][2] || '').trim();
      if (rowDate === today) {
        const status = String(attData[i][5] || '').toLowerCase().trim();
        if (status === 'present') todayPresent++;
        else if (status === 'late') todayLate++;
        else if (status === 'absent') todayAbsent++;
        
        todayAttendanceList.push({
          employeeId: String(attData[i][1] || ''),
          date: rowDate,
          punchIn: String(attData[i][3] || ''),
          punchOut: String(attData[i][4] || ''),
          status: String(attData[i][5] || ''),
          shift: String(attData[i][6] || '')
        });
      }
    }
    
    // Get pending leave requests
    const leaveSheet = getOrCreateSheet('Leaves');
    SpreadsheetApp.flush();
    const leaveData = leaveSheet.getDataRange().getValues();
    
    const pendingLeaves = [];
    for (let i = 1; i < leaveData.length; i++) {
      const rowStatus = String(leaveData[i][7] || '').trim().toLowerCase();
      if (rowStatus === 'pending') {
        pendingLeaves.push({
          leaveId: String(leaveData[i][0] || ''),
          employeeId: String(leaveData[i][1] || ''),
          leaveType: String(leaveData[i][2] || ''),
          startDate: String(leaveData[i][3] || ''),
          endDate: String(leaveData[i][4] || ''),
          days: parseInt(leaveData[i][5]) || 0,
          reason: String(leaveData[i][6] || ''),
          status: String(leaveData[i][7] || ''),
          appliedDate: String(leaveData[i][8] || '')
        });
      }
    }
    
    // Get pending punch corrections
    const corrSheet = getOrCreateSheet('PunchCorrections');
    SpreadsheetApp.flush();
    const corrData = corrSheet.getDataRange().getValues();
    
    const pendingCorrections = [];
    for (let i = 1; i < corrData.length; i++) {
      const rowStatus = String(corrData[i][8] || '').trim().toLowerCase();
      if (rowStatus === 'pending') {
        pendingCorrections.push({
          requestId: String(corrData[i][0] || ''),
          employeeId: String(corrData[i][1] || ''),
          date: String(corrData[i][2] || ''),
          currentPunchIn: String(corrData[i][3] || ''),
          currentPunchOut: String(corrData[i][4] || ''),
          requestedPunchIn: String(corrData[i][5] || ''),
          requestedPunchOut: String(corrData[i][6] || ''),
          reason: String(corrData[i][7] || ''),
          status: String(corrData[i][8] || ''),
          requestDate: String(corrData[i][9] || '')
        });
      }
    }
    
    return {
      success: true,
      dashboard: {
        totalEmployees: totalEmployees,
        activeEmployees: activeEmployees,
        todayStats: {
          present: todayPresent,
          late: todayLate,
          absent: Math.max(0, activeEmployees - todayPresent - todayLate),
          total: activeEmployees
        },
        todayAttendance: todayAttendanceList,
        pendingLeaves: pendingLeaves,
        pendingCorrections: pendingCorrections
      }
    };
  } catch (error) {
    console.error('getAdminDashboard error: ' + error.toString());
    return { success: false, message: 'Failed to get admin dashboard: ' + error.message };
  }
}

// ==========================================
// UTILITY FUNCTIONS
// ==========================================

function addAuditLog(userId, action, details) {
  try {
    const sheet = getOrCreateSheet('AuditLog');
    const logId = 'LOG' + String(Date.now()).slice(-6);
    sheet.appendRow([
      logId,
      new Date().toISOString(),
      userId,
      action,
      details,
      ''
    ]);
    SpreadsheetApp.flush();
  } catch (error) {
    console.error('Failed to add audit log: ' + error.toString());
  }
}

function getEmployeeNameById(employeeId) {
  try {
    const sheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(employeeId).trim()) {
        return String(data[i][1] || '');
      }
    }
    return employeeId;
  } catch (error) {
    return employeeId;
  }
}

// ==========================================
// SETUP AND INITIALIZATION
// ==========================================

function setupSpreadsheet() {
  try {
    const sheets = ['Employees', 'Attendance', 'Leaves', 'LeaveBalance', 'PunchCorrections', 
                   'Holidays', 'Roster', 'RosterConfig', 'Warnings', 'OTP', 'AuditLog', 'Settings'];
    
    sheets.forEach(sheetName => {
      getOrCreateSheet(sheetName);
    });
    
    const empSheet = getOrCreateSheet('Employees');
    SpreadsheetApp.flush();
    const empData = empSheet.getDataRange().getValues();
    
    if (empData.length <= 1) {
      empSheet.appendRow([
        'EMP000001',
        'System Admin',
        'admin@company.com',
        '9876543210',
        'Administration',
        'HR Manager',
        new Date().toISOString().split('T')[0],
        'admin123',
        'admin',
        'active',
        'A',
        new Date().toISOString(),
        'Sunday',
        '',
        '',
        ''
      ]);
      SpreadsheetApp.flush();
    }
    
    return 'Setup completed successfully!';
  } catch (error) {
    return 'Setup failed: ' + error.message;
  }
}

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  return 'Spreadsheet ID set to: ' + id;
}

// ==========================================
// WEB APP ROUTES
// ==========================================

function doGet(e) {
  const page = e.parameter.page || 'index';
  
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Admin Dashboard - ' + COMPANY_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(COMPANY_NAME + ' - Employee Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
