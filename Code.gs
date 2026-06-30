/*
  الزواج المغربي - Apps Script Backend
  -------------------------------------
  هيكل الجداول المطلوبة في Google Sheet (الأسماء مطابقة):
  1. Profiles
  2. Interests
  3. AdminChatRequests
  4. AdminChatMessages
  5. Messages
  6. RegistrationRequests

  طريقة النشر:
  - Deploy -> New deployment -> Web app
  - Execute as: Me
  - Who has access: Anyone
  - انسخ الرابط وألصقه في HTML في المتغير CONFIG.API_URL
*/

const CONFIG = {
  SHEET_ID: '1JZUCmGoz7xPf2UEU93z18VpFKb361EzJTHUAVIiQIoc', // ← عدّل هذا
  ADMIN_EMAIL: 'hassanbts20@gmail.com',          // ← عدّل هذا لبريد الموجه
  AUTO_APPROVE_REGISTRATION: true,           // false = يتطلب موافقة الموجه
  DEFAULT_REMAINING_MESSAGES: 5
};

/* ========== دوال أساسية ========== */
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    if (!action) {
      return jsonResponse({ status: 'error', msg: 'Missing action parameter' });
    }

    const router = {
      'login': handleLogin,
      'register': handleRegister,
      'update': handleUpdate,
      'delete': handleDelete,
      'list': handleList,
      'sendInterest': handleSendInterest,
      'respondInterest': handleRespondInterest,
      'getMyRequests': handleGetMyRequests,
      'requestAdminChat': handleRequestAdminChat,
      'getAdminChatStatus': handleGetAdminChatStatus,
      'getAdminChatMessages': handleGetAdminChatMessages,
      'sendAdminChatMessage': handleSendAdminChatMessage,
      'getConversations': handleGetConversations,
      'getMessages': handleGetMessages,
      'sendMessage': handleSendMessage,
      'adminDashboard': handleAdminDashboard,
      'adminReviewRegistrationRequest': handleAdminReviewRegistrationRequest,
      'adminDeleteUser': handleAdminDeleteUser,
      'adminGetChatRequests': handleAdminGetChatRequests,
      'adminOpenChatRequest': handleAdminOpenChatRequest
    };

    if (router[action]) {
      return router[action](e);
    }
    return jsonResponse({ status: 'error', msg: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', msg: err.message });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ========== أدوات الجداول ========== */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // نضع عناوين الأعمدة افتراضية
    const headers = {
      'Profiles': ['email', 'activationCode', 'name', 'age', 'gender', 'status', 'location', 'birthCity', 'height', 'specs', 'imageUrl', 'isAdmin', 'ratingAverage', 'ratingCount', 'createdAt'],
      'Interests': ['requestId', 'senderEmail', 'receiverEmail', 'status', 'createdAt', 'remainingAttempts'],
      'AdminChatRequests': ['requestId', 'userEmail', 'requestStatus', 'remainingUserMessages', 'createdAt'],
      'AdminChatMessages': ['requestId', 'senderType', 'senderEmail', 'messageText', 'createdAt'],
      'Messages': ['messageId', 'senderEmail', 'receiverEmail', 'messageText', 'createdAt'],
      'RegistrationRequests': ['requestId', 'name', 'email', 'age', 'gender', 'status', 'location', 'birthCity', 'height', 'specs', 'imageUrl', 'activationCode', 'decision', 'createdAt']
    };
    if (headers[name]) sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
  }
  return sheet;
}

function getRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] || '');
    return obj;
  });
}

function findRowIndex(sheet, columnName, value) {
  const data = sheet.getDataRange().getValues();
  if (!data.length) return -1;
  const headers = data[0];
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]).toLowerCase().trim() === String(value).toLowerCase().trim()) {
      return i + 1; // 1-based row index
    }
  }
  return -1;
}

function generateId(prefix) {
  return prefix + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);
}

function nowStr() {
  return new Date().toISOString();
}

function hashEmail(email) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'MATCHNEST::' + email.toLowerCase().trim(), Utilities.Charset.UTF_8);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').slice(0, 8).toUpperCase();
}

function verifyAuth(e) {
  const email = e.parameter.email || '';
  const code = e.parameter.code || e.parameter.activationCode || '';
  if (!email || !code) throw new Error('Email and code are required');
  const expected = hashEmail(email);
  if (code !== expected) throw new Error('Invalid activation code');
  return email;
}

function getProfileByEmail(email) {
  const sheet = getSheet('Profiles');
  const rowIndex = findRowIndex(sheet, 'email', email);
  if (rowIndex === -1) return null;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data[rowIndex - 1];
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i] || '');
  obj.isAdmin = obj.isAdmin === true || obj.isAdmin === 'true' || String(obj.email).toLowerCase() === String(CONFIG.ADMIN_EMAIL).toLowerCase();
  return obj;
}

function profileToFrontend(p) {
  if (!p) return null;
  return {
    email: p.email || '',
    name: p.name || '',
    age: p.age || '',
    gender: p.gender || '',
    status: p.status || '',
    location: p.location || '',
    birthCity: p.birthCity || '',
    height: p.height || '',
    specs: p.specs || '',
    imageUrl: p.imageUrl || '',
    isAdmin: p.isAdmin === true || p.isAdmin === 'true' || String(p.email).toLowerCase() === String(CONFIG.ADMIN_EMAIL).toLowerCase(),
    ratingAverage: Number(p.ratingAverage || 0),
    ratingCount: Number(p.ratingCount || 0)
  };
}

/* ========== المصادحة والتسجيل ========== */
function handleLogin(e) {
  const email = e.parameter.email || '';
  const code = e.parameter.code || e.parameter.activationCode || '';
  if (!email || !code) {
    return jsonResponse({ status: 'error', msg: 'Email and code required' });
  }
  const expected = hashEmail(email);
  if (code !== expected) {
    return jsonResponse({ status: 'error', msg: 'Invalid activation code' });
  }
  const profile = getProfileByEmail(email);
  if (!profile) {
    return jsonResponse({ status: 'error', msg: 'Profile not found' });
  }
  return jsonResponse({ status: 'success', profile: profileToFrontend(profile) });
}

function handleRegister(e) {
  const email = e.parameter.email || '';
  const code = e.parameter.activationCode || '';
  if (!email || !code) {
    return jsonResponse({ status: 'error', msg: 'Email and activationCode required' });
  }
  const expected = hashEmail(email);
  if (code !== expected) {
    return jsonResponse({ status: 'error', msg: 'Invalid activation code' });
  }

  const profiles = getSheet('Profiles');
  if (findRowIndex(profiles, 'email', email) !== -1) {
    return jsonResponse({ status: 'error', msg: 'This email is already registered' });
  }

  const newProfile = {
    email: email,
    activationCode: code,
    name: e.parameter.name || '',
    age: e.parameter.age || '',
    gender: e.parameter.gender || '',
    status: e.parameter.status || '',
    location: e.parameter.location || '',
    birthCity: e.parameter.birthCity || '',
    height: e.parameter.height || '',
    specs: e.parameter.specs || '',
    imageUrl: e.parameter.imageUrl || '',
    isAdmin: false,
    ratingAverage: 0,
    ratingCount: 0,
    createdAt: nowStr()
  };

  if (!CONFIG.AUTO_APPROVE_REGISTRATION) {
    const reqSheet = getSheet('RegistrationRequests');
    reqSheet.appendRow([
      generateId('REG'),
      newProfile.name,
      newProfile.email,
      newProfile.age,
      newProfile.gender,
      newProfile.status,
      newProfile.location,
      newProfile.birthCity,
      newProfile.height,
      newProfile.specs,
      newProfile.imageUrl,
      newProfile.activationCode,
      'pending',
      newProfile.createdAt
    ]);
    return jsonResponse({ status: 'pending_approval', msg: 'بانتظار موافقة الموجه' });
  }

  profiles.appendRow([
    newProfile.email,
    newProfile.activationCode,
    newProfile.name,
    newProfile.age,
    newProfile.gender,
    newProfile.status,
    newProfile.location,
    newProfile.birthCity,
    newProfile.height,
    newProfile.specs,
    newProfile.imageUrl,
    false,
    0,
    0,
    newProfile.createdAt
  ]);

  return jsonResponse({ status: 'success', profile: profileToFrontend(newProfile) });
}

/* ========== التحديث والحذف ========== */
function handleUpdate(e) {
  const email = verifyAuth(e);
  const sheet = getSheet('Profiles');
  const rowIndex = findRowIndex(sheet, 'email', email);
  if (rowIndex === -1) {
    return jsonResponse({ status: 'error', msg: 'Profile not found' });
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data[rowIndex - 1];
  const current = {};
  headers.forEach((h, i) => current[h] = row[i] || '');

  const updated = {
    name: e.parameter.name || current.name,
    age: e.parameter.age || current.age,
    gender: e.parameter.gender || current.gender,
    status: e.parameter.status || current.status,
    location: e.parameter.location || current.location,
    birthCity: e.parameter.birthCity || current.birthCity,
    height: e.parameter.height || current.height,
    specs: e.parameter.specs || current.specs,
    imageUrl: e.parameter.imageUrl || current.imageUrl
  };

  const setVal = (col, val) => {
    const idx = headers.indexOf(col);
    if (idx !== -1) sheet.getRange(rowIndex, idx + 1).setValue(val);
  };

  setVal('name', updated.name);
  setVal('age', updated.age);
  setVal('gender', updated.gender);
  setVal('status', updated.status);
  setVal('location', updated.location);
  setVal('birthCity', updated.birthCity);
  setVal('height', updated.height);
  setVal('specs', updated.specs);
  setVal('imageUrl', updated.imageUrl);

  return jsonResponse({ status: 'success', msg: 'Profile updated' });
}

function handleDelete(e) {
  const email = verifyAuth(e);
  const sheet = getSheet('Profiles');
  const rowIndex = findRowIndex(sheet, 'email', email);
  if (rowIndex === -1) {
    return jsonResponse({ status: 'error', msg: 'Profile not found' });
  }
  sheet.deleteRow(rowIndex);
  return jsonResponse({ status: 'success', msg: 'Account deleted' });
}

/* ========== قائمة الأعضاء ========== */
function handleList(e) {
  // لا يحتاج مصادقة - لعرض البطاقات
  const rows = getRows(getSheet('Profiles'));
  const profiles = rows.map(profileToFrontend).filter(p => p);

  // التأكد من وجود الموجه في القائمة
  const adminExists = profiles.some(p => p.isAdmin);
  if (!adminExists) {
    profiles.push({
      email: CONFIG.ADMIN_EMAIL,
      name: 'الموجه',
      age: '',
      gender: '',
      status: '',
      location: 'المغرب',
      birthCity: '',
      height: '',
      specs: 'الموجه متاح للمساعدة والإرشاد.',
      imageUrl: 'https://ui-avatars.com/api/?name=الموجه&background=E50914&color=fff&size=256',
      isAdmin: true,
      ratingAverage: 5,
      ratingCount: 0
    });
  }

  return jsonResponse({ status: 'success', profiles: profiles });
}

/* ========== طلبات الاهتمام ========== */
function handleSendInterest(e) {
  const senderEmail = e.parameter.senderEmail || '';
  const receiverEmail = e.parameter.receiverEmail || '';
  const code = e.parameter.activationCode || '';
  if (!senderEmail || !receiverEmail || !code) {
    return jsonResponse({ status: 'error', msg: 'Missing data' });
  }
  if (code !== hashEmail(senderEmail)) {
    return jsonResponse({ status: 'error', msg: 'Invalid code' });
  }
  if (senderEmail === receiverEmail) {
    return jsonResponse({ status: 'error', msg: 'Cannot send interest to yourself' });
  }

  const sheet = getSheet('Interests');
  const rows = getRows(sheet);
  const existing = rows.find(r => r.senderEmail === senderEmail && r.receiverEmail === receiverEmail);
  if (existing) {
    const remaining = Number(existing.remainingAttempts || 0);
    if (remaining <= 0) {
      return jsonResponse({ status: 'error', msg: 'No remaining attempts' });
    }
    const rowIndex = findRowIndex(sheet, 'requestId', existing.requestId);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 5).setValue('pending'); // status
      sheet.getRange(rowIndex, 6).setValue(remaining - 1); // remainingAttempts
    }
    return jsonResponse({ status: 'success', msg: 'Interest re-sent' });
  }

  sheet.appendRow([generateId('REQ'), senderEmail, receiverEmail, 'pending', nowStr(), 2]);
  return jsonResponse({ status: 'success', msg: 'Interest sent' });
}

function handleRespondInterest(e) {
  const email = verifyAuth(e);
  const requestId = e.parameter.requestId || '';
  const responseStatus = e.parameter.responseStatus || '';
  if (!requestId || !['accepted', 'rejected'].includes(responseStatus)) {
    return jsonResponse({ status: 'error', msg: 'Invalid response' });
  }
  const sheet = getSheet('Interests');
  const rowIndex = findRowIndex(sheet, 'requestId', requestId);
  if (rowIndex === -1) {
    return jsonResponse({ status: 'error', msg: 'Request not found' });
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data[rowIndex - 1];
  const receiver = row[headers.indexOf('receiverEmail')];
  if (receiver !== email) {
    return jsonResponse({ status: 'error', msg: 'Not authorized' });
  }
  sheet.getRange(rowIndex, headers.indexOf('status') + 1).setValue(responseStatus);
  return jsonResponse({ status: 'success', msg: 'Response recorded' });
}

function handleGetMyRequests(e) {
  const email = verifyAuth(e);
  const rows = getRows(getSheet('Interests'));
  const incoming = rows.filter(r => String(r.receiverEmail).toLowerCase() === email.toLowerCase());
  const outgoing = rows.filter(r => String(r.senderEmail).toLowerCase() === email.toLowerCase());

  const enrich = r => {
    const isIncoming = r.receiverEmail === email;
    const otherEmail = isIncoming ? r.senderEmail : r.receiverEmail;
    const otherProfile = getProfileByEmail(otherEmail);
    return {
      requestId: r.requestId,
      senderEmail: r.senderEmail,
      receiverEmail: r.receiverEmail,
      status: r.status,
      remainingAttempts: Number(r.remainingAttempts || 0),
      senderProfile: isIncoming ? profileToFrontend(otherProfile) : undefined,
      receiverProfile: !isIncoming ? profileToFrontend(otherProfile) : undefined
    };
  };

  return jsonResponse({
    status: 'success',
    incoming: incoming.map(enrich),
    outgoing: outgoing.map(enrich)
  });
}

/* ========== محادثة الموجه ========== */
function getOrCreateAdminRequest(email) {
  const sheet = getSheet('AdminChatRequests');
  const rows = getRows(sheet);
  let req = rows.find(r => String(r.userEmail).toLowerCase() === email.toLowerCase());
  if (!req) {
    const requestId = generateId('ADREQ');
    sheet.appendRow([requestId, email, 'pending', CONFIG.DEFAULT_REMAINING_MESSAGES, nowStr()]);
    req = { requestId: requestId, userEmail: email, requestStatus: 'pending', remainingUserMessages: CONFIG.DEFAULT_REMAINING_MESSAGES };
  }
  return req;
}

function handleRequestAdminChat(e) {
  const email = verifyAuth(e);
  const req = getOrCreateAdminRequest(email);
  return jsonResponse({ status: 'success', requestId: req.requestId, requestStatus: req.requestStatus });
}

function handleGetAdminChatStatus(e) {
  const email = verifyAuth(e);
  const req = getOrCreateAdminRequest(email);
  return jsonResponse({
    status: 'success',
    requestId: req.requestId,
    requestStatus: req.requestStatus,
    remainingUserMessages: Number(req.remainingUserMessages || 0)
  });
}

function handleGetAdminChatMessages(e) {
  verifyAuth(e);
  const requestId = e.parameter.requestId || '';
  if (!requestId) return jsonResponse({ status: 'error', msg: 'requestId required' });
  const rows = getRows(getSheet('AdminChatMessages')).filter(r => r.requestId === requestId);
  return jsonResponse({ status: 'success', messages: rows });
}

function handleSendAdminChatMessage(e) {
  const email = verifyAuth(e);
  const requestId = e.parameter.requestId || '';
  const messageText = e.parameter.messageText || '';
  if (!requestId || !messageText) {
    return jsonResponse({ status: 'error', msg: 'requestId and messageText required' });
  }

  const reqSheet = getSheet('AdminChatRequests');
  const reqRow = findRowIndex(reqSheet, 'requestId', requestId);
  if (reqRow === -1) return jsonResponse({ status: 'error', msg: 'Request not found' });

  const data = reqSheet.getDataRange().getValues();
  const headers = data[0];
  const row = data[reqRow - 1];
  const userEmail = row[headers.indexOf('userEmail')];
  const status = row[headers.indexOf('requestStatus')];
  let remaining = Number(row[headers.indexOf('remainingUserMessages')] || 0);

  const isUser = String(email).toLowerCase() === String(userEmail).toLowerCase();
  if (isUser && status !== 'opened') {
    return jsonResponse({ status: 'error', msg: 'Chat not opened yet' });
  }
  if (isUser && remaining <= 0) {
    return jsonResponse({ status: 'error', msg: 'No remaining messages. Wait for admin reply.' });
  }

  const msgSheet = getSheet('AdminChatMessages');
  msgSheet.appendRow([requestId, isUser ? 'user' : 'admin', email, messageText, nowStr()]);

  if (isUser) {
    remaining = Math.max(0, remaining - 1);
    reqSheet.getRange(reqRow, headers.indexOf('remainingUserMessages') + 1).setValue(remaining);
  } else {
    // رد الموجه يعيد تفعيل 5 رسائل
    reqSheet.getRange(reqRow, headers.indexOf('remainingUserMessages') + 1).setValue(CONFIG.DEFAULT_REMAINING_MESSAGES);
  }

  return jsonResponse({ status: 'success', msg: 'Message sent', remainingUserMessages: remaining });
}

/* ========== المحادثات الخاصة ========== */
function handleSendMessage(e) {
  const senderEmail = e.parameter.senderEmail || '';
  const receiverEmail = e.parameter.receiverEmail || '';
  const code = e.parameter.code || e.parameter.activationCode || '';
  const messageText = e.parameter.messageText || '';
  if (!senderEmail || !receiverEmail || !code || !messageText) {
    return jsonResponse({ status: 'error', msg: 'Missing data' });
  }
  if (code !== hashEmail(senderEmail)) {
    return jsonResponse({ status: 'error', msg: 'Invalid code' });
  }
  const sheet = getSheet('Messages');
  sheet.appendRow([generateId('MSG'), senderEmail, receiverEmail, messageText, nowStr()]);
  return jsonResponse({ status: 'success', msg: 'Message sent' });
}

function handleGetConversations(e) {
  const email = verifyAuth(e);
  const rows = getRows(getSheet('Messages'));
  const map = {};
  rows.forEach(r => {
    const isSender = String(r.senderEmail).toLowerCase() === email.toLowerCase();
    const other = isSender ? r.receiverEmail : r.senderEmail;
    if (!map[other]) map[other] = { otherEmail: other, lastMessage: r.messageText, lastAt: r.createdAt };
    else if (r.createdAt > map[other].lastAt) {
      map[other].lastMessage = r.messageText;
      map[other].lastAt = r.createdAt;
    }
  });
  const conversations = Object.values(map).map(c => {
    const otherProfile = getProfileByEmail(c.otherEmail);
    return { otherEmail: c.otherEmail, otherProfile: profileToFrontend(otherProfile), lastMessage: c.lastMessage };
  });
  return jsonResponse({ status: 'success', conversations: conversations });
}

function handleGetMessages(e) {
  const email = verifyAuth(e);
  const otherEmail = e.parameter.otherEmail || '';
  if (!otherEmail) return jsonResponse({ status: 'error', msg: 'otherEmail required' });
  const rows = getRows(getSheet('Messages')).filter(r =>
    (String(r.senderEmail).toLowerCase() === email.toLowerCase() && String(r.receiverEmail).toLowerCase() === otherEmail.toLowerCase()) ||
    (String(r.senderEmail).toLowerCase() === otherEmail.toLowerCase() && String(r.receiverEmail).toLowerCase() === email.toLowerCase())
  );
  return jsonResponse({ status: 'success', messages: rows });
}

/* ========== لوحة الموجه ========== */
function requireAdmin(e) {
  const email = verifyAuth(e);
  const profile = getProfileByEmail(email);
  if (!profile || !profile.isAdmin) {
    throw new Error('Admin only');
  }
  return email;
}

function handleAdminDashboard(e) {
  requireAdmin(e);
  const profiles = getRows(getSheet('Profiles'));
  const requests = getRows(getSheet('Interests'));
  const adminChatRequests = getRows(getSheet('AdminChatRequests'));
  const registrationRequests = getRows(getSheet('RegistrationRequests'));
  const messages = getRows(getSheet('Messages'));
  const adminMessages = getRows(getSheet('AdminChatMessages'));

  return jsonResponse({
    status: 'success',
    totals: {
      profiles: profiles.length,
      requests: requests.length,
      adminChatRequests: adminChatRequests.length,
      pendingRegistrationRequests: registrationRequests.filter(r => r.decision === 'pending' || r.decision === '').length,
      memberMessages: messages.length,
      adminMessages: adminMessages.length
    },
    registrationRequests: registrationRequests.filter(r => r.decision === 'pending' || r.decision === '')
  });
}

function handleAdminReviewRegistrationRequest(e) {
  requireAdmin(e);
  const requestId = e.parameter.requestId || '';
  const decision = e.parameter.decision || '';
  if (!requestId || !['approved', 'rejected'].includes(decision)) {
    return jsonResponse({ status: 'error', msg: 'Invalid decision' });
  }
  const sheet = getSheet('RegistrationRequests');
  const rowIndex = findRowIndex(sheet, 'requestId', requestId);
  if (rowIndex === -1) return jsonResponse({ status: 'error', msg: 'Request not found' });
  sheet.getRange(rowIndex, 13).setValue(decision); // decision column

  if (decision === 'approved') {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const row = data[rowIndex - 1];
    const p = {};
    headers.forEach((h, i) => p[h] = row[i] || '');
    const profiles = getSheet('Profiles');
    if (findRowIndex(profiles, 'email', p.email) === -1) {
      profiles.appendRow([p.email, p.activationCode, p.name, p.age, p.gender, p.status, p.location, p.birthCity, p.height, p.specs, p.imageUrl, false, 0, 0, nowStr()]);
    }
  }
  return jsonResponse({ status: 'success', msg: 'Decision recorded' });
}

function handleAdminDeleteUser(e) {
  requireAdmin(e);
  const targetEmail = e.parameter.targetEmail || '';
  if (!targetEmail) return jsonResponse({ status: 'error', msg: 'targetEmail required' });
  const sheet = getSheet('Profiles');
  const rowIndex = findRowIndex(sheet, 'email', targetEmail);
  if (rowIndex === -1) return jsonResponse({ status: 'error', msg: 'User not found' });
  sheet.deleteRow(rowIndex);
  return jsonResponse({ status: 'success', msg: 'User deleted' });
}

function handleAdminGetChatRequests(e) {
  requireAdmin(e);
  const rows = getRows(getSheet('AdminChatRequests'));
  const requests = rows.map(r => {
    const profile = getProfileByEmail(r.userEmail);
    return {
      requestId: r.requestId,
      userEmail: r.userEmail,
      userProfile: profileToFrontend(profile),
      requestStatus: r.requestStatus,
      remainingUserMessages: Number(r.remainingUserMessages || 0)
    };
  });
  return jsonResponse({ status: 'success', requests: requests });
}

function handleAdminOpenChatRequest(e) {
  requireAdmin(e);
  const requestId = e.parameter.requestId || '';
  if (!requestId) return jsonResponse({ status: 'error', msg: 'requestId required' });
  const sheet = getSheet('AdminChatRequests');
  const rowIndex = findRowIndex(sheet, 'requestId', requestId);
  if (rowIndex === -1) return jsonResponse({ status: 'error', msg: 'Request not found' });
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  sheet.getRange(rowIndex, headers.indexOf('requestStatus') + 1).setValue('opened');
  sheet.getRange(rowIndex, headers.indexOf('remainingUserMessages') + 1).setValue(CONFIG.DEFAULT_REMAINING_MESSAGES);
  return jsonResponse({ status: 'success', msg: 'Chat opened' });
}
