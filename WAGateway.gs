/**
 * Layanan WhatsApp Gateway Terpusat
 * Mendukung Multi-Provider: WAplus, Fonnte, StarSender
 */

function sendWhatsApp(target, message, token, sender, provider = 'waplus') {
  if (!target || !message || !token) return false;

  // Normalisasi nomor target (hilangkan +, ubah 08 jadi 62)
  target = String(target).replace(/\D/g, '');
  if (target.startsWith('0')) {
    target = '62' + target.substring(1);
  }

  try {
    switch(provider.toLowerCase()) {
      case 'waplus':
        // WAplus butuh parameter sender
        if(!sender) return false;
        return sendViaWaplus(target, message, token, sender);
      
      case 'fonnte':
      // Fonnte mengenali sender dari token
      return sendViaFonnte(target, message, token); 
    
    case 'starsender':
      // StarSender mengenali sender dari Device API key (token)
      return sendViaStarSender(target, message, token);
    
    case 'xsender':
      // Xsender butuh parameter sender
      if(!sender) return false;
      return sendViaXsender(target, message, token, sender);
    
    default:
      console.error("Provider WA tidak dikenal: " + provider);
      return false;
    }
  } catch (err) {
    console.error("Error mengirim WA via " + provider + ": " + err);
    return false;
  }
}

// -------------------------------------------------------------
// IMPLEMENTASI API WAPLUS
// -------------------------------------------------------------
function sendViaWaplus(target, message, token, sender) {
  const payload = {
    "api_key": token,
    "sender": sender,
    "number": target,
    "message": message
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://app.waplus.id/send-message', options);
  const json = JSON.parse(response.getContentText());
  
  return json.status === true;
}

// -------------------------------------------------------------
// IMPLEMENTASI API FONNTE
// -------------------------------------------------------------
function sendViaFonnte(target, message, token) {
  const options = {
    method: 'post',
    headers: { 'Authorization': token },
    payload: { 'target': target, 'message': message },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
  const json = JSON.parse(response.getContentText());
  
  return json.status === true;
}

// -------------------------------------------------------------
// IMPLEMENTASI API STARSENDER
// -------------------------------------------------------------
function sendViaStarSender(target, message, token) {
  // FIX BUG Content-Type Apps Script + Auto Trim Token
  const fetchUrl = "https://api.starsender.online/api/send";
  const options = {
    "method": "POST",
    "headers": { "Authorization": String(token).trim() },
    "contentType": "application/json", // Wajib di luar header
    "muteHttpExceptions": true,
    "payload": JSON.stringify({ 
      "messageType": "text", 
      "to": target, 
      "body": message 
    })
  };
  
  const response = UrlFetchApp.fetch(fetchUrl, options);
  const json = JSON.parse(response.getContentText());
  
  // Asumsi sukses jika tidak ada error (tergantung respon spesifik StarSender, biasanya mengembalikan status/success true)
  return !!json; 
}

// -------------------------------------------------------------
// IMPLEMENTASI API XSENDER
// -------------------------------------------------------------
function sendViaXsender(target, message, token, sender) {
  const fetchUrl = "https://xsender.id/send-message";
  const options = {
    "method": "POST",
    "contentType": "application/json",
    "muteHttpExceptions": true,
    "payload": JSON.stringify({ 
      "api_key": String(token).trim(), 
      "sender": sender, 
      "number": target, 
      "message": message 
    })
  };
  
  const response = UrlFetchApp.fetch(fetchUrl, options);
  const json = JSON.parse(response.getContentText());
  
  return json.status === true;
}
