function doGet(e) {
  // Return a test response for GET requests
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: 'Google Apps Script is working',
      timestamp: new Date(new Date().getTime() + 8 * 60 * 60 * 1000).toISOString().replace('T', ' ').replace('Z', ' UTC+8'),
      method: 'GET'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(request){
  return handleRequest(request);
}

//  Enter sheet name where data is to be written below
var SHEET_NAME = "RSVP_responses";

// REPLACE THIS WITH YOUR OWN SHEET ID
var SHEET_ID = "11Sp1j-FPY18T-WCTdqSCvMhP9ANqaCX12QXSsFlUhh8"; 
var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

function handleRequest(request) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.

  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SHEET_ID);
    var sheet = doc.getSheetByName(SHEET_NAME);

    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = request.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1; // get next row
    var row = [];
    
    // loop through the header columns
    for (i in headers){
      if (headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
        var utc8Time = new Date(new Date().getTime() + 8 * 60 * 60 * 1000);
        row.push(utc8Time.toISOString().replace('T', ' ').replace('Z', ' UTC+8'));
      } else if(headers[i] == "sn") {
        row.push(sheet.getLastRow());
      } else { // else use header name to get data
        row.push(request.parameter[headers[i]] || ''); // Add fallback for missing parameters
      }
    }
    
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

    // Format the new row for better readability
    var range = sheet.getRange(nextRow, 1, 1, row.length);
    range.setBorder(true, true, true, true, true, true);
    
    // Auto-resize columns to fit content
    sheet.autoResizeColumns(1, sheet.getLastColumn());

    // send email notification
    sendEmail(request.parameter);

    // Log successful submission
    console.log('Successfully processed RSVP submission for: ' + request.parameter.name);

    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "message": "RSVP submitted successfully"}))
          .setMimeType(ContentService.MimeType.JSON);
          
  } catch(e){
    // Log error for debugging
    console.error('Error processing RSVP submission:', e.toString());
    
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e.toString()}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
    
    // Initialize the sheet with headers and formatting
    var initResult = initializeSheet();
    
    if (initResult.success) {
      console.log('✅ Setup completed successfully');
    } else {
      console.error('❌ Setup failed:', initResult.error);
    }
    
    return initResult;
}

function createHeaders(sheet) {
  var headers = [
    'Timestamp',
    'name',
    'relation',
    'attendance',
    'guests',
    'dietary',
    'vegetarian-meals',
    'children-seats',
    'children-seats-count',
    'invitation',
    'address',
    'message',
    'email'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#8B4513');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setBorder(true, true, true, true, true, true);
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

function initializeSheet() {
  try {
    console.log('Initializing wedding RSVP sheet...');
    
    // Open the spreadsheet
    var doc = SpreadsheetApp.openById(SHEET_ID);
    
    // Check if the sheet exists, if not create it
    var sheet = doc.getSheetByName(SHEET_NAME);
    if (!sheet) {
      console.log('Sheet "' + SHEET_NAME + '" not found. Creating new sheet...');
      sheet = doc.insertSheet(SHEET_NAME);
      console.log('✅ Sheet created successfully');
    } else {
      console.log('✅ Sheet "' + SHEET_NAME + '" found');
    }
    
    // Check if headers already exist
    var lastRow = sheet.getLastRow();
    var existingHeaders = [];
    
    if (lastRow > 0) {
      existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    // If no headers exist or first cell is empty, create headers
    if (lastRow === 0 || !existingHeaders[0] || existingHeaders[0] === '') {
      console.log('No headers found. Creating headers...');
      createHeaders(sheet);
      console.log('✅ Headers created successfully');
    } else {
      console.log('✅ Headers already exist');
      
      // Verify that all expected headers are present
      var expectedHeaders = [
        'Timestamp',
        'name',
        'relation',
        'attendance',
        'guests',
        'dietary',
        'vegetarian-meals',
        'children-seats',
        'children-seats-count',
        'invitation',
        'address',
        'message',
        'email'
      ];
      
      var missingHeaders = [];
      expectedHeaders.forEach(function(header) {
        if (existingHeaders.indexOf(header) === -1) {
          missingHeaders.push(header);
        }
      });
      
      if (missingHeaders.length > 0) {
        console.log('⚠️ Missing headers detected: ' + missingHeaders.join(', '));
        console.log('Recreating headers to ensure all fields are present...');
        
        // Clear existing headers and recreate
        sheet.getRange(1, 1, 1, sheet.getLastColumn()).clearContent();
        createHeaders(sheet);
        console.log('✅ Headers updated successfully');
      }
    }
    
    // Set up some additional formatting for better usability
    
    // Set column widths for better readability
    sheet.setColumnWidth(1, 150); // Timestamp
    sheet.setColumnWidth(2, 200); // Name
    sheet.setColumnWidth(3, 150); // Relation
    sheet.setColumnWidth(4, 120); // Attendance
    sheet.setColumnWidth(5, 100); // Guests
    sheet.setColumnWidth(6, 150); // Dietary
    sheet.setColumnWidth(7, 150); // Vegetarian-meals
    sheet.setColumnWidth(8, 150); // Children-seats
    sheet.setColumnWidth(9, 150); // Children-seats-count
    sheet.setColumnWidth(10, 150); // Invitation
    sheet.setColumnWidth(11, 300); // Address
    sheet.setColumnWidth(12, 300); // Message
    sheet.setColumnWidth(13, 200); // Email
    
    // Add data validation for attendance column (column D)
    if (sheet.getLastColumn() >= 4) {
      var attendanceColumn = sheet.getRange('D:D');
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['yes', 'no'], true)
        .setAllowInvalid(false)
        .setHelpText('Please select either "yes" or "no"')
        .build();
      attendanceColumn.setDataValidation(rule);
    }
    
    // Add data validation for guests column (column E)
    if (sheet.getLastColumn() >= 5) {
      var guestsColumn = sheet.getRange('E:E');
      var guestsRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['1', '2', '3', '4', '5'], true)
        .setAllowInvalid(true)
        .setHelpText('Select number of guests (1-5)')
        .build();
      guestsColumn.setDataValidation(guestsRule);
    }
    
    // Add data validation for relation column (column C)
    if (sheet.getLastColumn() >= 3) {
      var relationColumn = sheet.getRange('C:C');
      var relationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['friends-groom', 'friends-bride', 'other'], true)
        .setAllowInvalid(true)
        .setHelpText('Select relation type')
        .build();
      relationColumn.setDataValidation(relationRule);
    }
    
    // Add data validation for dietary column (column F)
    if (sheet.getLastColumn() >= 6) {
      var dietaryColumn = sheet.getRange('F:F');
      var dietaryRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['', '1', '2'], true)
        .setAllowInvalid(true)
        .setHelpText('Select dietary preference')
        .build();
      dietaryColumn.setDataValidation(dietaryRule);
    }
    
    // Add data validation for vegetarian-meals column (column G)
    if (sheet.getLastColumn() >= 7) {
      var vegetarianMealsColumn = sheet.getRange('G:G');
      var vegetarianMealsRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['', '1', '2', '3', '4', '5'], true)
        .setAllowInvalid(true)
        .setHelpText('Select number of vegetarian meals needed')
        .build();
      vegetarianMealsColumn.setDataValidation(vegetarianMealsRule);
    }
    
    // Add data validation for children-seats column (column H)
    if (sheet.getLastColumn() >= 8) {
      var childrenSeatsColumn = sheet.getRange('H:H');
      var childrenSeatsRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['yes', 'no'], true)
        .setAllowInvalid(true)
        .setHelpText('Select if children seats are needed')
        .build();
      childrenSeatsColumn.setDataValidation(childrenSeatsRule);
    }
    
    // Add data validation for children-seats-count column (column I)
    if (sheet.getLastColumn() >= 9) {
      var childrenSeatsCountColumn = sheet.getRange('I:I');
      var childrenSeatsCountRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['', '1', '2', '3', '4', '5'], true)
        .setAllowInvalid(true)
        .setHelpText('Select number of children seats needed')
        .build();
      childrenSeatsCountColumn.setDataValidation(childrenSeatsCountRule);
    }
    
    // Add data validation for invitation column (column J)
    if (sheet.getLastColumn() >= 10) {
      var invitationColumn = sheet.getRange('J:J');
      var invitationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['yes', 'no'], true)
        .setAllowInvalid(true)
        .setHelpText('Select if paper invitation is needed')
        .build();
      invitationColumn.setDataValidation(invitationRule);
    }
    
    console.log('✅ Sheet initialization completed successfully!');
    console.log('📊 Sheet Details:');
    console.log('   - Sheet Name: ' + SHEET_NAME);
    console.log('   - Sheet ID: ' + SHEET_ID);
    console.log('   - Total Columns: ' + sheet.getLastColumn());
    console.log('   - Total Rows: ' + sheet.getLastRow());
    
    return {
      success: true,
      message: 'Sheet initialized successfully',
      sheetName: SHEET_NAME,
      columns: sheet.getLastColumn(),
      rows: sheet.getLastRow()
    };
    
  } catch(e) {
    console.error('❌ Error initializing sheet:', e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

function sendEmail(data) {
  try {
    var johnnyEmail = 'johnny114chiu@gmail.com';
    
    // Send confirmation email to the guest if they provided an email
    if (data.email && data.email.trim() !== '') {
      var guestSubject = 'RSVP 確認通知 - Johnny & Josephine 婚禮';
      var guestBody = createEmailBody(data, true); // true indicates guest email
      
      var guestOptions = {
        'name': 'Johnny & Josephine Wedding',
        'htmlBody': guestBody,
        'cc': johnnyEmail // CC Johnny on guest email
      };
      
      try {
        GmailApp.sendEmail(data.email, guestSubject, '', guestOptions);
        console.log('Guest confirmation email sent successfully to: ' + data.email);
      } catch(e) {
        console.error('Error sending guest email:', e.toString());
      }
    }
    
    // Send notification email to Johnny
    var adminSubject = 'New Wedding RSVP Response from ' + (data.name || 'Unknown Guest');
    var adminBody = createEmailBody(data, false); // false indicates admin email
    
    var adminOptions = {
      'name': 'Wedding RSVP System',
      'htmlBody': adminBody
    };
    
    try {
      GmailApp.sendEmail(johnnyEmail, adminSubject, '', adminOptions);
      console.log('Admin notification email sent successfully to: ' + johnnyEmail);
    } catch(e) {
      console.error('Error sending admin email:', e.toString());
    }
    
    console.log('Email notifications processed successfully');
    
  } catch(e) {
    console.error('Error in sendEmail function:', e.toString());
    // Don't throw error here - we don't want email issues to break form submission
  }
}

function createEmailBody(data, isGuestEmail) {
  var attendanceStatus = data.attendance === 'yes' ? '&#x2705; 欣然接受' : '&#x274C; 忍痛拒絕';
  var attendanceColor = data.attendance === 'yes' ? '#28a745' : '#dc3545';
  
  // Map dietary options to readable text
  var dietaryText = '';
  if (data.dietary === '') dietaryText = '葷食';
  else if (data.dietary === '1') dietaryText = '素食';
  else if (data.dietary === '2') dietaryText = '減肥中，我絕食';
  else dietaryText = data.dietary || 'Not specified';
  
  // Map vegetarian meals to readable text
  var vegetarianMealsText = '';
  if (data['vegetarian-meals'] && data.dietary === '1') {
    vegetarianMealsText = data['vegetarian-meals'] + ' 份素食';
  } else {
    vegetarianMealsText = 'Not applicable';
  }
  
  // Map children seats count to readable text
  var childrenSeatsCountText = '';
  if (data['children-seats-count'] && data['children-seats'] === 'yes') {
    childrenSeatsCountText = data['children-seats-count'] + ' 張兒童座椅';
  } else {
    childrenSeatsCountText = 'Not applicable';
  }
  
  // Map relation to readable text
  var relationText = '';
  if (data.relation === 'friends-groom') relationText = '男方親友';
  else if (data.relation === 'friends-bride') relationText = '女方親友';
  else if (data.relation === 'other') relationText = '其他';
  else relationText = data.relation || 'Not specified';
  
  // Map children seats to readable text
  var childrenSeatsText = '';
  if (data['children-seats'] === 'yes') childrenSeatsText = '需要';
  else if (data['children-seats'] === 'no') childrenSeatsText = '不需要';
  else childrenSeatsText = 'Not specified';
  
  // Map invitation to readable text
  var invitationText = '';
  if (data.invitation === 'yes') invitationText = '需要';
  else if (data.invitation === 'no') invitationText = '不需要';
  else invitationText = 'Not specified';
  
  var title = isGuestEmail ? 'RSVP 確認通知' : 'New Wedding RSVP Response';
  var subtitle = isGuestEmail ? '感謝您的回覆' : 'New response received';
  
  var html = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #8B4513, #A0522D); padding: 30px; text-align: center; color: white;">
        <h1 style="margin: 0; font-size: 28px; font-family: 'Dancing Script', cursive;">&#x1F495; ${title}</h1>
        <p style="margin: 10px 0 0 0; font-size: 16px;">${subtitle}</p>
      </div>
      
      <div style="background: #f8f9fa; padding: 30px;">
        <div style="background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
          
          ${isGuestEmail ? `
          <div style="text-align: center; margin-bottom: 25px;">
            <p style="font-size: 18px; color: #333; margin: 0;">親愛的 ${data.name || 'Guest'}，</p>
            <p style="font-size: 16px; color: #666; margin: 10px 0;">我們已收到您的 RSVP 回覆，以下是您提交的資訊：</p>
          </div>
          ` : ''}
          
          <div style="text-align: center; margin-bottom: 25px;">
            <span style="background: ${attendanceColor}; color: white; padding: 8px 20px; border-radius: 25px; font-weight: bold; font-size: 14px;">
              ${attendanceStatus}
            </span>
          </div>
          
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333; width: 30%;">
                &#x1F464; 姓名
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${data.name || 'Not provided'}
              </td>
            </tr>
            
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F465; 關係
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${relationText}
              </td>
            </tr>
            
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F4E7; 電子郵件
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                <a href="mailto:${data.email}" style="color: #8B4513; text-decoration: none;">
                  ${data.email || 'Not provided'}
                </a>
              </td>
            </tr>
            
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F465; 出席人數
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${data.guests || '1'} 人
              </td>
            </tr>
            
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F37D; 主菜選擇
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${dietaryText}
              </td>
            </tr>
            
            ${data.dietary === '1' && data['vegetarian-meals'] ? `
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F957; 素食餐數
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${vegetarianMealsText}
              </td>
            </tr>
            ` : ''}
            
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1FA91; 兒童座椅
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${childrenSeatsText}
              </td>
            </tr>
            
            ${data['children-seats'] === 'yes' && data['children-seats-count'] ? `
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1FA92; 兒童座椅數量
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${childrenSeatsCountText}
              </td>
            </tr>
            ` : ''}
            
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F48C; 紙本喜帖
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${invitationText}
              </td>
            </tr>
            
            ${data.address ? `
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F3E0; 寄送地址
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${data.address}
              </td>
            </tr>
            ` : ''}
            
            ${data.message ? `
            <tr>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; font-weight: bold; color: #333;">
                &#x1F48C; 給新人的話
              </td>
              <td style="padding: 12px 0; border-bottom: 1px solid #eee; color: #666;">
                ${data.message}
              </td>
            </tr>
            ` : ''}
            
            <tr>
              <td style="padding: 12px 0; font-weight: bold; color: #333;">
                &#x1F552; 提交時間
              </td>
              <td style="padding: 12px 0; color: #666;">
                ${new Date(new Date().getTime() + 8 * 60 * 60 * 1000).toISOString().replace('T', ' ').replace('Z', ' UTC+8')}
              </td>
            </tr>
          </table>
          
          ${isGuestEmail ? `
          <div style="text-align: center; margin-top: 30px; padding: 20px; background: #f8f9fa; border: 1px solid #eee; border-radius: 10px;">
            <h3 style="color: #8B4513; margin: 0 0 15px 0;">婚禮資訊</h3>
            <p style="margin: 5px 0; color: #333;"><strong>日期：</strong>2025年10月25日 (星期六)</p>
            <p style="margin: 5px 0; color: #333;"><strong>時間：</strong>下午5:30</p>
            <p style="margin: 5px 0; color: #333;"><strong>地點：</strong>晶宴會館-桃園館 三樓詠劇場</p>
            <p style="margin: 5px 0; color: #333;"><strong>地址：</strong>桃園市桃園區南平路166號</p>
            <p style="margin: 15px 0 0 0; color: #666; font-size: 14px;">期待與您共度這個特別的日子！</p>
          </div>

          <!-- 交通資訊 -->
          <div style="margin-top: 20px; padding: 20px; background: #f8f9fa; border: 1px solid #eee; border-radius: 10px;">
            <h3 style="color: #8B4513; text-align: center; margin: 6px 0 16px 0;">交通資訊</h3>

            <div style="text-align: center; margin: 10px 0 20px 0;">
              <img src="https://letsparty20251025.github.io/ParkingInfoDetail.jpg" alt="交通與停車資訊 Traffic & Parking Info" style="max-width: 100%; height: auto; border: 1px solid #eee; border-radius: 6px;" />
            </div>

            <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-bottom: 16px;">
              <tr>
                <td style="width: 33.33%; text-align: center; padding: 8px;">
                  <div style="color: #6c757d; font-size: 14px; margin-bottom: 4px;">Tel</div>
                  <div style="color: #555;">(03)355-3555</div>
                </td>
                <td style="width: 33.33%; text-align: center; padding: 8px;">
                  <div style="color: #6c757d; font-size: 14px; margin-bottom: 4px;">Address</div>
                  <div style="color: #555;">33076 桃園市桃園區南平路166號</div>
                </td>
                <td style="width: 33.33%; text-align: center; padding: 8px;">
                  <div style="color: #6c757d; font-size: 14px; margin-bottom: 4px;">Fax</div>
                  <div style="color: #555;">(03)355-0606</div>
                </td>
              </tr>
            </table>

            <div style="text-align: left;">
              <h4 style="margin: 14px 0 6px 0; color: #8B4513; font-size: 16px;">停車資訊</h4>
              <p style="margin: 0 0 12px 0; color: #333;">賓客停車</p>
              <p style="margin: 0 0 16px 0; color: #555; line-height: 1.6;">
                本館備有停車位置：<br/>
                P：晶宴會館附設停車場 <br/>
                P：晶宴特約停車場，免費三小時停車
              </p>

              <h4 style="margin: 14px 0 6px 0; color: #8B4513; font-size: 16px;">大眾交通工具</h4>
              <p style="margin: 0 0 6px 0; color: #333;">搭乘公車</p>
              <p style="margin: 0 0 12px 0; color: #555; line-height: 1.6;">
                A.請搭乘下列公車至「永昌新城站」下車：151、152。<br/>
                B.請搭乘下列公車至「經國南平路口」下車：201、702、5059。<br/>
                C.請搭乘下列公車至「中正藝文中心」下車：117、綠線GR。 ( 步行8分鐘 )
              </p>

              <p style="margin: 0 0 6px 0; color: #333;">搭乘客運</p>
              <p style="margin: 0 0 12px 0; color: #555; line-height: 1.6;">
                客運直達<br/>
                A.統聯客運南崁站：轉乘201公車至「經國南平路口」下車。<br/>
                B.桃園市政府：搭乘桃園客運188號至「南平同安街口」下車。<br/>
                <br/>
                捷運→客運<br/>
                A.台北市市府轉運站：搭乘9005至「同安街口」下車。(即到達)<br/>
                B.台北市圓山轉運站(玉門)：搭乘9023至「中正藝文特區」下車。(下車後需再直行至南平路右轉，步行約8分鐘)<br/>
                <br/>
                火車→客運<br/>
                內壢火車站：請搭乘統聯客運168至「南平同安街口」下車。<br/>
                <br/>
                高鐵→客運<br/>
                桃園高鐵站：請於桃園客運9號月台搭乘302及707A至「同安街口」下車。(下車即到達)
              </p>

              <h4 style="margin: 14px 0 6px 0; color: #8B4513; font-size: 16px;">自行開車</h4>
              <p style="margin: 0; color: #555; line-height: 1.6;">
                A.國道一號：南崁交流道下往桃園方向 → 靠右經國路 → 南平路右轉。<br/>
                B.國道二號：南桃園交流道下往桃園方向 → 靠左大興西路三段 → 中正路左轉 → 南平路右轉。
              </p>
            </div>
          </div>
          ` : ''}
        </div>
        
        <div style="text-align: center; margin-top: 25px; color: #666; font-size: 14px;">
          <p>此郵件由 RSVP 系統自動發送</p>
          ${isGuestEmail ? '<p>如有任何問題，請聯繫新人</p>' : ''}
        </div>
      </div>
    </div>
  `;
  
  return html;
}

// Optional: Function to get RSVP statistics
function getRSVPStats() {
  try {
    var doc = SpreadsheetApp.openById(SHEET_ID);
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    if (sheet.getLastRow() <= 1) {
      return {
        total: 0,
        attending: 0,
        notAttending: 0,
        totalGuests: 0
      };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    var stats = {
      total: data.length,
      attending: 0,
      notAttending: 0,
      totalGuests: 0
    };
    
    data.forEach(function(row) {
      var attendanceIndex = getColumnIndex(sheet, 'attendance');
      var guestsIndex = getColumnIndex(sheet, 'guests');
      
      if (row[attendanceIndex] === 'yes') {
        stats.attending++;
        stats.totalGuests += parseInt(row[guestsIndex]) || 1;
      } else if (row[attendanceIndex] === 'no') {
        stats.notAttending++;
      }
    });
    
    return stats;
    
  } catch(e) {
    console.error('Error getting RSVP stats:', e.toString());
    return null;
  }
}

function getColumnIndex(sheet, columnName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(columnName);
}

// Optional: Function to send reminder emails (can be triggered manually or with time-based triggers)
function sendReminderEmails() {
  // This is a placeholder for reminder email functionality
  // You can expand this to send reminder emails to people who haven't RSVPed yet
  console.log('Reminder email functionality - to be implemented based on your needs');
}

// Test function - run this to test your setup
function testSetup() {
  try {
    console.log('Testing Google Apps Script setup...');
    
    // Test sheet initialization
    console.log('🔧 Initializing sheet...');
    var initResult = initializeSheet();
    if (initResult.success) {
      console.log('✅ Sheet initialization successful');
    } else {
      throw new Error('Sheet initialization failed: ' + initResult.error);
    }
    
    // Test spreadsheet access
    var doc = SpreadsheetApp.openById(SHEET_ID);
    console.log('✅ Spreadsheet access successful');
    
    // Test sheet access
    var sheet = doc.getSheetByName(SHEET_NAME);
    console.log('✅ Sheet access successful');
    
    // Test email sending
    sendEmail({
      name: 'Test User',
      email: 'test@example.com',
      relation: 'friends-groom',
      attendance: 'yes',
      guests: '2',
      dietary: '1',
      'vegetarian-meals': '2',
      'children-seats': 'yes',
      'children-seats-count': '3',
      invitation: 'yes',
      address: 'Test Address',
      message: 'This is a test submission'
    });
    console.log('✅ Email test completed');
    
    console.log('🎉 Setup test completed successfully!');
    
  } catch(e) {
    console.error('❌ Setup test failed:', e.toString());
  }
}
