/**
 * Web Pro BD - Universal Backend Engine
 * Developer: AI Assistant for Web Pro BD
 */

const SHEET_ID = "1cp1_07k8mUGX3fWDGMoCn2NMjvCbappW8wpcRD8CZhs";

// ১. এই ফাংশনটি একবার রান করলে তোমার শিট পুরো সাজানো হয়ে যাবে
function setupProfessionalSystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const structure = [
    { name: "Settings", headers: ["Option_Name", "Value", "Description"], colors: "#2c3e50" },
    { name: "Products", headers: ["ID", "Title", "Category", "Price", "Discount_Price", "Image_URL", "Preview_URL", "Status", "Features"], colors: "#16a085" },
    { name: "Orders", headers: ["Order_ID", "Date", "Customer_Name", "Phone", "Product_Name", "Total_Amount", "Payment_Status", "Order_Status"], colors: "#2980b9" },
    { name: "Messages", headers: ["Date", "Name", "Phone/Email", "Message", "Reply_Status"], colors: "#8e44ad" },
    { name: "Reviews", headers: ["Client_Name", "Comment", "Rating", "Status"], colors: "#d35400" }
  ];

  structure.forEach(obj => {
    let sheet = ss.getSheetByName(obj.name);
    if (!sheet) sheet = ss.insertSheet(obj.name);
    
    sheet.clear(); // নতুন করে সাজানোর জন্য
    sheet.appendRow(obj.headers);
    
    // ডিজাইন ও ফরম্যাটিং
    const headerRange = sheet.getRange(1, 1, 1, obj.headers.length);
    headerRange.setBackground(obj.colors)
               .setFontColor("white")
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
    
    sheet.setFrozenRows(1); // উপরের রো ফ্রিজ করা
    sheet.setColumnWidth(1, 150);
  });

  // ডাটা ভ্যালিডেশন (ড্রপডাউন) তৈরি
  setupDropdowns(ss);
  
  // কিছু ডিফল্ট ডাটা ও সেটিংস অ্যাড করা
  populateDefaultSettings(ss);
  
  Browser.msgBox("Web Pro BD System: সফলভাবে শিট সাজানো হয়েছে!");
}

function setupDropdowns(ss) {
  const activeStatus = SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Inactive', 'Draft']).build();
  ss.getSheetByName("Products").getRange("H2:H100").setDataValidation(activeStatus);
  
  const orderStatus = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Processing', 'Completed', 'Cancelled']).build();
  ss.getSheetByName("Orders").getRange("H2:H500").setDataValidation(orderStatus);

  const payStatus = SpreadsheetApp.newDataValidation().requireValueInList(['Paid', 'Unpaid', 'Partial']).build();
  ss.getSheetByName("Orders").getRange("G2:G500").setDataValidation(payStatus);
}

function populateDefaultSettings(ss) {
  const sheet = ss.getSheetByName("Settings");
  const defaults = [
    ["Site_Name", "Web Pro BD", "আপনার ওয়েবসাইটের নাম"],
    ["WhatsApp", "8801xxxxxxxxx", "সরাসরি কন্টাক্ট নম্বর"],
    ["Primary_Color", "#3498db", "ওয়েবসাইটের মূল থিম কালার"],
    ["Banner_Title", "সাশ্রয়ী মূল্যে প্রিমিয়াম ওয়েব অ্যাপ", "হোম পেজের বড় টেক্সট"],
    ["SEO_Description", "Best Web Solution in BD", "গুগল সার্চের জন্য বর্ণনা"]
  ];
  sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
}

// ২. ওয়েবসাইট থেকে ডাটা রিড করার জন্য (API Get)
function doGet(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const result = {};
  
  ss.getSheets().forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    result[sheet.getName()] = data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  });

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ৩. ওয়েবসাইট থেকে অর্ডার বা মেসেজ সেভ করার জন্য (API Post)
function doPost(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const postData = JSON.parse(e.postData.contents);
  const sheet = ss.getSheetByName(postData.target);
  
  if (postData.target === "Orders") {
    sheet.appendRow([
      "ORD-" + Math.floor(Math.random()*10000), 
      new Date(), 
      postData.name, 
      postData.phone, 
      postData.product, 
      postData.price, 
      "Unpaid", 
      "Pending"
    ]);
  } else if (postData.target === "Messages") {
    sheet.appendRow([new Date(), postData.name, postData.contact, postData.message, "New"]);
  }

  return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
    .setMimeType(ContentService.MimeType.JSON);
}