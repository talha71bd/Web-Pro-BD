// =======================================================
// Web Pro BD - ভিআইপি ক্যাশ মেমো ও চালান স্টুডিও
// =======================================================

function doGet(e) {
  return HtmlService.createHtmlOutput(getHtmlContent())
    .setTitle("Web Pro BD - অফিসিয়াল চালানের মেমো")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// গুগল শিটে ডাটা সেভ করার আপডেট ব্যাকএন্ড ফাংশন
function saveInvoice(data) {
  try {
    var sheetId = "1U54bauvCMyVWzGF5xLevjrQhQIKevpkDT92t3bqoSME";
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName("Invoices") || ss.getSheets()[0];
    
    // শিট খালি থাকলে নতুন কাস্টম হেডার তৈরি
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "চালান নম্বর", 
        "তারিখ", 
        "গ্রাহকের নাম", 
        "ফোন নম্বর", 
        "ইমেইল",
        "ঠিকানা",
        "সার্ভিস বিবরণী", 
        "সাবটোটাল (টাকা)",
        "ডিসকাউন্ট (টাকা)",
        "মোট মূল্য (টাকা)", 
        "পরিশোধিত (টাকা)", 
        "বকেয়া (টাকা)", 
        "পেমেন্ট মেথড",
        "ট্রানজেকশন আইডি",
        "পেমেন্ট স্ট্যাটাস"
      ]);
      
      var headerRange = sheet.getRange(1, 1, 1, 15);
      headerRange.setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    sheet.appendRow([
      data.invoiceNo,
      data.date,
      data.customerName,
      data.customerPhone,
      data.customerEmail,
      data.customerAddress,
      data.servicesSummary,
      Number(data.subtotal),
      Number(data.discount),
      Number(data.totalPrice),
      Number(data.paidAmount),
      Number(data.dueAmount),
      data.paymentMethod,
      data.trxId,
      data.status
    ]);
    
    return { status: "success", message: "মেমোর যাবতীয় ডাটা গুগল শিটে সফলভাবে সেভ হয়েছে!" };
    
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// ফ্রন্টএন্ড ওয়েবসাইটের সম্পূর্ণ এইচটিএমএল ও কাস্টম ডিজাইন
function getHtmlContent() {
  return `<!DOCTYPE html>
<html lang="bn">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Web Pro BD - ক্যাশ মেমো ও চালান জেনারেটর</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Google Fonts & FontAwesome Icons -->
    <link href="https://fonts.googleapis.com/css2?family=Hind+Siliguri:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <!-- Photo & PDF Download Libraries -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>

    <style>
        body { font-family: 'Hind Siliguri', sans-serif; }
        @media print {
            .no-print { display: none !important; }
            body { background: white !important; padding: 0 !important; }
            #memoCard { box-shadow: none !important; border: none !important; padding: 0 !important; }
        }
    </style>
</head>
<body class="bg-slate-100 text-gray-800 min-h-screen p-3 md:p-6">

    <div class="max-w-7xl mx-auto grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        <!-- বামপাশে: ইনপুট ফর্ম (প্রিন্টের সময় দেখাবে না) -->
        <div class="lg:col-span-5 bg-white p-5 md:p-6 rounded-2xl shadow-xl border border-slate-200 no-print space-y-4">
            <div class="border-b pb-3">
                <h2 class="text-xl font-extrabold text-indigo-600 flex items-center">
                    <i class="fa-solid fa-receipt mr-2"></i> মেমো ইনপুট ফর্ম
                </h2>
                <p class="text-xs text-slate-500">তথ্য দিলে ডানপাশে লাইভ ভিআইপি চালান তৈরি হবে</p>
            </div>

            <form id="invoiceForm" onsubmit="handleFormSubmit(event)" class="space-y-3">
                
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="text-xs font-bold text-slate-600">চালান নম্বর</label>
                        <input type="text" id="inNo" value="WPB-1001" required oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50 font-mono">
                    </div>
                    <div>
                        <label class="text-xs font-bold text-slate-600">তারিখ</label>
                        <input type="date" id="inDate" required onchange="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                    </div>
                </div>

                <div class="border-t border-slate-100 pt-2 space-y-2">
                    <p class="text-xs font-bold text-indigo-600">গ্রাহকের বিবরণ (Customer Details)</p>
                    <div class="grid grid-cols-2 gap-3">
                        <input type="text" id="custName" placeholder="গ্রাহকের নাম *" required oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                        <input type="text" id="custPhone" placeholder="ফোন নম্বর *" required oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <input type="email" id="custEmail" placeholder="গ্রাহকের ইমেইল" oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                        <input type="text" id="custAddress" placeholder="গ্রাহকের ঠিকানা" oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                    </div>
                </div>

                <!-- ডায়নামিক আইটেম সার্ভিস যোগ করার অপশন -->
                <div class="border-t border-slate-100 pt-3">
                    <div class="flex justify-between items-center mb-2">
                        <label class="text-xs font-bold text-indigo-700">সার্ভিস / প্রোডাক্টের তালিকা</label>
                        <button type="button" onclick="addItemRow()" class="text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 px-2.5 py-1 rounded-lg font-bold transition">
                            <i class="fa-solid fa-plus mr-1"></i> আইটেম যোগ করুন
                        </button>
                    </div>
                    <div id="itemsContainer" class="space-y-2">
                        <div class="grid grid-cols-12 gap-2 item-row">
                            <input type="text" placeholder="যেমন: E-Commerce Website" required oninput="updatePreview()" class="col-span-7 p-2 text-xs border rounded-lg item-desc bg-slate-50">
                            <input type="number" placeholder="দাম (৳)" value="15000" required oninput="updatePreview()" class="col-span-4 p-2 text-xs border rounded-lg item-price bg-slate-50">
                            <button type="button" onclick="removeItemRow(this)" class="col-span-1 text-red-500 hover:text-red-700 text-center"><i class="fa-solid fa-trash"></i></button>
                        </div>
                    </div>
                </div>

                <div class="grid grid-cols-2 gap-3 border-t border-slate-100 pt-3">
                    <div>
                        <label class="text-xs font-bold text-slate-600">ডিসকাউন্ট / ছাড় (৳)</label>
                        <input type="number" id="discountAmount" value="0" oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                    </div>
                    <div>
                        <label class="text-xs font-bold text-slate-600">অগ্রিম / জমা (৳)</label>
                        <input type="number" id="paidAmount" value="5000" required oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50">
                    </div>
                </div>

                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="text-xs font-bold text-slate-600">পেমেন্ট মেথড</label>
                        <select id="payMethod" onchange="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50 font-bold">
                            <option value="bKash Personal">bKash (01516513987)</option>
                            <option value="Nagad Personal">Nagad (01516513987)</option>
                            <option value="Bank Transfer">Bank Transfer</option>
                            <option value="Cash Payment">Cash Payment</option>
                        </select>
                    </div>
                    <div>
                        <label class="text-xs font-bold text-slate-600">ট্রানজেকশন আইডি (TrxID)</label>
                        <input type="text" id="trxId" placeholder="TrxID (ঐচ্ছিক)" oninput="updatePreview()" class="w-full p-2.5 text-xs border rounded-xl bg-slate-50 font-mono">
                    </div>
                </div>

                <div class="p-3 bg-indigo-50 rounded-xl flex justify-between items-center text-sm font-bold text-indigo-900">
                    <span>অবশিষ্ট বকেয়া (Due):</span>
                    <span id="dueDisplay" class="text-red-600 text-base">৳ 10,000</span>
                </div>

                <button type="submit" id="submitBtn" class="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-lg transition flex items-center justify-center text-sm">
                    <i class="fa-solid fa-cloud-arrow-up mr-2"></i> গুগল শিটে ডাটা সেভ করুন
                </button>
            </form>
        </div>

        <!-- ডানপাশে: অতি সুন্দর ব্র্যান্ডেড ক্যাশ মেমো / চালান প্রিভিউ -->
        <div class="lg:col-span-7">
            <div id="memoCard" class="bg-white p-6 md:p-8 rounded-2xl shadow-2xl border border-slate-200 relative overflow-hidden">
                
                <!-- ওয়াটারমার্ক/স্ট্যাম্প -->
                <div id="statusStamp" class="absolute top-28 right-10 border-4 border-emerald-500 text-emerald-500 font-extrabold text-2xl px-5 py-1.5 rounded-xl rotate-12 opacity-80 uppercase tracking-widest pointer-events-none">
                    PAID
                </div>

                <!-- হেডার ব্র্যান্ডিং: Web Pro BD -->
                <div class="flex justify-between items-start border-b-2 border-indigo-600 pb-5">
                    <div>
                        <div class="flex items-center space-x-3">
                            <div class="bg-gradient-to-tr from-indigo-600 to-purple-600 text-white p-3 rounded-2xl shadow-md"><i class="fa-solid fa-laptop-code text-2xl"></i></div>
                            <div>
                                <h1 class="text-3xl font-black text-indigo-700 tracking-wide uppercase leading-none">Web Pro BD</h1>
                                <p class="text-[11px] text-purple-700 font-bold mt-1">প্রফেশনাল ওয়েবসাইট ও মোবাইল অ্যাপস ডেভেলপমেন্ট</p>
                                <p class="text-[10px] text-gray-500 font-bold">প্রোপ্রাইটর: আবু তালহা খান</p>
                            </div>
                        </div>
                        <div class="text-[11px] text-slate-600 mt-3 space-y-0.5">
                            <p><i class="fa-solid fa-location-dot text-indigo-600 mr-1.5"></i> ঠিকানা: ফার্মগেট, ঢাকা-১২১৫</p>
                            <p><i class="fa-brands fa-whatsapp text-emerald-600 mr-1.5"></i> হোয়াটসঅ্যাপ / ইমো: <b class="text-slate-800">01516513987</b></p>
                            <p><i class="fa-solid fa-envelope text-indigo-600 mr-1.5"></i> ইমেইল: abutalha46700@gmail.com</p>
                        </div>
                    </div>
                    <div class="text-right">
                        <span class="bg-gradient-to-r from-indigo-600 to-purple-600 text-white text-[11px] font-bold px-3.5 py-1.5 rounded-full uppercase tracking-widest shadow-md">অফিসিয়াল চালান</span>
                        <p class="text-sm font-extrabold text-slate-800 mt-3 font-mono" id="prevInNo">WPB-1001</p>
                        <p class="text-xs text-slate-500 font-semibold" id="prevDate">তারিখ: --/--/----</p>
                    </div>
                </div>

                <!-- লাইফটাইম অফার ব্যানার বিবেশ হাইলাইট -->
                <div class="my-4 p-2.5 bg-gradient-to-r from-indigo-50 to-purple-50 rounded-xl border border-indigo-100 text-center">
                    <p class="text-[11px] font-bold text-indigo-900">
                        ✨ <span class="text-purple-700">বিশেষ সুবিধা:</span> আমাদের থেকে নেওয়া ওয়েবসাইট ও অ্যাপসের কোনো বাৎসরিক বা মাসিক চার্জ নেই — এককালীন ক্রয়ে আজীবন সম্পূর্ণ ফ্রি!
                    </p>
                </div>

                <!-- গ্রাহকের তথ্য ও কিউআর কোড -->
                <div class="my-4 p-4 bg-slate-50 rounded-2xl border border-slate-100 flex justify-between items-center">
                    <div class="space-y-0.5">
                        <p class="text-[10px] text-indigo-600 font-extrabold uppercase tracking-wider">বিল প্রাপক (CUSTOMER DETAILS)</p>
                        <h3 class="text-base font-bold text-slate-900" id="prevCustName">গ্রাহকের নাম</h3>
                        <p class="text-xs text-slate-600" id="prevCustPhone">ফোন: 01700000000</p>
                        <p class="text-[11px] text-slate-500" id="prevCustEmail">ইমেইল: N/A</p>
                        <p class="text-[11px] text-slate-500" id="prevCustAddress">ঠিকানা: N/A</p>
                    </div>
                    
                    <div class="text-right space-y-1">
                        <img id="qrCodeImg" src="" alt="QR Code" class="w-16 h-16 rounded-lg border p-1 bg-white shadow-sm ml-auto">
                        <p class="text-[10px] font-bold text-slate-700">পেমেন্ট: <span id="prevPayMethod">bKash</span></p>
                        <p class="text-[10px] font-mono text-slate-500" id="prevTrxId">TrxID: N/A</p>
                    </div>
                </div>

                <!-- সার্ভিসের টেবিল -->
                <div class="my-4">
                    <table class="w-full text-left text-xs">
                        <thead>
                            <tr class="bg-indigo-600 text-white font-bold uppercase">
                                <th class="p-2.5 rounded-l-xl">ক্রম</th>
                                <th class="p-2.5">সার্ভিস / বিবরণ</th>
                                <th class="p-2.5 text-right rounded-r-xl">মূল্য (Price)</th>
                            </tr>
                        </thead>
                        <tbody id="memoItemsBody" class="divide-y border-b font-semibold">
                            <!-- JS দিয়ে আইটেম যুক্ত হবে -->
                        </tbody>
                    </table>
                </div>

                <!-- মোট হিসাব -->
                <div class="flex justify-end my-4">
                    <div class="w-64 space-y-1.5 text-xs">
                        <div class="flex justify-between text-slate-600">
                            <span>সাবটোটাল (Subtotal):</span>
                            <span class="font-bold text-slate-800" id="sumSubtotal">৳ 15,000</span>
                        </div>
                        <div class="flex justify-between text-amber-600">
                            <span>ডিসকাউন্ট (Discount):</span>
                            <span class="font-bold" id="sumDiscount">- ৳ 0</span>
                        </div>
                        <div class="flex justify-between text-slate-900 font-extrabold text-sm border-t pt-1">
                            <span>সর্বমোট (Total):</span>
                            <span id="sumTotal">৳ 15,000</span>
                        </div>
                        <div class="flex justify-between text-emerald-600 font-bold">
                            <span>জমা / অগ্রিম পরিশোধ:</span>
                            <span id="sumPaid">৳ 5,000</span>
                        </div>
                        <div class="flex justify-between border-t border-slate-300 pt-1 text-sm font-black text-red-600">
                            <span>অবশিষ্ট বকেয়া (Due):</span>
                            <span id="sumDue">৳ 10,000</span>
                        </div>
                    </div>
                </div>

                <!-- শর্তাবলী ও স্বাক্ষর -->
                <div class="border-t border-slate-200 pt-5 mt-6 grid grid-cols-2 gap-4 items-end">
                    <div class="text-[10px] text-slate-400 leading-relaxed">
                        <p class="font-bold text-slate-600">নোট ও শর্তাবলী:</p>
                        <p>১. ডেলিভারির ৩ দিনের মধ্যে বকেয়া পরিশোধ সাপেক্ষে একসেস হস্তান্তর করা হবে।</p>
                        <p>২. আমাদের ওয়েবসাইট ও অ্যাপে কোনো বাৎসরিক বা মাসিক ফি দিতে হয় না।</p>
                    </div>
                    <div class="text-center">
                        <div class="w-36 border-b border-slate-400 ml-auto mb-1"></div>
                        <p class="text-xs font-bold text-slate-800">আবু তালহা খান</p>
                        <p class="text-[10px] text-slate-500 font-semibold">Web Pro BD</p>
                    </div>
                </div>

            </div>

            <!-- মেমো ডাউনলোড ও প্রিন্ট অপশনসমূহ -->
            <div class="mt-6 grid grid-cols-3 gap-3 no-print">
                <button onclick="downloadAsPhoto()" class="py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl shadow-lg transition flex items-center justify-center text-xs">
                    <i class="fa-solid fa-image mr-1.5 text-sm"></i> ছবি (PNG Photo)
                </button>
                <button onclick="downloadAsPDF()" class="py-3 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl shadow-lg transition flex items-center justify-center text-xs">
                    <i class="fa-solid fa-file-pdf mr-1.5 text-sm"></i> PDF ডকুমেন্ট
                </button>
                <button onclick="window.print()" class="py-3 bg-slate-900 hover:bg-black text-white font-bold rounded-xl shadow-lg transition flex items-center justify-center text-xs">
                    <i class="fa-solid fa-print mr-1.5 text-sm"></i> প্রিন্ট করুন
                </button>
            </div>
        </div>

    </div>

    <script>
        document.addEventListener("DOMContentLoaded", function() {
            document.getElementById('inDate').valueAsDate = new Date();
            updatePreview();
        });

        function addItemRow() {
            const container = document.getElementById('itemsContainer');
            const row = document.createElement('div');
            row.className = 'grid grid-cols-12 gap-2 item-row';
            row.innerHTML = \`
                <input type="text" placeholder="সার্ভিসের নাম..." required oninput="updatePreview()" class="col-span-7 p-2 text-xs border rounded-lg item-desc bg-slate-50">
                <input type="number" placeholder="দাম (৳)" value="0" required oninput="updatePreview()" class="col-span-4 p-2 text-xs border rounded-lg item-price bg-slate-50">
                <button type="button" onclick="removeItemRow(this)" class="col-span-1 text-red-500 hover:text-red-700 text-center"><i class="fa-solid fa-trash"></i></button>
            \`;
            container.appendChild(row);
            updatePreview();
        }

        function removeItemRow(btn) {
            const rows = document.querySelectorAll('.item-row');
            if(rows.length > 1) {
                btn.closest('.item-row').remove();
                updatePreview();
            } else {
                alert("কমপক্ষে একটি আইটেম রাখা আবশ্যক!");
            }
        }

        function updatePreview() {
            const inNo = document.getElementById('inNo').value;
            const inDate = document.getElementById('inDate').value;
            const name = document.getElementById('custName').value || 'গ্রাহকের নাম';
            const phone = document.getElementById('custPhone').value || '01700000000';
            const email = document.getElementById('custEmail').value || 'N/A';
            const address = document.getElementById('custAddress').value || 'N/A';
            const method = document.getElementById('payMethod').value;
            const trxId = document.getElementById('trxId').value || 'N/A';

            const descs = document.querySelectorAll('.item-desc');
            const prices = document.querySelectorAll('.item-price');
            const tbody = document.getElementById('memoItemsBody');
            tbody.innerHTML = '';

            let subtotal = 0;

            descs.forEach((descInput, index) => {
                const desc = descInput.value || 'সার্ভিসের বিবরণ...';
                const price = parseFloat(prices[index].value) || 0;
                subtotal += price;

                tbody.innerHTML += \`
                    <tr>
                        <td class="p-2.5 text-slate-400 font-mono">\${index + 1}</td>
                        <td class="p-2.5 font-bold text-slate-800">\${desc}</td>
                        <td class="p-2.5 text-right font-extrabold text-slate-900">৳ \${price.toLocaleString('bn-BD')}</td>
                    </tr>
                \`;
            });

            const discount = parseFloat(document.getElementById('discountAmount').value) || 0;
            const totalPrice = Math.max(0, subtotal - discount);
            const paid = parseFloat(document.getElementById('paidAmount').value) || 0;
            const due = Math.max(0, totalPrice - paid);

            document.getElementById('prevInNo').innerText = inNo;
            document.getElementById('prevDate').innerText = 'তারিখ: ' + inDate;
            document.getElementById('prevCustName').innerText = name;
            document.getElementById('prevCustPhone').innerText = 'ফোন: ' + phone;
            document.getElementById('prevCustEmail').innerText = 'ইমেইল: ' + email;
            document.getElementById('prevCustAddress').innerText = 'ঠিকানা: ' + address;
            document.getElementById('prevPayMethod').innerText = method;
            document.getElementById('prevTrxId').innerText = 'TrxID: ' + trxId;

            document.getElementById('sumSubtotal').innerText = '৳ ' + subtotal.toLocaleString('bn-BD');
            document.getElementById('sumDiscount').innerText = '- ৳ ' + discount.toLocaleString('bn-BD');
            document.getElementById('sumTotal').innerText = '৳ ' + totalPrice.toLocaleString('bn-BD');
            document.getElementById('sumPaid').innerText = '৳ ' + paid.toLocaleString('bn-BD');
            document.getElementById('sumDue').innerText = '৳ ' + due.toLocaleString('bn-BD');
            document.getElementById('dueDisplay').innerText = '৳ ' + due.toLocaleString('bn-BD');

            // কিউআর কোড
            const qrData = \`Web Pro BD | Invoice: \${inNo} | Customer: \${name} | Total: \${totalPrice} | Due: \${due}\`;
            document.getElementById('qrCodeImg').src = \`https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=\${encodeURIComponent(qrData)}\`;

            // স্ট্যাম্প
            const stamp = document.getElementById('statusStamp');
            if (due === 0 && totalPrice > 0) {
                stamp.innerText = "PAID";
                stamp.className = "absolute top-28 right-10 border-4 border-emerald-500 text-emerald-500 font-extrabold text-2xl px-5 py-1.5 rounded-xl rotate-12 opacity-80 uppercase tracking-widest pointer-events-none";
            } else {
                stamp.innerText = "DUE";
                stamp.className = "absolute top-28 right-10 border-4 border-red-500 text-red-500 font-extrabold text-2xl px-5 py-1.5 rounded-xl rotate-12 opacity-80 uppercase tracking-widest pointer-events-none";
            }
        }

        // ছবি (Photo PNG) হিসেবে ডাউনলোড
        function downloadAsPhoto() {
            const card = document.getElementById('memoCard');
            html2canvas(card, { scale: 2 }).then(canvas => {
                const link = document.createElement('a');
                link.download = document.getElementById('inNo').value + '-WebProBD.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
            });
        }

        // PDF ডকুমেন্ট হিসেবে ডাউনলোড
        function downloadAsPDF() {
            const card = document.getElementById('memoCard');
            const opt = {
                margin:       0.3,
                filename:     document.getElementById('inNo').value + '-WebProBD.pdf',
                image:        { type: 'jpeg', quality: 0.98 },
                html2canvas:  { scale: 2 },
                jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
            };
            html2pdf().set(opt).from(card).save();
        }

        // ফর্ম সাবমিট ও গুগল শিটে অটো সেভ
        function handleFormSubmit(e) {
            e.preventDefault();

            const btn = document.getElementById('submitBtn');
            btn.disabled = true;
            btn.innerHTML = \`<i class="fa-solid fa-spinner animate-spin mr-2"></i> শিটে সেভ হচ্ছে...\`;

            const descs = document.querySelectorAll('.item-desc');
            const prices = document.querySelectorAll('.item-price');
            let subtotal = 0;
            let summaryList = [];

            descs.forEach((descInput, index) => {
                const desc = descInput.value || 'সার্ভিস';
                const price = parseFloat(prices[index].value) || 0;
                subtotal += price;
                summaryList.push(desc + ' (৳' + price + ')');
            });

            const discount = parseFloat(document.getElementById('discountAmount').value) || 0;
            const totalPrice = Math.max(0, subtotal - discount);
            const paid = parseFloat(document.getElementById('paidAmount').value) || 0;
            const due = Math.max(0, totalPrice - paid);

            const payload = {
                invoiceNo: document.getElementById('inNo').value,
                date: document.getElementById('inDate').value,
                customerName: document.getElementById('custName').value,
                customerPhone: document.getElementById('custPhone').value,
                customerEmail: document.getElementById('custEmail').value,
                customerAddress: document.getElementById('custAddress').value,
                servicesSummary: summaryList.join(', '),
                subtotal: subtotal,
                discount: discount,
                totalPrice: totalPrice,
                paidAmount: paid,
                dueAmount: due,
                paymentMethod: document.getElementById('payMethod').value,
                trxId: document.getElementById('trxId').value,
                status: due === 0 ? "PAID" : "DUE"
            };

            google.script.run
                .withSuccessHandler(function(response) {
                    btn.disabled = false;
                    btn.innerHTML = \`<i class="fa-solid fa-cloud-arrow-up mr-2"></i> গুগল শিটে ডাটা সেভ করুন\`;
                    if (response.status === 'success') {
                        alert("✅ " + response.message);
                    } else {
                        alert("❌ এরর: " + response.message);
                    }
                })
                .withFailureHandler(function(err) {
                    btn.disabled = false;
                    btn.innerHTML = \`<i class="fa-solid fa-cloud-arrow-up mr-2"></i> গুগল শিটে ডাটা সেভ করুন\`;
                    alert("❌ সমস্যা হয়েছে: " + err);
                })
                .saveInvoice(payload);
        }
    </script>
</body>
</html>`;
}