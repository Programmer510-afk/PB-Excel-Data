// 📌 এক্সেল ডেটা আনার API লিংক
const apiUrl = "https://prize-bond-data-table-ta12.onrender.com/api/excel-data";

// 🔄 এখানে সম্পূর্ণ ওয়ার্কবুক (Workbook) ডেটা রাখা হবে
let workbookData = null;

// 🔃 পুরো HTML লোড হলে fetch শুরু হবে
window.addEventListener('DOMContentLoaded', () => {
  fetch(apiUrl)
    .then(res => res.json())
    .then(data => {
      workbookData = data;

      // 🗂️ "Header" বাদ দিয়ে বাকি শীটগুলোর নাম descending (নিম্নক্রমে) সাজানো
      const sheetNames = Object.keys(workbookData)
        .filter(name => name !== "Header")
        .sort()
        .reverse();

      // ✅ প্রয়োজনীয় DOM এলিমেন্টগুলো ধরে রাখছি
      const dropdown = document.querySelector('.dropdown');
      const dropdownButton = document.getElementById('dropdown-button');
      const dropdownSelected = document.getElementById('dropdown-selected');
      const dropdownMenu = document.getElementById('dropdown-menu');
      const arrow = document.querySelector('.dropdown-arrow');

      // 🔄 পুরনো মেনু ক্লিয়ার করা
      dropdownMenu.innerHTML = '';

      // 🔽 প্রতিটি শীট নাম দিয়ে ড্রপডাউন মেনুর জন্য অপশন তৈরি করা
      sheetNames.forEach(sheet => {
        const item = document.createElement('a'); // 'a' ব্যবহার করছি কারণ CSS এটাকে স্টাইল করেছে
        item.classList.add('dropdown-item');
        item.textContent = sheet;
        item.href = '#';
        item.addEventListener('click', (e) => {
          e.preventDefault(); // ডিফল্ট লিঙ্ক ক্লিক রোধ

          // 🟢 বেছে নেওয়া শীট নাম বাটনে দেখাও
          dropdownSelected.innerHTML = `<b>${sheet}</b>`;

          // 🔴 ড্রপডাউন মেনু বন্ধ করো এবং অ্যারো রিসেট
          dropdown.classList.remove('show');
          arrow.style.transform = 'translateY(-50%) rotate(0deg)';

          // ✅ সংশ্লিষ্ট টেবিলগুলো আপডেট করো
          loadCol2TableData(sheet);
          loadCol3TableData(sheet);
          loadCol4TableData(sheet);
          loadColumnData(sheet);
        });

        dropdownMenu.appendChild(item); // মেনুতে অপশন যোগ
      });

      // ✅ পেজ লোডের সময় প্রথম শীট বেছে নিয়ে ডিফল্ট টেবিল লোড করা
      dropdownSelected.innerHTML = `<b>${sheetNames[0]}</b>`;
      loadCol2TableData(sheetNames[0]);
      loadCol3TableData(sheetNames[0]);
      loadCol4TableData(sheetNames[0]);
      loadColumnData(sheetNames[0]);
    });

  // 🔘 বাটনে ক্লিক করলে মেনু খুলবে বা বন্ধ হবে এবং অ্যারো ঘুরবে
  document.getElementById('dropdown-button').addEventListener('click', () => {
    const dropdown = document.querySelector('.dropdown');
    const arrow = document.querySelector('.dropdown-arrow');

    const isOpen = dropdown.classList.toggle('show'); // show ক্লাস টগল

    // 🌀 অ্যারো ঘোরাও বা রিসেট করো
    if (isOpen) {
      arrow.style.transform = 'translateY(-50%) rotate(180deg)';
    } else {
      arrow.style.transform = 'translateY(-50%) rotate(0deg)';
    }
  });

  // 📤 ড্রপডাউনের বাইরে ক্লিক করলে মেনু বন্ধ এবং অ্যারো রিসেট হবে
  document.addEventListener('click', (e) => {
    const dropdown = document.querySelector('.dropdown');
    const arrow = document.querySelector('.dropdown-arrow');

    // যদি ক্লিক করা এলিমেন্ট dropdown-এর বাইরে হয়
    if (!dropdown.contains(e.target)) {
      dropdown.classList.remove('show');
      arrow.style.transform = 'translateY(-50%) rotate(0deg)';
    }
  });
});


const button = document.getElementById("dropdown-button");

button.addEventListener("mouseenter", () => {
  button.classList.add("stop-shiny");
});

button.addEventListener("mouseleave", () => {
  button.classList.remove("stop-shiny");
});







// ✅ UPDATED FUNCTION: col2-table =====================================
function loadCol2TableData(sheetName) {
  if (!workbookData || !workbookData[sheetName]) return;

  const data = workbookData[sheetName];
  const table = document.getElementById("col2-table");
  table.innerHTML = "";

  const row1 = document.createElement("tr");
  const row2 = document.createElement("tr");


  // প্রথম দুইটা কলাম সব সময় থাকবে, তাই ২টি ফাঁকা td যোগ করলাম (যেখানে header দিতে পারো)
  // এখানে ফাঁকা রাখলাম, তুমি চাইলে সরাসরি header text বসাতে পারো
  
  
  // 🟩 A1 cell of HTML table ← Excel's F1 (row 0, column 5)
  const a1 = document.createElement("td");
  a1.classList.add("highlight");
  a1.textContent = (data[0]?.[5] ?? "").trim() || "1st Prize"; // fallback if blank
  row1.appendChild(a1);

  // 🟩 A2 cell of HTML table ← Excel's G1 (row 0, column 6)
  const a2 = document.createElement("td");
  a2.classList.add("highlight");
  a2.textContent = (data[0]?.[6] ?? "").trim()  || "2nd Prize"; // fallback if blank
  row2.appendChild(a2);

  // 🟨 B1 cell of HTML table ← Excel's F2 (row 1, column 5)
  const b1 = document.createElement("td");
  b1.classList.add("highlight");
  b1.textContent = data[1]?.[5] ?? "";
  row1.appendChild(b1);

  // 🟨 B2 cell of HTML table ← Excel's G2 (row 1, column 6)
  const b2 = document.createElement("td");
  b2.classList.add("highlight");
  b2.textContent = data[1]?.[6] ?? "";
  row2.appendChild(b2);




  // এক্সেল ডেটার প্রথম কলাম থেকে ডেটা row1 এ যোগ করা হচ্ছে (১ম সারা বাদ দিয়ে)
  for (let i = 1; i < data.length; i++) {
    const val = data[i]?.[0] ?? "";
    const td = document.createElement("td");
    td.textContent = val;
    row1.appendChild(td);
  }

  // এক্সেল ডেটার দ্বিতীয় কলাম থেকে ডেটা row2 তে যোগ করা হচ্ছে (১ম সারা বাদ দিয়ে)
  for (let i = 1; i < data.length; i++) {
    const val = data[i]?.[1] ?? "";
    const td = document.createElement("td");
    td.textContent = val;
    row2.appendChild(td);
  }

  // এখন খালি কলাম গুলো লুকানো হবে, তবে প্রথম দুইটা কলাম (index 0,1) সব সময় দেখাবে
  const row1Tds = Array.from(row1.querySelectorAll('td'));
  const row2Tds = Array.from(row2.querySelectorAll('td'));

  for (let col = 2; col < row1Tds.length; col++) {
    const cell1Text = row1Tds[col].textContent.trim();
    const cell2Text = row2Tds[col].textContent.trim();

    if (cell1Text === "" && cell2Text === "") {
      row1Tds[col].style.display = "none";
      row2Tds[col].style.display = "none";
    } else {
      row1Tds[col].style.display = "";
      row2Tds[col].style.display = "";
    }
  }

  table.appendChild(row1);
  table.appendChild(row2);
}



// ✅ col3-table (3rd Prize) লোড করার ফাংশন
function loadCol3TableData(sheetName) {
  if (!workbookData || !workbookData[sheetName]) return;

  const data = workbookData[sheetName];
  const table = document.getElementById("col3-table");

  const filteredData = data.slice(1).map(row => row[2])
    .filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== '');


  // 🔁 H1 = data[0][7], fallback = "3rd Prize"
  // 🔁 H2 = data[1][7], optional second part
  const part1 = (data[0]?.[7] ?? "").toString().trim() || "3rd Prize";
  const part2 = (data[1]?.[7] ?? "").toString().trim();

  let labelText = part1;
  if (part2) {
     labelText += ` ${part2}`; // 👉 শুধু একটা space দিয়ে part2 যোগ হবে
  }



  const headerRow = table.querySelector("tr.highlight");
  if (headerRow) {
    const headerCell = headerRow.querySelector("td");
    if (headerCell) {
      headerCell.colSpan = 2;
      headerCell.textContent = `${labelText} : Total - ${filteredData.length}`;
    }
  }

  table.querySelectorAll("tr:not(.highlight)").forEach(tr => tr.remove());

  for (let i = 0; i < filteredData.length; i += 2) {
    const tr = document.createElement('tr');
    for (let j = 0; j < 2; j++) {
      const td = document.createElement('td');
      td.textContent = filteredData[i + j] !== undefined ? filteredData[i + j] : '';
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
}




// ✅ col4-table (4th Prize) লোড করার ফাংশন
function loadCol4TableData(sheetName) {
  if (!workbookData || !workbookData[sheetName]) return;

  const data = workbookData[sheetName];
  const table = document.getElementById("col4-table");

  const filteredData = data.slice(1).map(row => row[3])
    .filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== '');


  // 📦 I1 = data[0][8], fallback = "4th Prize"
  // 📦 I2 = data[1][8], optional second part
  const part1 = (data[0]?.[8] ?? "").toString().trim() || "4th Prize";
  const part2 = (data[1]?.[8] ?? "").toString().trim();

  let labelText = part1;
  if (part2) {
  labelText += ` ${part2}`; // 👉 শুধু একটা space দিয়ে 2nd part যোগ হবে (কোনো কমা বা অন্য কিছু না)
  }


  const headerRow = table.querySelector("tr.highlight");
  if (headerRow) {
    const headerCell = headerRow.querySelector("td");
    if (headerCell) {
      headerCell.colSpan = 2;
      headerCell.textContent = `${labelText} : Total - ${filteredData.length}`;
    }
  }

  table.querySelectorAll("tr:not(.highlight)").forEach(tr => tr.remove());

  for (let i = 0; i < filteredData.length; i += 2) {
    const tr = document.createElement('tr');
    for (let j = 0; j < 2; j++) {
      const td = document.createElement('td');
      td.textContent = filteredData[i + j] !== undefined ? filteredData[i + j] : '';
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
}


// ✅ col5-table (5th Prize) লোড করার ফাংশন
function loadColumnData(sheetName) {
  if (!workbookData || !workbookData[sheetName]) return;

  const data = workbookData[sheetName];
  const table = document.getElementById("col5-table");

  const filteredData = data.slice(1).map(row => row[4])
    .filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== '');


  // 📦 J1 = data[0][9] (1st part), fallback = "5th Prize"
  // 📦 J2 = data[1][9] (2nd part), optional
  const part1 = (data[0]?.[9] ?? "").toString().trim() || "5th Prize";
  const part2 = (data[1]?.[9] ?? "").toString().trim();

  let labelText = part1;
  if (part2) {
   labelText += ` ${part2}`; // 👉 শুধু একটা space দিয়ে জোড়া হবে, অন্য কোনো সিম্বল না
  }


  const headerRow = table.querySelector("tr.highlight");
  if (headerRow) {
    const headerCell = headerRow.querySelector("td");
    if (headerCell) {
      headerCell.colSpan = 5;
      headerCell.textContent = `${labelText} : Total - ${filteredData.length}`;
    }
  }

  table.querySelectorAll("tr:not(.highlight)").forEach(tr => tr.remove());

  for (let i = 0; i < filteredData.length; i += 5) {
    const tr = document.createElement('tr');
    for (let j = 0; j < 5; j++) {
      const td = document.createElement('td');
      td.textContent = filteredData[i + j] !== undefined ? filteredData[i + j] : '';
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
}

document.addEventListener('DOMContentLoaded', function () {
  
  

  // For mouse
  dropdown.addEventListener('mousedown', setActiveColor);
  dropdown.addEventListener('mouseup', resetColor);
  dropdown.addEventListener('mouseleave', resetColor);

  // For touch devices
  dropdown.addEventListener('touchstart', setActiveColor);
  dropdown.addEventListener('touchend', resetColor);
  dropdown.addEventListener('touchcancel', resetColor);
});
