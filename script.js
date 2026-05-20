import {
  initializeApp
} from "https://www.gstatic.com/firebasejs/12.13.0/firebase-app.js";

import {
  getFirestore,
  doc,
  setDoc,
  onSnapshot
} from "https://www.gstatic.com/firebasejs/12.13.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyA-QjCZ2Z5rSLNZDXkIHDu-uXKi0VOhiFo",
  authDomain: "rekap-kehadiran.firebaseapp.com",
  projectId: "rekap-kehadiran",
  storageBucket: "rekap-kehadiran.firebasestorage.app",
  messagingSenderId: "544538520853",
  appId: "1:544538520853:web:17b768306620ea0b058d1c"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const attendanceRef = doc(db, "attendance", "main");

const dates = [];

for (let i = 26; i <= 31; i++) dates.push(i);
for (let i = 1; i <= 25; i++) dates.push(i);

let attendance = [];

onSnapshot(attendanceRef, (snapshot) => {
  attendance = snapshot.exists() ? (snapshot.data().data || []) : [];

  attendance.forEach(p => {
    if (!p.records) p.records = {};
  });

  renderTable();
});

async function saveData() {
  await setDoc(attendanceRef, { data: attendance });
}

function renderTable() {

  const headerRow = document.getElementById("headerRow");
  const tableBody = document.getElementById("tableBody");

  const longestNameLength = Math.max(
    ...attendance.map(p => p.name.length),
    10
  );

  document.documentElement.style.setProperty(
    "--name-column-width",
    `${longestNameLength}ch`
  );

  headerRow.innerHTML = `<th>NAMA</th>`;

  dates.forEach(d => {
    headerRow.innerHTML += `<th>${d}</th>`;
  });

  headerRow.innerHTML += `<th>HK</th>`;

  tableBody.innerHTML = "";

  attendance.forEach((person, i) => {

    let row = `<tr>`;

    row += `<td>${person.name}</td>`;

    dates.forEach(date => {

      const checked = person.records?.[date];

      row += `
        <td>
          <input type="checkbox"
            ${checked ? "checked" : ""}
            onchange="toggleAttendance(${i}, ${date}, this)"
          >
        </td>
      `;
    });

    row += `
      <td class="hk" id="hk-${i}">
        ${getHK(person.records || {})}
      </td>
    `;

    row += `</tr>`;

    tableBody.innerHTML += row;
  });
}

window.toggleAttendance = async function(i, date, checkbox) {

  attendance[i].records = attendance[i].records || {};
  attendance[i].records[date] = checkbox.checked;

  document.getElementById(`hk-${i}`).innerText =
    getHK(attendance[i].records);

  await saveData();
};

function getHK(records) {
  return Object.values(records).filter(v => v).length;
}

/* ================= ADD MODAL ================= */

window.openAddModal = function () {
  document.getElementById("addModal").style.display = "flex";
  document.getElementById("addNameInput").value = "";
};

window.closeAddModal = function () {
  document.getElementById("addModal").style.display = "none";
};

window.confirmAddName = async function () {

  const input = document.getElementById("addNameInput");
  const name = input.value.trim();

  if (!name) return alert("Masukkan nama");

  attendance.push({
    name: name.toUpperCase(),
    records: {}
  });

  attendance.sort((a,b) =>
    a.name.localeCompare(b.name)
  );

  await saveData();
  closeAddModal();
};

/* ================= DELETE MODAL ================= */

window.openDeleteModal = function () {

  const modal = document.getElementById("deleteModal");
  const select = document.getElementById("deleteSelect");

  select.innerHTML = "";

  attendance.forEach((p,i) => {
    select.innerHTML += `<option value="${i}">${p.name}</option>`;
  });

  modal.style.display = "flex";
};

window.closeDeleteModal = function () {
  document.getElementById("deleteModal").style.display = "none";
};

window.confirmDeleteName = async function () {

  const index = document.getElementById("deleteSelect").value;

  if (index === "") return;

  attendance.splice(index, 1);

  await saveData();
  closeDeleteModal();
};

/* ================= EXPORT ================= */

window.exportExcel = function () {

  const data = [];

  const activeDates = dates.filter(date =>
    attendance.some(p => p.records?.[date])
  );

  const header = ["NAMA", ...activeDates.map(String), "HK"];

  data.push(header);

  attendance.forEach(p => {

    const row = [p.name];

    activeDates.forEach(date => {
      row.push(p.records?.[date] ? "☑" : "");
    });

    row.push(getHK(p.records || {}));

    data.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Rekap");

  XLSX.writeFile(wb, "rekap-kehadiran.xlsx");
};
