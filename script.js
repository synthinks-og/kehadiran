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

const attendanceRef = doc(
  db,
  "attendance",
  "main"
);

const dates = [];

for (let i = 26; i <= 31; i++) {
  dates.push(i);
}

for (let i = 1; i <= 25; i++) {
  dates.push(i);
}

let attendance = [];

onSnapshot(
  attendanceRef,
  (snapshot) => {

    if (snapshot.exists()) {
      attendance =
        snapshot.data().data || [];
    } else {
      attendance = [];
    }

    attendance.forEach((person) => {
      if (!person.records) {
        person.records = {};
      }
    });

    renderTable();

  }
);

async function saveData() {

  await setDoc(
    attendanceRef,
    {
      data: attendance
    }
  );

}

function renderTable() {

  const headerRow =
    document.getElementById(
      "headerRow"
    );

  const tableBody =
    document.getElementById(
      "tableBody"
    );

  const longestNameLength =
    Math.max(
      ...attendance.map(
        (p) => p.name.length
      ),
      10
    );

  document.documentElement
    .style
    .setProperty(
      "--name-column-width",
      `${longestNameLength}ch`
    );

  headerRow.innerHTML = `
    <th>NAMA</th>
  `;

  dates.forEach((date) => {

    headerRow.innerHTML += `
      <th>${date}</th>
    `;

  });

  headerRow.innerHTML += `
    <th>HK</th>
  `;

  tableBody.innerHTML = "";

  attendance.forEach((person, personIndex) => {

    let row = `<tr>`;

    row += `
      <td class="name-column">
        ${person.name}
      </td>
    `;

    dates.forEach((date) => {

      const checked =
        person.records &&
        person.records[date];

      row += `
        <td>
          <input
            type="checkbox"
            ${checked ? "checked" : ""}
            onchange="
              window.toggleAttendance(
                ${personIndex},
                ${date},
                this
              )
            "
          >
        </td>
      `;

    });

    row += `
      <td
        class="hk"
        id="hk-${personIndex}"
      >
        ${getHK(
          person.records || {}
        )}
      </td>
    `;

    row += `</tr>`;

    tableBody.innerHTML += row;

  });

}

window.toggleAttendance =
async function (
  personIndex,
  date,
  checkbox
) {

  if (
    !attendance[personIndex]
      .records
  ) {

    attendance[personIndex]
      .records = {};

  }

  attendance[personIndex]
    .records[date] =
    checkbox.checked;

  document.getElementById(
    `hk-${personIndex}`
  ).innerText =
    getHK(
      attendance[personIndex]
        .records
    );

  await saveData();

};

function getHK(records) {

  return Object.values(records)
    .filter((v) => v)
    .length;

}

window.addName =
async function () {

  const input =
    document.getElementById(
      "newName"
    );

  const newName =
    input.value.trim();

  if (!newName) {
    alert("Masukkan nama");
    return;
  }

  attendance.push({
    name:
      newName.toUpperCase(),
    records: {}
  });

  attendance.sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  input.value = "";

  await saveData();

};

window.openDeleteModal =
function () {

  const modal =
    document.getElementById(
      "deleteModal"
    );

  const select =
    document.getElementById(
      "deleteSelect"
    );

  select.innerHTML = "";

  attendance.forEach((person, index) => {

    select.innerHTML += `
      <option value="${index}">
        ${person.name}
      </option>
    `;

  });

  modal.style.display =
    "flex";

};

window.closeDeleteModal =
function () {

  document.getElementById(
    "deleteModal"
  ).style.display =
    "none";

};

window.confirmDeleteName =
async function () {

  const select =
    document.getElementById(
      "deleteSelect"
    );

  const index =
    select.value;

  if (
    index === "" ||
    index === null
  ) {
    return;
  }

  attendance.splice(index, 1);

  await saveData();

  closeDeleteModal();

};

window.exportExcel =
function () {

  const data = [];

  const activeDates =
    dates.filter((date) => {

      return attendance.some(
        (person) => {

          return (
            person.records &&
            person.records[date]
          );

        }
      );

    });

  // HEADER
  const header = ["NAMA"];

  activeDates.forEach((date) => {

    header.push(
      date.toString()
    );

  });

  header.push("HK");

  data.push(header);

  // DATA
  attendance.forEach((person) => {

    const row = [person.name];

    activeDates.forEach((date) => {

      row.push(
        person.records &&
        person.records[date]
          ? "☑"
          : ""
      );

    });

    row.push(
      getHK(
        person.records || {}
      )
    );

    data.push(row);

  });

  // SHEET
  const ws =
    XLSX.utils.aoa_to_sheet(data);

  // WIDTH
  const nameWidth =
    Math.max(
      ...attendance.map(
        (p) => p.name.length
      ),
      10
    );

  ws["!cols"] = [

    { wch: nameWidth },

    ...activeDates.map(
      () => ({
        wch: 4
      })
    ),

    { wch: 5 }

  ];

  // RANGE
  const range =
    XLSX.utils.decode_range(
      ws["!ref"]
    );

  // STYLE
  for (
    let R = range.s.r;
    R <= range.e.r;
    ++R
  ) {

    for (
      let C = range.s.c;
      C <= range.e.c;
      ++C
    ) {

      const cellAddress =
        XLSX.utils.encode_cell({
          r: R,
          c: C
        });

      if (!ws[cellAddress])
        continue;

      ws[cellAddress].s = {

        alignment: {

          horizontal:
            "center",

          vertical:
            "center"

        },

        border: {

          top: {
            style: "thin",
            color: {
              rgb: "D1D5DB"
            }
          },

          bottom: {
            style: "thin",
            color: {
              rgb: "D1D5DB"
            }
          },

          left: {
            style: "thin",
            color: {
              rgb: "D1D5DB"
            }
          },

          right: {
            style: "thin",
            color: {
              rgb: "D1D5DB"
            }
          }

        }

      };

      // HEADER BIRU
      if (
        R === 0 &&
        C !==
          activeDates.length +
            1
      ) {

        ws[cellAddress]
          .s.fill = {

          fgColor: {
            rgb: "DBEAFE"
          }

        };

        ws[cellAddress]
          .s.font = {

          bold: true

        };

      }

      // HK KUNING
      if (
        C ===
        activeDates.length + 1
      ) {

        ws[cellAddress]
          .s.fill = {

          fgColor: {
            rgb: "FEF08A"
          }

        };

        ws[cellAddress]
          .s.font = {

          bold: true

        };

      }

      // NAMA LEFT
      if (
        C === 0 &&
        R !== 0
      ) {

        ws[cellAddress]
          .s.alignment = {

          horizontal:
            "left",

          vertical:
            "center"

        };

      }

    }

  }

  const wb =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    ws,
    "Rekap Kehadiran"
  );

  XLSX.writeFile(
    wb,
    "rekap-kehadiran.xlsx"
  );

};
