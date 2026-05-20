const dates = [];

// 26 → 31
for (let i = 26; i <= 31; i++) {
  dates.push(i);
}

// 1 → 25
for (let i = 1; i <= 25; i++) {
  dates.push(i);
}

let attendance = [];

try {

  attendance =
    JSON.parse(
      localStorage.getItem("attendance")
    ) || [];

} catch {

  attendance = [];

}

attendance.forEach((person) => {

  if (!person.records) {

    person.records = {};

  }

});

function saveData() {

  localStorage.setItem(
    "attendance",
    JSON.stringify(attendance)
  );

}

function renderTable() {

  const headerRow =
    document.getElementById("headerRow");

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
    <th>AKSI</th>
  `;

  const tableBody =
    document.getElementById("tableBody");

  tableBody.innerHTML = "";

  attendance.forEach((person, personIndex) => {

    let row = `<tr>`;

    row += `
      <td>${person.name}</td>
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
              toggleAttendance(
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
        ${getHK(person.records || {})}
      </td>
    `;

    row += `
      <td>
        <button
          onclick="deleteName(${personIndex})"
          class="delete-btn"
        >
          Hapus
        </button>
      </td>
    `;

    row += `</tr>`;

    tableBody.innerHTML += row;

  });

}

function toggleAttendance(
  personIndex,
  date,
  checkbox
) {

  if (
    !attendance[personIndex].records
  ) {

    attendance[personIndex].records =
      {};

  }

  attendance[personIndex].records[date] =
    checkbox.checked;

  saveData();

  document.getElementById(
    `hk-${personIndex}`
  ).innerText =
    getHK(
      attendance[personIndex].records
    );

}

function getHK(records) {

  return Object.values(records)
    .filter((v) => v).length;

}

function addName() {

  const input =
    document.getElementById("newName");

  const newName =
    input.value.trim();

  if (!newName) {

    alert("Masukkan nama");

    return;

  }

  attendance.push({

    name: newName.toUpperCase(),

    records: {},

  });

  attendance.sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  input.value = "";

  saveData();

  renderTable();

}

function deleteName(index) {

  const confirmDelete = confirm(
    "Hapus nama ini?"
  );

  if (!confirmDelete) return;

  attendance.splice(index, 1);

  saveData();

  renderTable();

}

function exportExcel() {

  const data = [];

  // HEADER
  const header = ["NAMA"];

  dates.forEach((date) => {

    header.push(date.toString());

  });

  header.push("HK");

  data.push(header);

  // DATA
  attendance.forEach((person) => {

    const row = [person.name];

    dates.forEach((date) => {

      row.push(

        person.records &&
        person.records[date]

          ? "☑"

          : ""

      );

    });

    row.push(
      getHK(person.records || {})
    );

    data.push(row);

  });

  const ws =
    XLSX.utils.aoa_to_sheet(data);

  // LEBAR KOLOM
  ws["!cols"] = [

    { wch: 35 },

    ...dates.map(() => ({
      wch: 4
    })),

    { wch: 6 }

  ];

  // STYLE
  const range =
    XLSX.utils.decode_range(ws["!ref"]);

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

      if (!ws[cellAddress]) continue;

      ws[cellAddress].s = {

        border: {

          top: {
            style: "thin",
            color: { rgb: "000000" }
          },

          bottom: {
            style: "thin",
            color: { rgb: "000000" }
          },

          left: {
            style: "thin",
            color: { rgb: "000000" }
          },

          right: {
            style: "thin",
            color: { rgb: "000000" }
          }

        },

        alignment: {

          horizontal: "center",
          vertical: "center"

        }

      };

      // Header biru
      if (R === 0 && C !== dates.length + 1) {

        ws[cellAddress].s.fill = {

          fgColor: {
            rgb: "9DC3E6"
          }

        };

        ws[cellAddress].s.font = {

          bold: true

        };

      }

      // HK kuning
      if (C === dates.length + 1) {

        ws[cellAddress].s.fill = {

          fgColor: {
            rgb: "FFD966"
          }

        };

        ws[cellAddress].s.font = {

          bold: true

        };

      }

      // Nama rata kiri
      if (C === 0 && R !== 0) {

        ws[cellAddress].s.alignment = {

          horizontal: "left",
          vertical: "center"

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

}

renderTable();
