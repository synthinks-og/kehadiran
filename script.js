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

// FIX DATA LAMA
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

  const tableBody =
    document.getElementById("tableBody");

  // AUTO WIDTH KOLOM NAMA
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

  // HEADER
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

  // RESET BODY
  tableBody.innerHTML = "";

  // RENDER DATA
  attendance.forEach((person, personIndex) => {

    let row = `<tr>`;

    // NAMA
    row += `
      <td class="name-column">
        ${person.name}
      </td>
    `;

    // TANGGAL
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

    // HK
    row += `
      <td
        class="hk"
        id="hk-${personIndex}"
      >
        ${getHK(person.records || {})}
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

  // UPDATE HK REALTIME
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

  // SORT ABJAD
  attendance.sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  input.value = "";

  saveData();

  renderTable();

}

function openDeleteModal() {

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

  modal.style.display = "flex";

}

function closeDeleteModal() {

  document.getElementById(
    "deleteModal"
  ).style.display = "none";

}

function confirmDeleteName() {

  const select =
    document.getElementById(
      "deleteSelect"
    );

  const index = select.value;

  if (
    index === "" ||
    index === null
  ) {

    return;

  }

  attendance.splice(index, 1);

  saveData();

  renderTable();

  closeDeleteModal();

}

function exportExcel() {

  const data = [];

  // TANGGAL YANG ADA KEHADIRAN
  const activeDates = dates.filter((date) => {

    return attendance.some((person) => {

      return (
        person.records &&
        person.records[date]
      );

    });

  });

  // HEADER
  const header = ["NAMA"];

  activeDates.forEach((date) => {

    header.push(date.toString());

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
      getHK(person.records || {})
    );

    data.push(row);

  });

  // SHEET
  const ws =
    XLSX.utils.aoa_to_sheet(data);

  // AUTO WIDTH NAMA
  const nameWidth =
    Math.max(
      ...attendance.map(
        (p) => p.name.length
      ),
      10
    );

  // LEBAR KOLOM
  ws["!cols"] = [

    { wch: nameWidth },

    ...activeDates.map(() => ({
      wch: 4
    })),

    { wch: 5 }

  ];

  // RANGE
  const range =
    XLSX.utils.decode_range(ws["!ref"]);

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

      if (!ws[cellAddress]) continue;

      ws[cellAddress].s = {

        alignment: {

          horizontal: "center",
          vertical: "center"

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
        C !== activeDates.length + 1
      ) {

        ws[cellAddress].s.fill = {

          fgColor: {
            rgb: "DBEAFE"
          }

        };

        ws[cellAddress].s.font = {

          bold: true

        };

      }

      // HK KUNING
      if (
        C === activeDates.length + 1
      ) {

        ws[cellAddress].s.fill = {

          fgColor: {
            rgb: "FEF08A"
          }

        };

        ws[cellAddress].s.font = {

          bold: true

        };

      }

      // NAMA RATA KIRI
      if (
        C === 0 &&
        R !== 0
      ) {

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
