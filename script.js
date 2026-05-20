const dates = [];

// Tanggal 26 → 31
for (let i = 26; i <= 31; i++) {
  dates.push(i);
}

// Tanggal 1 → 25
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

// Fix data lama rusak
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
                ${date}
              )
            "
          >
        </td>
      `;

    });

    row += `
      <td class="hk">
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
  date
) {

  if (
    !attendance[personIndex].records
  ) {

    attendance[personIndex].records =
      {};

  }

  attendance[personIndex].records[date] =
    !attendance[personIndex].records[date];

  saveData();

  // Refresh tabel realtime
  renderTable();

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

  // Urut abjad
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

  const header = ["NAMA"];

  dates.forEach((date) => {

    header.push(date.toString());

  });

  header.push("HK");

  data.push(header);

  attendance.forEach((person) => {

    const row = [person.name];

    dates.forEach((date) => {

      row.push(

        person.records &&
        person.records[date]

          ? "✓"

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

  // Lebar kolom excel
  ws["!cols"] = [

    { wch: 30 },

    ...dates.map(() => ({
      wch: 2.5
    })),

    { wch: 5 },

    { wch: 10 }

  ];

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

// Jalankan pertama kali
renderTable();
