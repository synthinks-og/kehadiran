const dates = [];

for (let i = 25; i <= 31; i++) {
  dates.push(i);
}

for (let i = 1; i <= 25; i++) {
  dates.push(i);
}

let attendance =
  JSON.parse(localStorage.getItem("attendance")) || [];

function saveData() {
  localStorage.setItem(
    "attendance",
    JSON.stringify(attendance)
  );
}

function login() {
  const username =
    document.getElementById("username").value;

  const password =
    document.getElementById("password").value;

  if (
    username === "admin" &&
    password === "admin123"
  ) {
    document.getElementById(
      "loginPage"
    ).style.display = "none";

    document.getElementById(
      "app"
    ).style.display = "block";

    renderTable();
  } else {
    alert("Login gagal");
  }
}

function renderTable() {
  const headerRow =
    document.getElementById("headerRow");

  headerRow.innerHTML = "<th>NAMA</th>";

  dates.forEach((date) => {
    headerRow.innerHTML += `<th>${date}</th>`;
  });

  headerRow.innerHTML += `<th>HK</th>`;

  const tableBody =
    document.getElementById("tableBody");

  tableBody.innerHTML = "";

  attendance.forEach((person, personIndex) => {
    let row = `<tr>`;

    row += `<td>${person.name}</td>`;

    dates.forEach((date) => {
      const checked =
        person.records &&
        person.records[date];

      row += `
        <td>
          <input
            type="checkbox"
            ${checked ? "checked" : ""}
            onchange="toggleAttendance(${personIndex}, ${date})"
          >
        </td>
      `;
    });

    row += `
      <td class="hk">
        ${getHK(person.records || {})}
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

  renderTable();
}

function getHK(records) {
  return Object.values(records).filter(
    (v) => v
  ).length;
}

function addName() {
  const input =
    document.getElementById("newName");

  const newName = input.value.trim();

  if (!newName) {
    alert("Masukkan nama");
    return;
  }

  attendance.push({
    name: newName,
    records: {},
  });

  input.value = "";

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

  ws["!cols"] = [
    { wch: 35 },
    ...dates.map(() => ({ wch: 5 })),
    { wch: 8 },
  ];

  const range =
    XLSX.utils.decode_range(ws["!ref"]);

  for (
    let C = range.s.c;
    C <= range.e.c;
    ++C
  ) {
    const address =
      XLSX.utils.encode_cell({
        r: 0,
        c: C,
      });

    if (!ws[address]) continue;

    ws[address].s = {
      font: {
        bold: true,
      },
      alignment: {
        horizontal: "center",
      },
      fill: {
        fgColor: {
          rgb: "D9EAF7",
        },
      },
    };
  }

  const wb =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    ws,
    "Rekap"
  );

  XLSX.writeFile(
    wb,
    "rekap-kehadiran.xlsx"
  );
}
