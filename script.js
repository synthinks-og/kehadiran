const dates = [];

for(let i=25; i<=31; i++){
  dates.push(i);
}

for(let i=1; i<=25; i++){
  dates.push(i);
}

let attendance = JSON.parse(localStorage.getItem("attendance")) || [
  {
    name:"MUHAMMAD NURUL HIDAYAT",
    records:{}
  },
  {
    name:"REVI AKBAR",
    records:{}
  }
];

function saveData(){
  localStorage.setItem("attendance", JSON.stringify(attendance));
}

function login(){
  const username = document.getElementById("username").value;
  const password = document.getElementById("password").value;

  if(username === "admin" && password === "admin123"){
    document.getElementById("loginPage").style.display = "none";
    document.getElementById("app").style.display = "block";

    renderTable();
  }else{
    alert("Login gagal");
  }
}

function renderTable(){

  const headerRow = document.getElementById("headerRow");

  headerRow.innerHTML = "<th>NAMA</th>";

  dates.forEach(date => {
    headerRow.innerHTML += `<th>${date}</th>`;
  });

  headerRow.innerHTML += "<th>HK</th>";

  const tableBody = document.getElementById("tableBody");

  tableBody.innerHTML = "";

  attendance.forEach((person, personIndex) => {

    let row = `<tr>`;

    row += `<td>${person.name}</td>`;

    dates.forEach(date => {

      const checked = person.records[date] ? "checked" : "";

      row += `
        <td>
          <input
            type="checkbox"
            ${checked}
            onchange="toggleAttendance(${personIndex}, ${date})"
          >
        </td>
      `;
    });

    row += `<td class="hk">${getHK(person.records)}</td>`;

    row += `</tr>`;

    tableBody.innerHTML += row;
  });
}

function toggleAttendance(personIndex, date){

  attendance[personIndex].records[date] =
    !attendance[personIndex].records[date];

  saveData();

  renderTable();
}

function getHK(records){

  return Object.values(records).filter(v => v).length;
}

function addName(){

  const newName = document.getElementById("newName").value;

  if(!newName.trim()) return;

  attendance.push({
    name:newName,
    records:{}
  });

  document.getElementById("newName").value = "";

  saveData();

  renderTable();
}

function exportExcel(){

  const table = document.getElementById("attendanceTable");

  const wb = XLSX.utils.table_to_book(table, {
    sheet:"Rekap Kehadiran"
  });

  XLSX.writeFile(wb, "rekap-kehadiran.xlsx");
}