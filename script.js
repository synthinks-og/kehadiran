<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
  />

  <title>Rekap Kehadiran</title>

  <link
    rel="stylesheet"
    href="style.css"
  />

  <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style/dist/xlsx.bundle.js"></script>
</head>

<body>

<div id="app">

  <div class="topbar">

    <h1>Rekap Kehadiran</h1>

    <div class="actions">

      <input
        type="text"
        id="newName"
        placeholder="Tambah nama"
        onkeydown="
          if(event.key === 'Enter'){
            addName()
          }
        "
      >

      <button onclick="addName()">
        Tambah
      </button>

      <button onclick="openDeleteModal()">
        Hapus Nama
      </button>

      <button onclick="exportExcel()">
        Export Excel
      </button>

    </div>

  </div>

  <div class="table-container">

    <table id="attendanceTable">

      <thead>

        <tr id="headerRow">
          <th>NAMA</th>
        </tr>

      </thead>

      <tbody id="tableBody"></tbody>

    </table>

  </div>

</div>

<!-- MODAL HAPUS -->

<div
  id="deleteModal"
  class="modal"
>

  <div class="modal-content">

    <h2>Hapus Nama</h2>

    <select id="deleteSelect">

    </select>

    <div class="modal-actions">

      <button
        class="cancel-btn"
        onclick="closeDeleteModal()"
      >
        Batal
      </button>

      <button
        class="delete-confirm-btn"
        onclick="confirmDeleteName()"
      >
        Hapus
      </button>

    </div>

  </div>

</div>

<footer class="footer">

  Copyright ©
  <a
    href="https://instagram.com/synthinks_"
    target="_blank"
  >
    @Synthinks
  </a>

</footer>

<script src="script.js"></script>

</body>
</html>
