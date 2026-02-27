/* ============================================
   Sınav Salon Dağıtım - Main Application
   Offline-capable, no external dependencies
   ============================================ */

(function () {
  'use strict';

  // ===== STATE =====
  const state = {
    step: 1,
    students: [],
    rooms: [],
    distribution: null,
    editingRoomId: null,
    studentPage: 1,
    studentsPerPage: 25,
    roomPage: 1,
    roomsPerPage: 15,
    examInfo: {
      name: '',
      date: '',
      time: '',
      institution: '',
      location: ''
    },
    resultPage: 1,
    resultsPerPage: 1,
    opticalFormType: null // 'lgs' or 'tyt'
  };

  // ===== DOM REFS =====
  let DOM = {};

  // ===== INIT =====
  function init() {
    cacheDOMRefs();
    setupEventListeners();
    setupDragAndDrop('student-upload-area', 'student-file-input', handleStudentUpload);
    setupDragAndDrop('room-upload-area', 'room-file-input', handleRoomUpload);

    setupDragAndDrop('exam-upload-area', 'exam-file-input', handleExamInfoUpload);
    updateUI();
  }

  function cacheDOMRefs() {
    DOM.steps = {
      1: document.getElementById('step-1'),
      2: document.getElementById('step-2'),
      3: document.getElementById('step-3'),
      4: document.getElementById('step-4'),
      5: document.getElementById('step-5')
    };
    DOM.indicators = {
      1: document.getElementById('step-indicator-1'),
      2: document.getElementById('step-indicator-2'),
      3: document.getElementById('step-indicator-3'),
      4: document.getElementById('step-indicator-4'),
      5: document.getElementById('step-indicator-5')
    };
    DOM.connectors = {
      1: document.getElementById('connector-1'),
      2: document.getElementById('connector-2'),
      3: document.getElementById('connector-3'),
      4: document.getElementById('connector-4')
    };
    DOM.studentCount = document.getElementById('student-count');
    DOM.roomCount = document.getElementById('room-count');
    DOM.totalCapacity = document.getElementById('total-capacity');
    DOM.studentPreview = document.getElementById('student-preview');
    DOM.studentUploadArea = document.getElementById('student-upload-area');
    DOM.examUploadArea = document.getElementById('exam-upload-area');
    DOM.examPreview = document.getElementById('exam-info-preview');
    DOM.toastContainer = document.getElementById('toast-container');
  }

  // ===== TOAST NOTIFICATIONS =====
  function showToast(message, type) {
    type = type || 'info';
    var icons = {
      success: '✅',
      error: '❌',
      warning: '⚠️',
      info: 'ℹ️'
    };
    var toast = document.createElement('div');
    toast.className = 'toast toast-' + type;
    toast.innerHTML = '<span>' + (icons[type] || icons.info) + '</span><span>' + message + '</span>';
    DOM.toastContainer.appendChild(toast);

    setTimeout(function () {
      toast.classList.add('toast-removing');
      setTimeout(function () { toast.remove(); }, 300);
    }, 3500);
  }

  // ===== DRAG & DROP =====
  function setupDragAndDrop(areaId, fileInputId, callback) {
    var area = document.getElementById(areaId);
    if (!area) return;

    function preventDefaults(e) {
      e.preventDefault();
      e.stopPropagation();
    }

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (evt) {
      area.addEventListener(evt, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(function (evt) {
      area.addEventListener(evt, function () { area.classList.add('drag-over'); }, false);
    });

    ['dragleave', 'drop'].forEach(function (evt) {
      area.addEventListener(evt, function () { area.classList.remove('drag-over'); }, false);
    });

    area.addEventListener('dragover', function (e) {
      e.dataTransfer.dropEffect = 'copy';
    });

    area.addEventListener('drop', function (e) {
      var dt = e.dataTransfer;
      var files = dt.files;
      if (files.length > 0) callback(files[0]);
    }, false);

    // Click to open file dialog (avoid double trigger)
    area.addEventListener('click', function (e) {
      // If user clicked the button directly, let button logic handle it.
      // If user clicked outside button (on the div), trigger logic.
      if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
      document.getElementById(fileInputId).click();
    });
  }

  // ===== UI UPDATE =====
  function updateUI() {
    // Steps visibility & indicators
    for (var key = 1; key <= 5; key++) {
      if (key === state.step) {
        DOM.steps[key].classList.remove('hidden');
        DOM.steps[key].classList.add('animate-in');
        DOM.indicators[key].classList.add('active');
        DOM.indicators[key].classList.remove('completed');
      } else {
        DOM.steps[key].classList.add('hidden');
        DOM.indicators[key].classList.remove('active');
      }
      if (key < state.step) {
        DOM.indicators[key].classList.add('completed');
      } else if (key > state.step) {
        DOM.indicators[key].classList.remove('completed');
      }
    }

    // Connectors
    for (var c = 1; c <= 4; c++) {
      DOM.connectors[c].classList.remove('completed', 'active');
      if (c < state.step) DOM.connectors[c].classList.add('completed');
      else if (c === state.step) DOM.connectors[c].classList.add('active');
    }

    // Counts
    DOM.studentCount.textContent = state.students.length;
    DOM.roomCount.textContent = state.rooms.length;
    var totalCap = state.rooms.reduce(function (sum, r) { return sum + r.capacity; }, 0);
    DOM.totalCapacity.textContent = totalCap;
  }

  function setStep(step) {
    state.step = step;
    updateUI();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ===== EXCEL PARSING =====
  function parseExcelFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var data = new Uint8Array(e.target.result);
          var workbook = XLSX.read(data, { type: 'array' });
          var sheetName = workbook.SheetNames[0];
          var worksheet = workbook.Sheets[sheetName];
          var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          resolve(jsonData);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = function (err) { reject(err); };
      reader.readAsArrayBuffer(file);
    });
  }

  // ===== STUDENT HANDLING =====
  function handleStudentUpload(file) {
    if (typeof XLSX === 'undefined') {
      showToast('Excel kütüphanesi (XLSX) yüklenemedi. Lütfen sayfayı yenileyin.', 'error');
      return;
    }



    parseExcelFile(file).then(function (data) {
      if (!data || data.length === 0) {
        showToast('Dosya boş veya okunamadı.', 'error');
        return;
      }
      if (data.length === 1) {
        showToast('Dosyada öğrenci kaydı bulunamadı (sadece başlık var).', 'error');
        return;
      }

      var rows = data.slice(1);
      state.students = rows.map(function (row) {
        var name = (row[0] != null) ? String(row[0]) : '';
        var surname = (row[1] != null) ? String(row[1]) : '';
        var fullName = (name + ' ' + surname).trim();

        return {
          id: (row[3] != null) ? String(row[3]) : '',
          tc: (row[4] != null) ? String(row[4]) : '',
          name: fullName,
          phone: (row[2] != null) ? String(row[2]) : '',
          classRef: (row[5] != null) ? String(row[5]) : '',
          department: (row[6] != null) ? String(row[6]) : ''
        };
      }).filter(function (s) { return s.name.length > 0; });
      state.studentPage = 1;

      renderStudentTable();
      updateUI();
      DOM.studentPreview.classList.remove('hidden');
      DOM.studentUploadArea.classList.add('hidden');


    }).catch(function (err) {
      console.error(err);
      showToast('Dosya okuma hatası: ' + err.message, 'error');
    });
  }

  function renderStudentTable() {
    var thead = document.querySelector('#student-table thead');
    var tbody = document.querySelector('#student-table tbody');

    thead.innerHTML =
      '<tr>' +
      '<th>#</th>' +
      '<th>No</th>' +
      '<th>TC No</th>' +
      '<th>Ad Soyad</th>' +
      '<th>Sınıf</th>' +
      '<th>Bölüm</th>' +
      '<th>Telefon</th>' +
      '</tr>';

    tbody.innerHTML = '';

    var perPage = state.studentsPerPage;
    var page = state.studentPage;
    var totalPages = Math.ceil(state.students.length / perPage);
    if (page > totalPages) page = totalPages;
    if (page < 1) page = 1;
    state.studentPage = page;

    var startIdx = (page - 1) * perPage;
    var endIdx = Math.min(startIdx + perPage, state.students.length);
    var list = state.students.slice(startIdx, endIdx);

    list.forEach(function (s, i) {
      var tr = document.createElement('tr');
      tr.innerHTML =
        '<td>' + (startIdx + i + 1) + '</td>' +
        '<td>' + escapeHtml(s.id) + '</td>' +
        '<td>' + escapeHtml(s.tc) + '</td>' +
        '<td>' + escapeHtml(s.name) + '</td>' +
        '<td>' + escapeHtml(s.classRef) + '</td>' +
        '<td>' + escapeHtml(s.department) + '</td>' +
        '<td>' + escapeHtml(s.phone) + '</td>';
      tbody.appendChild(tr);
    });

    // Render pagination
    renderStudentPagination(totalPages);
  }

  function renderStudentPagination(totalPages) {
    var existingPag = document.getElementById('student-pagination');
    if (existingPag) existingPag.remove();

    if (totalPages <= 1) return;

    var container = document.createElement('div');
    container.id = 'student-pagination';
    container.className = 'pagination';

    // Prev button
    var prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn' + (state.studentPage <= 1 ? ' disabled' : '');
    prevBtn.innerHTML = '&laquo;';
    prevBtn.disabled = state.studentPage <= 1;
    prevBtn.addEventListener('click', function () {
      if (state.studentPage > 1) { state.studentPage--; renderStudentTable(); }
    });
    container.appendChild(prevBtn);

    // Page numbers
    var startPage = Math.max(1, state.studentPage - 2);
    var endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

    if (startPage > 1) {
      container.appendChild(createPageBtn(1));
      if (startPage > 2) {
        var dots = document.createElement('span');
        dots.className = 'pagination-dots';
        dots.textContent = '...';
        container.appendChild(dots);
      }
    }

    for (var p = startPage; p <= endPage; p++) {
      container.appendChild(createPageBtn(p));
    }

    if (endPage < totalPages) {
      if (endPage < totalPages - 1) {
        var dots2 = document.createElement('span');
        dots2.className = 'pagination-dots';
        dots2.textContent = '...';
        container.appendChild(dots2);
      }
      container.appendChild(createPageBtn(totalPages));
    }

    // Next button
    var nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn' + (state.studentPage >= totalPages ? ' disabled' : '');
    nextBtn.innerHTML = '&raquo;';
    nextBtn.disabled = state.studentPage >= totalPages;
    nextBtn.addEventListener('click', function () {
      if (state.studentPage < totalPages) { state.studentPage++; renderStudentTable(); }
    });
    container.appendChild(nextBtn);

    // Info text
    var info = document.createElement('span');
    info.className = 'pagination-info';
    info.textContent = state.students.length + ' öğrenci, Sayfa ' + state.studentPage + '/' + totalPages;
    container.appendChild(info);

    // Insert after table
    var tableContainer = document.querySelector('#student-table').closest('.data-table-container');
    tableContainer.parentNode.insertBefore(container, tableContainer.nextSibling);
  }

  function createPageBtn(pageNum) {
    var btn = document.createElement('button');
    btn.className = 'pagination-btn' + (pageNum === state.studentPage ? ' active' : '');
    btn.textContent = pageNum;
    btn.addEventListener('click', function () {
      state.studentPage = pageNum;
      renderStudentTable();
    });
    return btn;
  }

  // ===== ROOM HANDLING =====
  function handleRoomUpload(file) {
    parseExcelFile(file).then(function (data) {
      if (!data || data.length < 2) {
        showToast('Dosyada salon bilgisi bulunamadı.', 'error');
        return;
      }

      var rows = data.slice(1);
      var newRooms = rows.map(function (row, index) {
        return {
          id: Date.now() + index,
          name: row[0] || ('Salon ' + (index + 1)),
          capacity: parseInt(row[1]) || 20,
          priority: parseInt(row[2]) || 999
        };
      }).filter(function (r) { return r.capacity > 0; });

      state.rooms = state.rooms.concat(newRooms);
      state.rooms.sort(function (a, b) { return a.priority - b.priority; });
      updateUI();
      renderRoomTable();
      showToast(newRooms.length + ' salon başarıyla eklendi.', 'success');

    }).catch(function (err) {
      console.error(err);
      showToast('Dosya okuma hatası: ' + err.message, 'error');
    });
  }

  function renderRoomTable() {
    var tbody = document.querySelector('#room-table tbody');
    var nameInput = document.getElementById('manual-room-name');
    var capInput = document.getElementById('manual-room-capacity');
    var priorityInput = document.getElementById('manual-room-priority');
    var btnAdd = document.getElementById('btn-add-room');

    tbody.innerHTML = '';
    state.rooms.sort(function (a, b) { return a.priority - b.priority; });

    // Pagination calculations
    var perPage = state.roomsPerPage;
    var page = state.roomPage;
    var totalPages = Math.ceil(state.rooms.length / perPage);
    if (totalPages < 1) totalPages = 1;
    if (page > totalPages) page = totalPages;
    if (page < 1) page = 1;
    state.roomPage = page;

    var startIdx = (page - 1) * perPage;
    var endIdx = Math.min(startIdx + perPage, state.rooms.length);
    var pagedRooms = state.rooms.slice(startIdx, endIdx);

    pagedRooms.forEach(function (r, i) {
      var globalIndex = startIdx + i;
      var tr = document.createElement('tr');
      if (r.id === state.editingRoomId) tr.classList.add('editing');

      tr.innerHTML =
        '<td>' + escapeHtml(r.name) + '</td>' +
        '<td>' + r.capacity + '</td>' +
        '<td>' + (r.priority === 999 ? '—' : r.priority) + '</td>' +
        '<td><button class="btn-delete" data-index="' + globalIndex + '"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg> Sil</button></td>';

      // Double click to edit
      tr.addEventListener('dblclick', function () {
        state.editingRoomId = r.id;
        nameInput.value = r.name;
        capInput.value = r.capacity;
        priorityInput.value = r.priority === 999 ? '' : r.priority;
        btnAdd.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>';
        btnAdd.classList.remove('btn-accent');
        btnAdd.classList.add('btn-warning');
        renderRoomTable();
        nameInput.focus();
      });

      tbody.appendChild(tr);
    });

    // Delete handlers
    document.querySelectorAll('#room-table .btn-delete').forEach(function (btn) {
      btn.addEventListener('click', function (e) {
        var idx = parseInt(e.currentTarget.getAttribute('data-index'));
        var room = state.rooms[idx];
        if (room && state.editingRoomId === room.id) {
          cancelEditMode();
        }
        state.rooms.splice(idx, 1);
        updateUI();
        renderRoomTable();
        showToast('Salon silindi.', 'warning');
      });
    });

    // Render room pagination
    renderRoomPagination(totalPages);
  }

  function renderRoomPagination(totalPages) {
    var existingPag = document.getElementById('room-pagination');
    if (existingPag) existingPag.remove();

    if (totalPages <= 1) return;

    var container = document.createElement('div');
    container.id = 'room-pagination';
    container.className = 'pagination';

    // Prev button
    var prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn' + (state.roomPage <= 1 ? ' disabled' : '');
    prevBtn.innerHTML = '&laquo;';
    prevBtn.disabled = state.roomPage <= 1;
    prevBtn.addEventListener('click', function () {
      if (state.roomPage > 1) { state.roomPage--; renderRoomTable(); }
    });
    container.appendChild(prevBtn);

    // Page numbers
    var startPage = Math.max(1, state.roomPage - 2);
    var endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

    if (startPage > 1) {
      container.appendChild(createRoomPageBtn(1));
      if (startPage > 2) {
        var dots = document.createElement('span');
        dots.className = 'pagination-dots';
        dots.textContent = '...';
        container.appendChild(dots);
      }
    }

    for (var p = startPage; p <= endPage; p++) {
      container.appendChild(createRoomPageBtn(p));
    }

    if (endPage < totalPages) {
      if (endPage < totalPages - 1) {
        var dots2 = document.createElement('span');
        dots2.className = 'pagination-dots';
        dots2.textContent = '...';
        container.appendChild(dots2);
      }
      container.appendChild(createRoomPageBtn(totalPages));
    }

    // Next button
    var nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn' + (state.roomPage >= totalPages ? ' disabled' : '');
    nextBtn.innerHTML = '&raquo;';
    nextBtn.disabled = state.roomPage >= totalPages;
    nextBtn.addEventListener('click', function () {
      if (state.roomPage < totalPages) { state.roomPage++; renderRoomTable(); }
    });
    container.appendChild(nextBtn);

    // Info text
    var info = document.createElement('span');
    info.className = 'pagination-info';
    info.textContent = state.rooms.length + ' salon, Sayfa ' + state.roomPage + '/' + totalPages;
    container.appendChild(info);

    // Insert after room table
    var tableContainer = document.querySelector('#room-table').closest('.data-table-container');
    tableContainer.parentNode.insertBefore(container, tableContainer.nextSibling);
  }

  function createRoomPageBtn(pageNum) {
    var btn = document.createElement('button');
    btn.className = 'pagination-btn' + (pageNum === state.roomPage ? ' active' : '');
    btn.textContent = pageNum;
    btn.addEventListener('click', function () {
      state.roomPage = pageNum;
      renderRoomTable();
    });
    return btn;
  }

  function cancelEditMode() {
    state.editingRoomId = null;
    var nameInput = document.getElementById('manual-room-name');
    var capInput = document.getElementById('manual-room-capacity');
    var priorityInput = document.getElementById('manual-room-priority');
    var btnAdd = document.getElementById('btn-add-room');

    nameInput.value = '';
    capInput.value = '';
    priorityInput.value = '';
    btnAdd.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>';
    btnAdd.classList.remove('btn-warning');
    btnAdd.classList.add('btn-accent');
  }

  // ===== DISTRIBUTION =====
  function distributeStudents() {
    var studentsToDistribute = state.students.slice();
    shuffleArray(studentsToDistribute);

    var result = {};
    state.rooms.forEach(function (r) { result[r.id] = []; });

    var overflow = [];
    var currentRoomIndex = 0;

    studentsToDistribute.forEach(function (student) {
      var placed = false;

      while (currentRoomIndex < state.rooms.length) {
        var room = state.rooms[currentRoomIndex];
        if (result[room.id].length < room.capacity) {
          result[room.id].push(student);
          placed = true;
          break;
        } else {
          currentRoomIndex++;
        }
      }

      if (!placed) overflow.push(student);
    });

    state.distribution = { results: result, overflow: overflow };
    state.resultPage = 1;
    renderDistributionResults();
  }

  function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = array[i];
      array[i] = array[j];
      array[j] = temp;
    }
  }

  function renderDistributionResults() {
    var summaryEl = document.getElementById('distribution-summary');
    var resultsEl = document.getElementById('distribution-results');
    var totalDistributed = state.students.length - state.distribution.overflow.length;
    var hasOverflow = state.distribution.overflow.length > 0;

    // Summary cards
    summaryEl.innerHTML =
      '<div class="summary-grid">' +
      '<div class="summary-card summary-total">' +
      '<div class="summary-value">' + state.students.length + '</div>' +
      '<div class="summary-label">Toplam Öğrenci</div>' +
      '</div>' +
      '<div class="summary-card summary-placed">' +
      '<div class="summary-value">' + totalDistributed + '</div>' +
      '<div class="summary-label">Yerleşen</div>' +
      '</div>' +
      '<div class="summary-card ' + (hasOverflow ? 'summary-overflow' : 'summary-overflow no-overflow') + '">' +
      '<div class="summary-value">' + state.distribution.overflow.length + '</div>' +
      '<div class="summary-label">' + (hasOverflow ? 'Açıkta Kalan' : 'Herkes Yerleşti!') + '</div>' +
      '</div>' +
      '</div>' +
      (hasOverflow ?
        '<div style="margin-bottom: 1.5rem;"><button id="btn-export-overflow" class="btn btn-danger btn-sm">⚠️ Açıkta Kalanları İndir (Excel)</button></div>'
        : '');

    // Overflow export button
    var btnOverflow = document.getElementById('btn-export-overflow');
    if (btnOverflow) {
      btnOverflow.addEventListener('click', function () {
        exportOverflowToExcel(state.distribution.overflow);
      });
    }

    // Room result cards
    var html = '';

    // Pagination calculations
    var perPage = state.resultsPerPage;
    var page = state.resultPage;
    var totalPages = Math.ceil(state.rooms.length / perPage);
    if (totalPages < 1) totalPages = 1;
    if (page > totalPages) page = totalPages;
    if (page < 1) page = 1;
    state.resultPage = page;

    var startIdx = (page - 1) * perPage;
    var endIdx = Math.min(startIdx + perPage, state.rooms.length);
    var pagedRooms = state.rooms.slice(startIdx, endIdx);

    pagedRooms.forEach(function (room) {
      var students = state.distribution.results[room.id];
      html +=
        '<div class="room-result-card">' +
        '<div class="room-result-header">' +
        '<h3>' + escapeHtml(room.name) + '</h3>' +
        '<span class="capacity-badge">' + students.length + ' / ' + room.capacity + '</span>' +
        '</div>' +
        '<div class="data-table-container">' +
        '<table>' +
        '<thead><tr><th>Sıra</th><th>No</th><th>TC No</th><th>Ad Soyad</th><th>Sınıf</th><th>Bölüm</th><th>Telefon</th></tr></thead>' +
        '<tbody>' +
        (students.length === 0 ?
          '<tr><td colspan="7" style="text-align:center; color:var(--text-muted);">Öğrenci yok</td></tr>' :
          students.map(function (s, i) {
            return '<tr>' +
              '<td>' + (i + 1) + '</td>' +
              '<td>' + escapeHtml(String(s.id)) + '</td>' +
              '<td>' + escapeHtml(String(s.tc)) + '</td>' +
              '<td>' + escapeHtml(s.name) + '</td>' +
              '<td>' + escapeHtml(String(s.classRef)) + '</td>' +
              '<td>' + escapeHtml(String(s.department)) + '</td>' +
              '<td>' + escapeHtml(String(s.phone)) + '</td>' +
              '</tr>';
          }).join('')) +
        '</tbody>' +
        '</table>' +
        '</div>' +
        '</div>';
    });

    resultsEl.innerHTML = html;

    // Render Pagination
    renderResultPagination(totalPages);

    // Overflow section - Only show on last page or if there's only one page
    if (hasOverflow && (state.resultPage === totalPages)) {
      var overflowHtml =
        '<div class="room-result-card overflow-result-card" style="margin-top: 2rem;">' +
        '<div class="room-result-header">' +
        '<h3>⚠️ Açıkta Kalan Öğrenciler</h3>' +
        '<span class="capacity-badge">' + state.distribution.overflow.length + ' öğrenci</span>' +
        '</div>' +
        '<div class="data-table-container">' +
        '<table>' +
        '<thead><tr><th>No</th><th>TC No</th><th>Ad Soyad</th><th>Sınıf</th><th>Bölüm</th><th>Telefon</th></tr></thead>' +
        '<tbody>' +
        state.distribution.overflow.map(function (s) {
          return '<tr>' +
            '<td>' + escapeHtml(String(s.id)) + '</td>' +
            '<td>' + escapeHtml(String(s.tc)) + '</td>' +
            '<td>' + escapeHtml(s.name) + '</td>' +
            '<td>' + escapeHtml(String(s.classRef)) + '</td>' +
            '<td>' + escapeHtml(String(s.department)) + '</td>' +
            '<td>' + escapeHtml(String(s.phone)) + '</td>' +
            '</tr>';
        }).join('') +
        '</tbody>' +
        '</table>' +
        '</div>' +
        '</div>';

      resultsEl.insertAdjacentHTML('beforeend', overflowHtml);
    }
  }

  function renderResultPagination(totalPages) {
    var existingPag = document.getElementById('result-pagination');
    if (existingPag) existingPag.remove();

    if (totalPages <= 1) return;

    var container = document.createElement('div');
    container.id = 'result-pagination';
    container.className = 'pagination';
    container.style.justifyContent = 'center';
    container.style.marginBottom = '2rem';

    // Prev button
    var prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn' + (state.resultPage <= 1 ? ' disabled' : '');
    prevBtn.innerHTML = '&laquo;';
    prevBtn.disabled = state.resultPage <= 1;
    prevBtn.addEventListener('click', function () {
      if (state.resultPage > 1) { state.resultPage--; renderDistributionResults(); }
    });
    container.appendChild(prevBtn);

    // Page numbers
    var startPage = Math.max(1, state.resultPage - 2);
    var endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

    if (startPage > 1) {
      container.appendChild(createResultPageBtn(1));
      if (startPage > 2) {
        var dots = document.createElement('span');
        dots.className = 'pagination-dots';
        dots.textContent = '...';
        container.appendChild(dots);
      }
    }

    for (var p = startPage; p <= endPage; p++) {
      container.appendChild(createResultPageBtn(p));
    }

    if (endPage < totalPages) {
      if (endPage < totalPages - 1) {
        var dots2 = document.createElement('span');
        dots2.className = 'pagination-dots';
        dots2.textContent = '...';
        container.appendChild(dots2);
      }
      container.appendChild(createResultPageBtn(totalPages));
    }

    // Next button
    var nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn' + (state.resultPage >= totalPages ? ' disabled' : '');
    nextBtn.innerHTML = '&raquo;';
    nextBtn.disabled = state.resultPage >= totalPages;
    nextBtn.addEventListener('click', function () {
      if (state.resultPage < totalPages) { state.resultPage++; renderDistributionResults(); }
    });
    container.appendChild(nextBtn);

    // Info text
    var info = document.createElement('span');
    info.className = 'pagination-info';
    info.textContent = state.rooms.length + ' salon, Sayfa ' + state.resultPage + '/' + totalPages;
    container.appendChild(info);

    // Insert into resultsEl
    var resultsEl = document.getElementById('distribution-results');
    resultsEl.appendChild(container);
  }

  function createResultPageBtn(pageNum) {
    var btn = document.createElement('button');
    btn.className = 'pagination-btn' + (pageNum === state.resultPage ? ' active' : '');
    btn.textContent = pageNum;
    btn.addEventListener('click', function () {
      state.resultPage = pageNum;
      renderDistributionResults();
    });
    return btn;
  }

  // ===== EXCEL EXPORT =====
  function exportToExcel(data, rooms) {
    var wb = XLSX.utils.book_new();

    // Overview sheet
    var overviewData = [];
    rooms.forEach(function (room) {
      var students = data.results[room.id] || [];
      students.forEach(function (s, i) {
        overviewData.push({
          'Salon Adı': room.name,
          'Sıra No': i + 1,
          'Öğrenci No': s.id,
          'TC No': s.tc,
          'Ad Soyad': s.name,
          'Sınıf': s.classRef,
          'Bölüm': s.department,
          'Telefon': s.phone
        });
      });
    });

    if (data.overflow && data.overflow.length > 0) {
      data.overflow.forEach(function (s) {
        overviewData.push({
          'Salon Adı': 'YERLEŞEMEDİ',
          'Sıra No': '-',
          'Öğrenci No': s.id,
          'TC No': s.tc,
          'Ad Soyad': s.name,
          'Sınıf': s.classRef,
          'Bölüm': s.department,
          'Telefon': s.phone
        });
      });
    }

    var wsOverview = XLSX.utils.json_to_sheet(overviewData);
    XLSX.utils.book_append_sheet(wb, wsOverview, 'Genel Liste');

    // Per-room sheets
    rooms.forEach(function (room) {
      var students = data.results[room.id] || [];
      if (students.length > 0) {
        var roomData = students.map(function (s, i) {
          return {
            'Sıra No': i + 1,
            'Öğrenci No': s.id,
            'TC No': s.tc,
            'Ad Soyad': s.name,
            'Sınıf': s.classRef,
            'Bölüm': s.department,
            'Telefon': s.phone
          };
        });
        var wsRoom = XLSX.utils.json_to_sheet(roomData);

        var sheetName = room.name.replace(/[\\/?*[\]:]/g, ' ').trim();
        if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
        if (!sheetName) sheetName = 'Salon ' + room.id;

        var uniqueName = sheetName;
        var counter = 1;
        while (wb.SheetNames.indexOf(uniqueName) !== -1) {
          uniqueName = sheetName.substring(0, 28) + '(' + counter + ')';
          counter++;
        }

        XLSX.utils.book_append_sheet(wb, wsRoom, uniqueName);
      }
    });

    XLSX.writeFile(wb, 'sinav_dagitim_sonuclari.xlsx');
    showToast('Excel dosyası indirildi.', 'success');
  }

  function exportRoomsToExcel(rooms) {
    var wb = XLSX.utils.book_new();
    var sorted = rooms.slice().sort(function (a, b) { return (a.priority || 999) - (b.priority || 999); });
    var exportData = sorted.map(function (r) {
      return {
        'Salon Adı': r.name,
        'Kapasite': r.capacity,
        'Öncelik': r.priority === 999 ? '' : r.priority
      };
    });
    var ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Salonlar');
    XLSX.writeFile(wb, 'salon_listesi.xlsx');
    showToast('Salon listesi indirildi.', 'success');
  }

  function exportOverflowToExcel(overflowData) {
    var wb = XLSX.utils.book_new();
    var exportData = overflowData.map(function (s) {
      return {
        'Öğrenci No': s.id,
        'TC No': s.tc,
        'Ad Soyad': s.name,
        'Sınıf': s.classRef,
        'Bölüm': s.department,
        'Telefon': s.phone
      };
    });
    var ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [{ wch: 15 }, { wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 20 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Açıkta Kalanlar');
    XLSX.writeFile(wb, 'acikta_kalanlar.xlsx');
    showToast('Açıkta kalanlar listesi indirildi.', 'success');
  }

  // ===== EXAM INFO HANDLING =====
  function handleExamInfoUpload(file) {
    parseExcelFile(file).then(function (data) {
      if (!data || data.length < 2) { // Header + 1 row
        showToast('Dosyada bilgi bulunamadı.', 'error');
        return;
      }

      var row = data[1]; // First data row
      // Expected structure: [Exam Name, Exam Date, Exam Location]
      var name = row[0] ? String(row[0]).trim() : '';
      var dateRaw = row[1];
      var timeRaw = row[2]; // Assuming time is next
      var location = row[3] ? String(row[3]).trim() : '';
      var institution = row[4] ? String(row[4]).trim() : '';

      // Date parsing
      var dateStr = '';
      if (dateRaw) {
        if (typeof dateRaw === 'number') {
          var dateObj = new Date(Math.round((dateRaw - 25569) * 86400 * 1000));
          dateStr = dateObj.toISOString().split('T')[0];
        } else {
          dateStr = String(dateRaw);
        }
      }

      state.examInfo = {
        name: name,
        date: dateStr,
        time: timeRaw ? String(timeRaw) : '',
        location: location,
        institution: institution
      };

      updateExamInfoUI();
      showToast('Sınav bilgileri yüklendi.', 'success');

    }).catch(function (err) {
      console.error(err);
      showToast('Dosya okuma hatası: ' + err.message, 'error');
    });
  }

  function updateExamInfoUI() {
    document.getElementById('exam-name').value = state.examInfo.name;
    document.getElementById('exam-date').value = state.examInfo.date;
    document.getElementById('exam-time').value = state.examInfo.time;
    document.getElementById('exam-institution').value = state.examInfo.institution;
    document.getElementById('exam-location').value = state.examInfo.location;

    document.getElementById('exam-preview-name').textContent = state.examInfo.name || '—';
    document.getElementById('exam-preview-date').textContent = state.examInfo.date || '—';
    // Add time/location to preview if needed, or just keep simple preview
    document.getElementById('exam-preview-location').textContent = state.examInfo.location || '—';

    DOM.examUploadArea.classList.add('hidden');
    DOM.examPreview.classList.remove('hidden');
  }

  function downloadExamInfoTemplate() {
    var wb = XLSX.utils.book_new();
    var wsData = [
      ['Sınav Adı', 'Sınav Tarihi (YYYY-MM-DD)', 'Sınav Saati', 'Adres', 'Sınav Yeri'],
      ['2026 Bahar Final', '2026-06-15', '09:30', 'Merkez Kampüs, Bursa', 'Edesis Eğitim Kurumları']
    ];
    var ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 30 }, { wch: 20 }, { wch: 15 }, { wch: 30 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, ws, 'SinavBilgi');
    XLSX.writeFile(wb, 'sinav_bilgi_sablonu.xlsx');
  }

  // ===== PDF EXPORT =====
  function generatePDF(distributionData, rooms) {
    var jspdf = window.jspdf;
    var doc = new jspdf.jsPDF();

    var sortedRooms = rooms.slice().sort(function (a, b) { return (a.priority || 999) - (b.priority || 999); });
    var isFirstPage = true;

    sortedRooms.forEach(function (room) {
      var students = distributionData.results[room.id] || [];
      if (students.length === 0) return;

      if (!isFirstPage) {
        doc.addPage();
      } else {
        isFirstPage = false;
      }

      doc.setFontSize(16);
      doc.text('Salon: ' + transliterate(room.name), 14, 20);
      doc.setFontSize(10);
      doc.text('Kapasite: ' + students.length + ' / ' + room.capacity, 14, 28);

      var tableData = students.map(function (s, i) {
        return [
          i + 1,
          transliterate(String(s.id)),
          transliterate(String(s.tc)),
          transliterate(s.name),
          transliterate(String(s.classRef)),
          transliterate(String(s.department)),
          transliterate(String(s.phone))
        ];
      });

      doc.autoTable({
        startY: 35,
        head: [['SIRA', 'NO', 'TC', 'AD SOYAD', 'SINIF', 'BOLUM', 'TEL']],
        body: tableData,
        theme: 'grid',
        headStyles: { fillColor: [79, 70, 229] },
        styles: { fontSize: 9, cellPadding: 2, font: 'helvetica' },
        columnStyles: {
          0: { cellWidth: 15 },
          1: { cellWidth: 20 },
          2: { cellWidth: 25 },
          6: { cellWidth: 25 }
        }
      });
    });

    // Overflow page
    var overflow = distributionData.overflow || [];
    if (overflow.length > 0) {
      if (!isFirstPage) doc.addPage();
      doc.setFontSize(16);
      doc.setTextColor(220, 38, 38);
      doc.text('Acikta Kalanlar Listesi', 14, 20);
      doc.setTextColor(0, 0, 0);

      var tableData = overflow.map(function (s, i) {
        return [
          i + 1,
          transliterate(String(s.id)),
          transliterate(String(s.tc)),
          transliterate(s.name),
          transliterate(String(s.classRef)),
          transliterate(String(s.department)),
          transliterate(String(s.phone))
        ];
      });

      doc.autoTable({
        startY: 30,
        head: [['SIRA', 'NO', 'TC', 'AD SOYAD', 'SINIF', 'BOLUM', 'TEL']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [220, 38, 38] },
        styles: { fontSize: 9, font: 'helvetica' }
      });
    }

    doc.save('sinav_dagitim_raporu.pdf');
    showToast('PDF raporu indirildi.', 'success');
  }



  // ===== PDF ENTRY DOCS =====
  function generateEntryDocumentsPDF() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error('PDF kütüphanesi yüklenemedi. Sayfayı yenileyip tekrar deneyin.');
    }
    var jsPDF = window.jspdf.jsPDF;
    var doc = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });

    // Add Fonts
    if (window.fontRobotoRegular && window.fontRobotoBold) {
      doc.addFileToVFS('Roboto-Regular.ttf', window.fontRobotoRegular);
      doc.addFileToVFS('Roboto-Bold.ttf', window.fontRobotoBold);
      doc.addFont('Roboto-Regular.ttf', 'Roboto', 'normal');
      doc.addFont('Roboto-Bold.ttf', 'Roboto', 'bold');
      doc.setFont('Roboto', 'normal');
    }

    var logoImg = document.getElementById('header-logo').src;

    // Sort rooms
    var sortedRooms = state.rooms.slice().sort(function (a, b) { return (a.priority || 999) - (b.priority || 999); });

    var docIndex = 0;

    sortedRooms.forEach(function (room) {
      var students = state.distribution.results[room.id] || [];

      students.forEach(function (student, i) {
        var position = docIndex % 2; // 0 = top, 1 = bottom
        if (docIndex > 0 && position === 0) {
          doc.addPage();
        }

        // Calculate Y offset (Top: 10mm, Bottom: 158mm)
        // A4 height = 297mm. Half = 148.5mm.
        var startY = position === 0 ? 10 : 158;

        drawEntryDocument(doc, startY, student, room, i + 1, logoImg);

        docIndex++;
      });
    });

    doc.save('sinav_giris_belgeleri.pdf');
    showToast(docIndex + ' adet giriş belgesi oluşturuldu.', 'success');
  }

  function drawEntryDocument(doc, y, student, room, seatNo, logoImg) {
    // Colors
    var blueColor = [100, 149, 237]; // CornflowerBlue
    var redColor = [255, 105, 120]; // Light Red/Pinkish

    var width = 190;
    var x = 10;

    // --- Header ---
    // Logo (Centered now)
    var headerEnd = y + 20; // Reduced initial spacing
    if (logoImg) {
      try {
        var logoW = 18; // Slightly smaller logo
        var logoH = 18;
        var logoX = (210 - logoW) / 2;
        doc.addImage(logoImg, 'PNG', logoX, y, logoW, logoH);
        headerEnd = y + 20;
      } catch (e) { /* ignore */ }
    }

    // Title removed as requested
    // doc.setFontSize(14);
    // if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    // else doc.setFont(undefined, 'bold');

    // var title = state.examInfo.institution || 'SINAV GİRİŞ BELGESİ';
    // doc.text(title.toLocaleUpperCase('tr-TR'), 105, headerEnd + 8, { align: 'center' });

    // Exam Name removed from top header as requested

    // --- Box 1: Student Info ---
    var box1Y = headerEnd + 8; // Reduced gap (was 15)

    // Header (Purple Gradient-ish)
    doc.setFillColor(83, 109, 254); // Indigo/Purple
    doc.roundedRect(x, box1Y, width, 12, 2, 2, 'F'); // Height increased to 12

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14); // Font increased to 14

    if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    else doc.setFont(undefined, 'bold');

    doc.text('ÖĞRENCİ BİLGİLERİ', 105, box1Y + 8, { align: 'center' }); // Y adjusted

    // Body (White with border)
    doc.setDrawColor(200, 200, 200); // Light grey border
    doc.setFillColor(255, 255, 255); // White bg
    doc.roundedRect(x, box1Y + 12, width, 24, 2, 2, 'FD'); // Start Y +12

    doc.setFontSize(10);

    // Helper for Label:Value pairs
    function drawField(label, value, xPos, yPos) {
      doc.setTextColor(100, 100, 100); // Grey Label
      if (window.fontRobotoRegular) doc.setFont('Roboto', 'normal');
      else doc.setFont(undefined, 'normal');
      doc.text(label, xPos, yPos);

      // Value
      var labelWidth = doc.getTextWidth(label);
      doc.setTextColor(0, 0, 0); // Black Value
      if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
      else doc.setFont(undefined, 'bold');
      doc.text(value, xPos + labelWidth + 2, yPos);
    }

    // Row 1
    drawField('Adı ve Soyadı:', student.name.toLocaleUpperCase('tr-TR'), x + 5, box1Y + 18); // Check Y: 12 + 6 = 18
    drawField('Sınıf:', String(student.classRef).toLocaleUpperCase('tr-TR'), x + 120, box1Y + 18);

    // Row 2
    drawField('TC Kimlik No:', String(student.tc), x + 5, box1Y + 24); // 18 + 6 = 24
    drawField('Telefon:', String(student.phone), x + 120, box1Y + 24);

    // Row 3
    // drawField('Okul:', (state.examInfo.institution || '-').toLocaleUpperCase('tr-TR'), x + 5, box1Y + 26);

    // --- Box 2: Exam Info ---
    var box2Y = box1Y + 42; // Increased gap (was 38)

    // Header (Pink/Red)
    doc.setFillColor(255, 64, 129); // Pink
    doc.roundedRect(x, box2Y, width, 12, 2, 2, 'F'); // Height 12

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14); // Font 14

    if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    else doc.setFont(undefined, 'bold');

    doc.text('SINAV GİRİŞ BİLGİLERİ', 105, box2Y + 8, { align: 'center' }); // Y adjusted

    // Body Box
    doc.setDrawColor(200, 200, 200);
    doc.setFillColor(255, 255, 255);
    doc.roundedRect(x, box2Y + 12, width, 50, 2, 2, 'S'); // Start Y +12

    // 1. Exam Name Strip 
    doc.setFillColor(225, 245, 254);
    doc.roundedRect(x + 2, box2Y + 14, width - 4, 10, 2, 2, 'F'); // Y +14

    doc.setTextColor(0, 0, 0);
    doc.setFontSize(11); // Reset font for content
    if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    doc.text((state.examInfo.name || '').toLocaleUpperCase('tr-TR'), 105, box2Y + 20.5, { align: 'center' }); // Y +20.5

    // 2. Time & Salon Strip
    doc.setFillColor(255, 224, 178);
    doc.roundedRect(x + 2, box2Y + 26, width - 4, 16, 2, 2, 'F'); // Y +26

    // Time
    doc.setTextColor(50, 50, 50);
    doc.setFontSize(9);
    if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    doc.text('SINAV SAATİ', x + 50, box2Y + 31, { align: 'center' }); // Y +31

    doc.setTextColor(0, 0, 0);
    doc.setFontSize(14);
    doc.text((state.examInfo.time || '--:--'), x + 50, box2Y + 38, { align: 'center' }); // Y +38

    // Salon
    doc.setTextColor(50, 50, 50);
    doc.setFontSize(9);
    doc.text('SALON NO / SIRA NO', x + 140, box2Y + 31, { align: 'center' }); // Y +31

    doc.setTextColor(0, 0, 0);
    doc.setFontSize(14);
    doc.text(room.name.toLocaleUpperCase('tr-TR') + ' / ' + seatNo, x + 140, box2Y + 38, { align: 'center' }); // Y +38

    // 3. Footer Info (White area)
    doc.setFontSize(10);

    // Row 1
    var dateStr = state.examInfo.date || '';
    drawField('Sınav Tarihi:', formatDateTR(dateStr), x + 10, box2Y + 49); // Y +49
    drawField('Sınav Yeri:', (state.examInfo.institution || '').toLocaleUpperCase('tr-TR'), x + 100, box2Y + 49); // Y +49

    // Row 2
    drawField('Adres:', (state.examInfo.location || '').toLocaleUpperCase('tr-TR'), x + 10, box2Y + 56); // Y +56

    // Dashed separator line if top
    if (y < 100) {
      doc.setLineDash([2, 2], 0);
      doc.setDrawColor(200, 200, 200);
      doc.line(0, 148.5, 210, 148.5);
      doc.setLineDash([], 0);
      doc.setDrawColor(0, 0, 0);
    }
  }

  // Turkish char transliteration for PDF (jsPDF default font limitation)
  function transliterate(text) {
    if (!text) return '';
    var map = {
      'ğ': 'g', 'Ğ': 'G',
      'ş': 's', 'Ş': 'S',
      'ı': 'i', 'İ': 'I',
      'ç': 'c', 'Ç': 'C',
      'ö': 'o', 'Ö': 'O',
      'ü': 'u', 'Ü': 'U'
    };
    return text.replace(/[ğĞşŞıİçÇöÖüÜ]/g, function (ch) {
      return map[ch] || ch;
    });
  }

  // Format YYYY-MM-DD to DD.MM.YYYY
  function formatDateTR(dateStr) {
    if (!dateStr) return '';
    var parts = dateStr.split('-');
    if (parts.length === 3) {
      return parts[2] + '.' + parts[1] + '.' + parts[0];
    }
    return dateStr;
  }

  // ===== ESCAPE HTML =====
  function escapeHtml(text) {
    var div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // ===== EVENT LISTENERS =====
  function setupEventListeners() {
    // Student file input
    document.getElementById('student-file-input').addEventListener('change', function (e) {
      if (e.target.files.length > 0) handleStudentUpload(e.target.files[0]);
    });

    // Reload students
    document.getElementById('btn-reload-students').addEventListener('click', function () {
      state.students = [];
      DOM.studentPreview.classList.add('hidden');
      DOM.studentUploadArea.classList.remove('hidden');
      document.getElementById('student-file-input').value = '';
      updateUI();
      showToast('Öğrenci listesi temizlendi. Yeni dosya yükleyebilirsiniz.', 'info');
    });

    // Nav: Step 1 → 2
    document.getElementById('btn-to-step-2').addEventListener('click', function () {
      if (state.students.length === 0) {
        showToast('Lütfen önce öğrenci listesi yükleyin.', 'warning');
        return;
      }
      setStep(2);
    });

    // Nav: Step 2 → 1
    document.getElementById('btn-back-to-step-1').addEventListener('click', function () { setStep(1); });

    // Room file input
    document.getElementById('room-file-input').addEventListener('change', function (e) {
      if (e.target.files.length > 0) handleRoomUpload(e.target.files[0]);
    });

    // Manual room add
    document.getElementById('btn-add-room').addEventListener('click', function () {
      var nameInput = document.getElementById('manual-room-name');
      var capInput = document.getElementById('manual-room-capacity');
      var priorityInput = document.getElementById('manual-room-priority');
      var btnAdd = document.getElementById('btn-add-room');

      var name = nameInput.value.trim();
      var cap = parseInt(capInput.value);
      var priority = parseInt(priorityInput.value) || 999;

      if (!name || isNaN(cap) || cap <= 0) {
        showToast('Lütfen geçerli bir salon adı ve kapasite giriniz.', 'warning');
        return;
      }

      // Check duplicate priority
      if (priority !== 999) {
        var duplicate = state.rooms.find(function (r) { return r.priority === priority && r.id !== state.editingRoomId; });
        if (duplicate) {
          if (!confirm(priority + ' öncelik sırası zaten "' + duplicate.name + '" salonunda kullanılıyor. Devam edilsin mi?')) {
            return;
          }
        }
      }

      if (state.editingRoomId) {
        var roomIndex = state.rooms.findIndex(function (r) { return r.id === state.editingRoomId; });
        if (roomIndex !== -1) {
          state.rooms[roomIndex].name = name;
          state.rooms[roomIndex].capacity = cap;
          state.rooms[roomIndex].priority = priority;
        }
        cancelEditMode();
        showToast('Salon güncellendi.', 'success');
      } else {
        state.rooms.push({
          id: Date.now(),
          name: name,
          capacity: cap,
          priority: priority
        });
        showToast('"' + name + '" salonu eklendi.', 'success');
      }

      state.rooms.sort(function (a, b) { return a.priority - b.priority; });
      nameInput.value = '';
      capInput.value = '';
      priorityInput.value = '';
      updateUI();
      renderRoomTable();
    });

    // Export rooms
    document.getElementById('btn-export-rooms').addEventListener('click', function () {
      if (state.rooms.length === 0) {
        showToast('Dışa aktarılacak salon yok.', 'warning');
        return;
      }
      exportRoomsToExcel(state.rooms);
    });

    // Nav: Start distribution
    document.getElementById('btn-to-step-3').addEventListener('click', function () {
      if (state.rooms.length === 0) {
        showToast('Lütfen en az bir salon tanımlayın.', 'warning');
        return;
      }
      state.rooms.sort(function (a, b) { return a.priority - b.priority; });
      distributeStudents();
      setStep(3);
    });

    // Nav: Step 3 → 2
    document.getElementById('btn-back-to-step-2').addEventListener('click', function () { setStep(2); });

    // Restart
    document.getElementById('btn-restart').addEventListener('click', function () {
      if (confirm('Tüm veriler silinecek ve başa dönülecek. Onaylıyor musunuz?')) {
        location.reload();
      }
    });

    // Export Excel
    document.getElementById('btn-export-excel').addEventListener('click', function () {
      if (!state.distribution) return;
      exportToExcel(state.distribution, state.rooms);
    });

    // Export PDF
    document.getElementById('btn-export-pdf').addEventListener('click', function () {
      if (!state.distribution) return;
      var btn = document.getElementById('btn-export-pdf');
      var originalHTML = btn.innerHTML;
      btn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v4m0 12v4M4.93 4.93l2.83 2.83m8.48 8.48l2.83 2.83M2 12h4m12 0h4M4.93 19.07l2.83-2.83m8.48-8.48l2.83-2.83"/></svg> Hazırlanıyor...';
      btn.disabled = true;

      try {
        generatePDF(state.distribution, state.rooms);
      } catch (e) {
        console.error(e);
        showToast('PDF oluşturulurken hata oluştu.', 'error');
      } finally {
        btn.innerHTML = originalHTML;
        btn.disabled = false;
      }
    });

    // Nav: Step 3 → 4
    document.getElementById('btn-to-step-4').addEventListener('click', function () {
      setStep(4);
    });

    // Nav: Step 4 → 3
    document.getElementById('btn-back-to-step-3').addEventListener('click', function () {
      setStep(3);
    });

    // Nav: Step 4 → 5
    document.getElementById('btn-to-step-5').addEventListener('click', function () {
      setStep(5);
    });

    // Nav: Step 5 → 4
    document.getElementById('btn-back-to-step-4').addEventListener('click', function () {
      setStep(4);
    });

    // Step 5: Optical Form Selection
    var optRadios = document.getElementsByName('optical-type');
    for (var i = 0; i < optRadios.length; i++) {
      optRadios[i].addEventListener('change', function (e) {
        state.opticalFormType = e.target.value;

        // Update UI selection
        document.getElementById('optical-type-lgs').classList.remove('selected');
        document.getElementById('optical-type-tyt').classList.remove('selected');

        if (state.opticalFormType === 'lgs') {
          document.getElementById('optical-type-lgs').classList.add('selected');
          document.getElementById('optical-info-bar').innerHTML = '<span class="text-primary"><strong>LGS</strong> seçildi. Öğrenci bilgileri LGS formatında kodlanacak.</span>';
        } else if (state.opticalFormType === 'tyt') {
          document.getElementById('optical-type-tyt').classList.add('selected');
          document.getElementById('optical-info-bar').innerHTML = '<span class="text-primary"><strong>TYT / AYT</strong> seçildi. Öğrenci bilgileri YKS formatında kodlanacak.</span>';
        }

        document.getElementById('btn-generate-optical').disabled = false;
      });
    }

    // Step 5: Generate Optical Forms
    document.getElementById('btn-generate-optical').addEventListener('click', function () {
      if (!state.opticalFormType) {
        showToast('Lütfen optik form türü seçiniz.', 'warning');
        return;
      }

      var btn = document.getElementById('btn-generate-optical');
      var originalHTML = btn.innerHTML;
      btn.innerHTML = '⏳ Oluşturuluyor...';
      btn.disabled = true;

      setTimeout(function () {
        try {
          generateOpticalFormsPDF();
        } catch (e) {
          console.error(e);
          showToast('Hata: ' + e.message, 'error');
        } finally {
          btn.innerHTML = originalHTML;
          btn.disabled = false;
        }
      }, 50);
    });

    // Exam Info Template Download
    document.getElementById('download-exam-template').addEventListener('click', function (e) {
      e.preventDefault();
      downloadExamInfoTemplate();
    });

    // Exam File Input
    document.getElementById('exam-file-input').addEventListener('change', function (e) {
      if (e.target.files.length > 0) handleExamInfoUpload(e.target.files[0]);
    });

    // Manual Exam Inputs
    ['exam-name', 'exam-date', 'exam-time', 'exam-institution', 'exam-location'].forEach(function (id) {
      document.getElementById(id).addEventListener('input', function (e) {
        var key = id.replace('exam-', '');
        state.examInfo[key] = e.target.value;

        // Show preview if user types manually
        if (state.examInfo.name || state.examInfo.date || state.examInfo.location) {
          DOM.examUploadArea.classList.add('hidden');
          DOM.examPreview.classList.remove('hidden');
        }

        var previewEl = document.getElementById('exam-preview-' + key);
        if (previewEl) previewEl.textContent = state.examInfo[key] || '—';
      });
    });

    // Generate Individual Docs PDF
    document.getElementById('btn-generate-entry-docs').addEventListener('click', function () {
      var btn = document.getElementById('btn-generate-entry-docs');
      var originalHTML = btn.innerHTML;
      btn.innerHTML = '⏳ Oluşturuluyor...';
      btn.disabled = true;

      // Use timeout to allow UI to update
      setTimeout(function () {
        try {
          generateEntryDocumentsPDF();
        } catch (e) {
          console.error(e);
          showToast('Hata: ' + (e.message || e), 'error');
        } finally {
          btn.innerHTML = originalHTML;
          btn.disabled = false;
        }
      }, 50);
    });

    // Enter key on room inputs
    ['manual-room-name', 'manual-room-capacity', 'manual-room-priority'].forEach(function (id) {
      document.getElementById(id).addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          document.getElementById('btn-add-room').click();
        }
      });
    });

    // Project Export
    document.getElementById('btn-export-project').addEventListener('click', function () {
      exportProject();
    });

    // Project Import
    document.getElementById('btn-import-project').addEventListener('click', function () {
      document.getElementById('project-file-input').click();
    });

    document.getElementById('project-file-input').addEventListener('change', function (e) {
      if (e.target.files.length > 0) {
        importProject(e.target.files[0]);
      }
    });

  } // End setupEventListeners

  function exportProject() {
    var data = {
      version: '1.0',
      state: {
        step: state.step,
        students: state.students,
        rooms: state.rooms,
        distribution: state.distribution,
        examInfo: state.examInfo,
        opticalFormType: state.opticalFormType
      }
    };

    var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    var timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    a.href = url;
    a.download = 'sinav_ayarlari_' + timestamp + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Ayar dosyası indirildi.', 'success');
  }

  function importProject(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = JSON.parse(e.target.result);
        if (!data || !data.state) throw new Error('Geçersiz ayar dosyası.');

        // Restore state
        var s = data.state;
        state.students = s.students || [];
        state.rooms = s.rooms || [];
        state.distribution = s.distribution || null;
        state.examInfo = s.examInfo || { name: '', date: '', time: '', institution: '', location: '' };
        state.opticalFormType = s.opticalFormType || null;
        state.step = s.step || 1;
        state.studentPage = 1;
        state.roomPage = 1;

        // Sync UI
        updateUI();
        if (state.students.length > 0) {
          DOM.studentPreview.classList.remove('hidden');
          DOM.studentUploadArea.classList.add('hidden');
          renderStudentTable();
        } else {
          DOM.studentPreview.classList.add('hidden');
          DOM.studentUploadArea.classList.remove('hidden');
        }

        renderRoomTable();
        updateExamInfoUI();

        if (state.distribution) {
          renderDistributionResults();
        }

        // Handle optical form selection sync if needed
        if (state.opticalFormType) {
          var radio = document.querySelector('input[name="optical-type"][value="' + state.opticalFormType + '"]');
          if (radio) {
            radio.checked = true;
            // Trigger the change event logic manually or simulate event
            var event = new Event('change');
            radio.dispatchEvent(event);
          }
        }

        showToast('Ayarlar başarıyla geri yüklendi.', 'success');
        document.getElementById('project-file-input').value = '';

      } catch (err) {
        console.error(err);
        showToast('Ayar yükleme hatası: ' + err.message, 'error');
      }
    };
    reader.readAsText(file);
  }



  // ==========================================
  // OPTICAL FORM GENERATION
  // ==========================================

  function generateOpticalFormsPDF() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error('PDF kütüphanesi yüklenemedi. Sayfayı yenileyip tekrar deneyin.');
    }

    if (!state.opticalFormType) {
      throw new Error('Lütfen bir optik form türü seçiniz.');
    }

    var jsPDF = window.jspdf.jsPDF;
    var doc = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });

    // Add Fonts
    if (window.fontRobotoRegular && window.fontRobotoBold) {
      doc.addFileToVFS('Roboto-Regular.ttf', window.fontRobotoRegular);
      doc.addFileToVFS('Roboto-Bold.ttf', window.fontRobotoBold);
      doc.addFont('Roboto-Regular.ttf', 'Roboto', 'normal');
      doc.addFont('Roboto-Bold.ttf', 'Roboto', 'bold');
      doc.setFont('Roboto', 'normal');
    }

    var sortedRooms = state.rooms.slice().sort(function (a, b) { return (a.priority || 999) - (b.priority || 999); });
    var docIndex = 0;

    sortedRooms.forEach(function (room) {
      var students = state.distribution.results[room.id] || [];
      students.forEach(function (student) {
        if (docIndex > 0) {
          doc.addPage();
        }
        drawOpticalForm(doc, student, room, state.opticalFormType);
        docIndex++;
      });
    });

    var fileName = 'optik_formlar_' + state.opticalFormType + '.pdf';
    doc.save(fileName);
    showToast(docIndex + ' adet optik form oluşturuldu.', 'success');
  }

  function drawAlignmentMarks(doc, pageHeight) {
    doc.setFillColor(0, 0, 0); // Black

    var startY = 10;
    var endY = pageHeight - 10;

    // Sadece yatay çizgiler (Horizontal timing marks)
    var markX = 4;
    var markWidth = 5; // Biraz daha geniş yaptım ki net görünsün
    var markHeight = 1.5;
    var pitch = 5;

    for (var y = startY; y <= endY - markHeight; y += pitch) {
      doc.rect(markX, y, markWidth, markHeight, 'F');
    }
  }

  function drawOpticalForm(doc, student, room, type) {
    var pageWidth = 210;
    var pageHeight = 297;
    // Form içeriği sola çok yakın olduğu için yatay ayar çizgileriyle çakışıyor, 
    // margin artırılarak form hafifçe sağa kaydırıldı (8 -> 12).
    var margin = 12;

    // Pink/Magenta theme color matching reference
    var pinkR = 220, pinkG = 50, pinkB = 120;

    // Draw alignment/timing marks on the left edge
    drawAlignmentMarks(doc, pageHeight);

    if (type === 'lgs') {
      drawLGSOpticalForm(doc, student, room, pageWidth, pageHeight, margin, pinkR, pinkG, pinkB);
    } else {
      drawTYTOpticalForm(doc, student, room, pageWidth, pageHeight, margin);
    }
  }

  function drawLGSOpticalForm(doc, student, room, pageWidth, pageHeight, margin, pR, pG, pB) {
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    // ============ LAYOUT CONSTANTS ============
    var leftX = margin;
    var rightX = pageWidth / 2 + 10;

    // ============ LEFT SIDE: QR CODE ============
    var infoX = rightX - 25; // Pre-calculated to find the left area width
    var leftAreaW = infoX - margin;

    // Calculate the height of the right-side header + info box + dikkat box + kitapçık türü to find the available vertical space
    // Header(8) + gap(2) + InfoBox(28) + gap(2) + Dikkat(9) + gap(4) + Kitapçık(8) = 61
    var leftAreaH = 61;

    // We want the QR code to fill a good portion of the space but remain centered
    var qrSize = 36;

    // Center horizontally and vertically within the left area
    var qrX = margin + (leftAreaW - qrSize) / 2;
    var qrY = margin + (leftAreaH - qrSize) / 2;

    var qrText = (student.name || '') + ' - ' + (student.id || '') + ' - ' + (student.tc || '') + ' - ' + (student.classRef || '');
    var qrData = generateQRCodeDataURL(qrText);
    if (qrData) {
      doc.addImage(qrData, 'PNG', qrX, qrY, qrSize, qrSize);
    }

    // ============ RIGHT SIDE: STUDENT INFO HEADER & FIELDS ============
    var infoWidth = pageWidth - margin - infoX;
    var infoY = margin;

    // 1. Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(infoX, infoY, infoWidth, 8, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10);
    setFont('bold');
    doc.text('İLKOKUL & ORTAOKUL CEVAP KAĞIDI', infoX + infoWidth / 2, infoY + 5.5, { align: 'center' });
    doc.setTextColor(0);

    // 2. Bordered Box below header
    var boxY = infoY + 8 + 2; // Added 2 unit gap
    var boxH = 28;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(infoX, boxY, infoWidth, boxH);

    // 3. Info Fields
    var infoFields = [
      { label: 'Adı Soyadı', value: (student.name || '').toLocaleUpperCase('tr-TR') },
      { label: 'Öğrenci Numarası', value: String(student.id || '') },
      { label: 'TC Kimlik No', value: String(student.tc || '') },
      { label: 'Sınıfı', value: String(student.classRef || '') }
    ];

    var startTextY = boxY + 5.5; // Start earlier to fit 4 fields in 28 units height
    var lineGap = 6;             // Tighter gap

    var maxLabelW = 0;
    doc.setFontSize(8);
    setFont('bold');
    infoFields.forEach(function (f) {
      var w = doc.getTextWidth(f.label + ' :');
      if (w > maxLabelW) maxLabelW = w;
    });

    var labelsEndX = infoX + maxLabelW + 6; // Where labels end
    var valStartX = labelsEndX + 2;         // Where lines start

    infoFields.forEach(function (f, i) {
      var ty = startTextY + i * lineGap;

      // Label (Pink) - Right Aligned
      doc.setFontSize(8);
      setFont('bold');
      doc.setTextColor(pR, pG, pB);
      var lw = doc.getTextWidth(f.label + ' :');
      doc.text(f.label + ' :', labelsEndX - lw, ty);

      // Dotted Line
      var lineEndX = infoX + infoWidth - 3;
      doc.setDrawColor(180, 180, 180); // Light grey dots
      doc.setLineWidth(0.2);
      doc.setLineDash([0.5, 0.8], 0);
      doc.line(valStartX, ty + 1, lineEndX, ty + 1);
      doc.setLineDash([], 0);

      // Value
      doc.setTextColor(0);
      setFont('normal');
      doc.setFontSize(8);
      doc.text(f.value, valStartX + 2, ty - 0.5);
    });

    // ============ DİKKAT SECTION ============
    var dikkatY = boxY + boxH + 2;
    var dikkatH = 9;
    var gapBetweenBoxes = 2;

    // 1. DİKKAT Box (Pink)
    var dikkatW = 16;
    doc.setFillColor(pR, pG, pB);
    doc.rect(infoX, dikkatY, dikkatW, dikkatH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(7);
    setFont('bold');
    doc.text('DİKKAT', infoX + dikkatW / 2, dikkatY + (dikkatH / 2) + 1, { align: 'center' });
    doc.setTextColor(0);

    // 2. Yanlış kodlama Box
    var yanlisX = infoX + dikkatW + gapBetweenBoxes;
    var yanlisW = 38;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(yanlisX, dikkatY, yanlisW, dikkatH);

    doc.setTextColor(pR, pG, pB);
    doc.setFontSize(5.5);
    setFont('bold');
    doc.text('Yanlış kodlama', yanlisX + yanlisW / 2, dikkatY + 3, { align: 'center' });
    doc.setTextColor(0);

    var startCircX = yanlisX + 3.5;
    var circY = dikkatY + 6;
    var gap = 5.5;

    doc.setDrawColor(0);
    doc.setFillColor(0);
    doc.setTextColor(0);

    // 1. Dot in center
    doc.setLineWidth(0.3);
    doc.circle(startCircX, circY, 1.8, 'S');
    doc.circle(startCircX, circY, 0.5, 'F');

    // 2. Vertical Oval / Bean (Simulated)
    doc.ellipse(startCircX + gap, circY, 1.0, 1.8, 'F');

    // 3. Tick
    doc.circle(startCircX + gap * 2, circY, 1.8, 'S');
    doc.setFontSize(5);
    doc.text('✔', startCircX + gap * 2 - 1, circY + 1.2);

    // 4. Cross
    doc.circle(startCircX + gap * 3, circY, 1.8, 'S');
    doc.text('X', startCircX + gap * 3 - 1, circY + 1.2);

    // 5. Scribble
    doc.circle(startCircX + gap * 4, circY, 1.8, 'S');
    doc.line(startCircX + gap * 4 - 1, circY, startCircX + gap * 4 + 1, circY);

    // 6. Dash
    doc.circle(startCircX + gap * 5, circY, 1.8, 'S');
    doc.text('-', startCircX + gap * 5 - 0.5, circY + 1);

    // 3. Doğru kodlama Box
    var dogruX = yanlisX + yanlisW;
    var dogruW = 22;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(dogruX, dikkatY, dogruW, dikkatH);

    doc.setTextColor(pR, pG, pB);
    doc.text('Doğru kodlama', dogruX + dogruW / 2, dikkatY + 3, { align: 'center' });
    doc.setTextColor(0);

    // Correct bubble
    doc.setDrawColor(0);
    doc.setFillColor(0);
    doc.circle(dogruX + dogruW / 2, circY, 1.8, 'F');

    // 4. Text on Right
    var textX = dogruX + dogruW + 2;
    var remainingW = (infoX + infoWidth) - textX;
    doc.setTextColor(0);
    doc.setFontSize(5.5);
    setFont('bold');

    doc.text('Kodlamalarınızı lütfen yumuşak', textX + remainingW / 2, dikkatY + 3, { align: 'center' });
    doc.text('kurşun kalem ile yapınız.', textX + remainingW / 2, dikkatY + 6, { align: 'center' });


    // ============ KİTAPÇIK TÜRÜ ============
    var kitapcikY = dikkatY + dikkatH + 4; // Start below DİKKAT
    var row1H = 8;
    var titleW = 24;
    var boxGap = 4;
    var optionsW = 28;
    var gap2 = 14;
    var textX2 = infoX + titleW + boxGap + optionsW + gap2 - 32; // Exact shifted left

    doc.setFillColor(pR, pG, pB);
    doc.rect(infoX, kitapcikY, titleW, row1H, 'F');
    doc.setTextColor(255);
    doc.setFontSize(7);
    setFont('bold');
    doc.text('KİTAPÇIK TÜRÜ', infoX + titleW / 2, kitapcikY + 5.3, { align: 'center' });

    var optionsX = infoX + titleW + boxGap;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(optionsX, kitapcikY, optionsW, row1H);

    var ktBubs = ['A', 'B', 'C', 'D'];
    var bubGap = optionsW / (ktBubs.length + 1);

    for (var k = 0; k < ktBubs.length; k++) {
      var bx = optionsX + bubGap * (k + 1);
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.3);
      doc.circle(bx, kitapcikY + 4, 2, 'S');
      doc.setTextColor(pR, pG, pB);
      doc.setFontSize(5.5);
      setFont('bold');
      doc.text(ktBubs[k], bx, kitapcikY + 4.65, { align: 'center' });
    }

    var textRemainingW = (pageWidth - margin) - textX2;
    doc.setFontSize(6);
    setFont('bold');
    doc.setTextColor(0);
    doc.text('Kitapçık Türünü', textX2 + textRemainingW / 2, kitapcikY + 3, { align: 'center' });
    doc.text('Kodlamayı Unutmayınız.', textX2 + textRemainingW / 2, kitapcikY + 6, { align: 'center' });

    // Reset colors
    doc.setTextColor(0);
    doc.setDrawColor(0);

    // ============ BOTTOM BLOCK (SÖZEL / SAYISAL) ============
    var bottomY = kitapcikY + row1H + 14;
    drawLGSLayout(doc, bottomY, pR, pG, pB);

    // ============ FOOTER ============
  }

  function drawPinkBubbleGrid(doc, value, x, y, cols, gapX, gapY, bubbleR, pR, pG, pB) {
    var rows = 10; // 0-9

    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    // Draw digit values at top
    for (var c = 0; c < cols; c++) {
      doc.setFontSize(6.5);
      setFont('bold');
      doc.text(value[c], x + c * gapX + gapX / 2, y, { align: 'center' });
    }

    // Draw grid
    var gridStartY = y + 2;
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var bx = x + c * gapX + gapX / 2;
        var by = gridStartY + r * gapY;
        var digit = r.toString();
        var isFilled = (value[c] === digit);

        if (isFilled) {
          doc.setFillColor(pR, pG, pB);
          doc.circle(bx, by, bubbleR, 'F');
          doc.setTextColor(255, 255, 255);
          doc.setFontSize(4.5);
          var tw = doc.getTextWidth(digit);
          doc.text(digit, bx - tw / 2, by + 0.8);
          doc.setTextColor(0);
        } else {
          doc.setDrawColor(pR, pG, pB);
          doc.setLineWidth(0.2);
          doc.circle(bx, by, bubbleR, 'S');
          doc.setFontSize(4.5);
          doc.setTextColor(pR, pG, pB);
          var tw2 = doc.getTextWidth(digit);
          doc.text(digit, bx - tw2 / 2, by + 0.8);
          doc.setTextColor(0);
        }
      }
    }
  }

  function drawLGSLayout(doc, startY, pR, pG, pB) {
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };
    // Sola eklediğimiz ayar çizgilerine(timing marks) çarpmaması için 
    // sol margin değerini 14 yapıyoruz. Sağ tarafı ise sağa kaydırmak için daraltıyoruz.
    var leftMargin = 14;
    var rightMargin = 6;
    var pageWidth = 210;

    // İçerik genişliğini asimetrik marginlere göre hesaplıyoruz
    var contentWidth = pageWidth - leftMargin - rightMargin; // 190

    // Split: SÖZEL (Left 4 cols) | SAYISAL (Right 2 cols)
    var colW = contentWidth / 6.2;
    var sozelW = colW * 4;
    var sayisalW = colW * 2;
    var gap = contentWidth - sozelW - sayisalW;

    var sayisalX = leftMargin + sozelW + gap;

    // Headers
    var headerH = 7;
    // SÖZEL Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(leftMargin, startY, sozelW, headerH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(10);
    setFont('bold');
    doc.text('SÖZEL BÖLÜM', leftMargin + sozelW / 2, startY + 5, { align: 'center' });

    // SAYISAL Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(sayisalX, startY, sayisalW, headerH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(10);
    setFont('bold');
    doc.text('SAYISAL BÖLÜM', sayisalX + sayisalW / 2, startY + 5, { align: 'center' });
    doc.setTextColor(0);

    // Columns
    var colY = startY + headerH + 4; // Increased gap to 4
    var subHeadH = 10;
    var gapBelowSubHead = 3.0; // Increased gap to 3.0
    var borderY = colY + subHeadH + gapBelowSubHead;
    var gridTopPadding = 1.0;
    var rowGap = 7.2; // Further increased row gap for LGS as requested

    // SÖZEL Columns
    var sozelCols = [
      { title: 'TÜRKÇE', q: 20 },
      { title: 'SOSYAL BİLGİLER\nİNKILAP TARİHİ VE\nATATÜRKÇÜLÜK', q: 10 },
      { title: 'DİN KÜLTÜRÜ\nVE\nAHLAK BİLGİSİ', q: 10 },
      { title: 'İNGİLİZCE', q: 10 }
    ];
    var sColGap = 2.5; // Sütunlar arası boşluk
    var sColW = (sozelW - 3 * sColGap) / 4;

    for (var i = 0; i < 4; i++) {
      var cx = leftMargin + i * (sColW + sColGap);

      // SubHeader
      doc.setFillColor(pR, pG, pB);
      doc.rect(cx, colY, sColW, subHeadH, 'F');
      doc.setTextColor(255);
      doc.setFontSize(7);

      var lines = sozelCols[i].title.split('\n');
      var totalTextH = lines.length * 2.5;
      var startTextY = colY + (subHeadH - totalTextH) / 2 + 1.8;
      for (var li = 0; li < lines.length; li++) {
        doc.text(lines[li], cx + sColW / 2, startTextY + li * 2.5, { align: 'center' });
      }

      // Answers
      // Use the new rowGap
      drawPinkAnswerColumn(doc, sozelCols[i].q, cx, borderY + rowGap / 2 + gridTopPadding, sColW, pR, pG, pB, rowGap);

      // Individual Column Border
      var colGridH = sozelCols[i].q * rowGap + 2.0;
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.3);
      doc.rect(cx, borderY, sColW, colGridH);
    }

    // SAYISAL Columns
    var sayisalColsVals = [
      { title: 'MATEMATİK', q: 20 },
      { title: 'FEN BİLİMLERİ', q: 20 }
    ];
    var mColGap = 2.5; // Sütunlar arası boşluk
    var mColW = (sayisalW - 1 * mColGap) / 2;

    for (var j = 0; j < 2; j++) {
      var cx = sayisalX + j * (mColW + mColGap);

      // SubHeader
      doc.setFillColor(pR, pG, pB);
      doc.rect(cx, colY, mColW, subHeadH, 'F');
      doc.setTextColor(255);
      doc.setFontSize(7);
      doc.text(sayisalColsVals[j].title, cx + mColW / 2, colY + subHeadH / 2 + 1.2, { align: 'center' });

      // Answers
      drawPinkAnswerColumn(doc, sayisalColsVals[j].q, cx, borderY + rowGap / 2 + gridTopPadding, mColW, pR, pG, pB, rowGap);

      // Individual Column Border
      var colGridH = sayisalColsVals[j].q * rowGap + 2.0;
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.3);
      doc.rect(cx, borderY, mColW, colGridH);
    }

  }

  function drawPinkAnswerColumn(doc, count, x, y, availWidth, pR, pG, pB, rGap) {
    if (!rGap) rGap = 4.5;
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    var opts = ['A', 'B', 'C', 'D'];

    // Add some padding inside the column border
    var colLeftPadding = 2.5; // Space between left border and numbers

    var leftSpace = 6.5; // Reduced from 8.0: Space between numbers and first bubble
    var rightPadding = 3.8; // Increased from 3.0: Space between last bubble and right border
    var activeWidth = availWidth - colLeftPadding - leftSpace - rightPadding;

    // Distribute bubbles across activeWidth
    var step = activeWidth / (opts.length - 1);
    var startX = x + colLeftPadding + leftSpace;

    for (var i = 1; i <= count; i++) {
      var rowY = y + (i - 1) * rGap;

      // Zebra striping for even rows to improve readability
      if (i % 2 === 0) {
        doc.setFillColor(250, 230, 240); // Very light pink
        doc.rect(x, rowY - rGap / 2, availWidth, rGap, 'F');
      }

      // Question number
      doc.setFontSize(6.5);
      setFont('bold');
      doc.setTextColor(pR, pG, pB); // Pink like bubbles
      doc.text(String(i), x + colLeftPadding, rowY + 0.1, { baseline: 'middle' });

      // Bubbles
      for (var o = 0; o < opts.length; o++) {
        var obx = startX + o * step;
        doc.setDrawColor(pR, pG, pB);
        doc.setLineWidth(0.3); // Standardized thickness
        doc.circle(obx, rowY, 1.85, 'S'); // Make bubbles slightly smaller for more gap (radius 1.85)

        doc.setFontSize(6); // Bigger text inside bubbles
        setFont('bold');
        doc.setTextColor(pR, pG, pB);
        doc.text(opts[o], obx, rowY + 0.1, { align: 'center', baseline: 'middle' }); // Automatically center
      }
    }
  }

  // TYT/AYT form matching "LİSE GRUBU CEVAP KAĞIDI" reference
  function drawTYTOpticalForm(doc, student, room, pageWidth, pageHeight, margin) {
    var pR = 220, pG = 50, pB = 120;
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    var leftX = margin;
    var rightX = pageWidth / 2 + 10;

    // ============ LEFT SIDE: QR CODE ============
    var infoX = rightX - 25; // Pre-calculated to find the left area width
    var leftAreaW = infoX - margin;

    // Calculate the height of the right-side header + info box + dikkat box to find the available vertical space
    // Header(8) + gap(2) + InfoBox(28) + gap(2) + Dikkat(9) = 49
    var leftAreaH = 49;

    // We want the QR code to fill a good portion of the space but remain centered
    var qrSize = 36;

    // Center horizontally and vertically within the left area
    var qrX = margin + (leftAreaW - qrSize) / 2;
    var qrY = margin + (leftAreaH - qrSize) / 2;

    var qrText = (student.name || '') + ' - ' + (student.id || '') + ' - ' + (student.tc || '') + ' - ' + (student.classRef || '');
    var qrData = generateQRCodeDataURL(qrText);
    if (qrData) {
      doc.addImage(qrData, 'PNG', qrX, qrY, qrSize, qrSize);
    }

    // ============ RIGHT SIDE: STUDENT INFO HEADER & FIELDS ============
    // Align with image: "LİSE GRUBU CEVAP KAĞIDI" pink header, then bordered box below.

    var infoX = rightX - 25; // Shift left a bit to give more space
    var infoWidth = pageWidth - margin - infoX;
    var infoY = margin;

    // 1. Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(infoX, infoY, infoWidth, 8, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(11);
    setFont('bold');
    doc.text('LİSE GRUBU CEVAP KAĞIDI', infoX + infoWidth / 2, infoY + 5.5, { align: 'center' });
    doc.setTextColor(0);

    // 2. Bordered Box below header
    var boxY = infoY + 8 + 4; // Increased from 2 to 4
    var boxH = 28;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.4);
    doc.rect(infoX, boxY, infoWidth, boxH);

    // 3. Info Fields (Adı Soyadı, Öğrenci Numarası, TC Kimlik No, Sınıfı as requested)
    var infoFields = [
      { label: 'Adı Soyadı', value: (student.name || '').toLocaleUpperCase('tr-TR') },
      { label: 'Öğrenci Numarası', value: String(student.id || '') },
      { label: 'TC Kimlik No', value: String(student.tc || '') },
      { label: 'Sınıfı', value: String(student.classRef || '') }
    ];

    var startTextY = boxY + 5.5; // Start earlier to fit 4 fields in 28 units height
    var lineGap = 6;             // Tighter gap

    // Pre-calculate fixed start X for values so they align vertically
    var maxLabelW = 0;
    doc.setFontSize(8);
    setFont('bold');
    infoFields.forEach(function (f) {
      var w = doc.getTextWidth(f.label + ' :');
      if (w > maxLabelW) maxLabelW = w;
    });

    var labelsEndX = infoX + maxLabelW + 6; // Where labels end
    var valStartX = labelsEndX + 2;         // Where lines start

    infoFields.forEach(function (f, i) {
      var ty = startTextY + i * lineGap;

      // Label (Pink) - Right Aligned
      doc.setFontSize(8);
      setFont('bold');
      doc.setTextColor(pR, pG, pB);
      var lw = doc.getTextWidth(f.label + ' :');
      doc.text(f.label + ' :', labelsEndX - lw, ty);

      // Dotted Line
      var lineEndX = infoX + infoWidth - 3;
      doc.setDrawColor(180, 180, 180); // Light grey dots
      doc.setLineWidth(0.2);
      doc.setLineDash([0.5, 0.8], 0);
      doc.line(valStartX, ty + 1, lineEndX, ty + 1);
      doc.setLineDash([], 0);

      // Value
      doc.setTextColor(0);
      setFont('normal'); // Value normal weight? Or bold?
      doc.setFontSize(8);
      doc.text(f.value, valStartX + 2, ty - 0.5);
    });

    // ============ NUMARANIZ BOX REMOVED AS REQUESTED ============


    // ============ DİKKAT SECTION ============
    // Visual: Pink "DİKKAT" box left, then white box with "Yanlış kodlama" examples, then "Doğru kodlama" example, then text.
    var dikkatY = boxY + boxH + 6; // Increased from 2 to 6 unit gap below info box
    var dikkatH = 9; // Increased height from 7 to 9
    var gapBetweenBoxes = 2; // Requested gap

    // 1. DİKKAT Box (Pink)
    var dikkatW = 16;
    doc.setFillColor(pR, pG, pB);
    doc.rect(infoX, dikkatY, dikkatW, dikkatH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(7); // Slightly larger
    setFont('bold');
    doc.text('DİKKAT', infoX + dikkatW / 2, dikkatY + (dikkatH / 2) + 1, { align: 'center' }); // Vertically centered
    doc.setTextColor(0);

    // 2. Yanlış kodlama Box REMOVED AS REQUESTED

    // 3. Doğru kodlama Box
    var dogruX = infoX + dikkatW + gapBetweenBoxes;
    var dogruW = 22;
    doc.setDrawColor(pR, pG, pB);
    doc.rect(dogruX, dikkatY, dogruW, dikkatH);

    doc.setTextColor(pR, pG, pB);
    doc.setFontSize(5.5);
    setFont('bold');
    doc.text('Doğru kodlama', dogruX + dogruW / 2, dikkatY + 3, { align: 'center' }); // Adjusted Y
    doc.setTextColor(0);

    // Correct bubble (Black as requested)
    var circY = dikkatY + 6; // Center in height
    doc.setDrawColor(0);
    doc.setFillColor(0);
    doc.circle(dogruX + dogruW / 2, circY, 1.8, 'F');

    // 4. Text on Right (Centered in the remaining right area, Black)
    var textX = dogruX + dogruW + 2; // Shifted slightly left from +3 to +2
    var remainingW = (infoX + infoWidth) - textX; // Calculate remaining width for centering
    doc.setTextColor(0); // Black text requested
    doc.setFontSize(5.5);
    setFont('bold');

    // Guidance text removed as requested


    // ============ KİTAPÇIK TÜRÜ (MOVED NEXT TO DİKKAT) ============
    var ktStartX = dogruX + dogruW + 14; // increased gap between Doğru Kodlama and Kitapçık Türü
    var ktStartY = dikkatY + 0.5; // vertically align with DİKKAT (height 9 vs 8)
    var row1H = 8;
    var titleW = 24;
    var boxGap = 4;
    var optionsW = 28;

    // Left part: Pink box with text "KİTAPÇIK TÜRÜ"
    doc.setFillColor(pR, pG, pB);
    doc.rect(ktStartX, ktStartY, titleW, row1H, 'F');
    doc.setTextColor(255);
    doc.setFontSize(7);
    setFont('bold');
    doc.text('KİTAPÇIK TÜRÜ', ktStartX + titleW / 2, ktStartY + 5.3, { align: 'center' });

    // Middle part: White box with pink border for bubbles
    var optionsX = ktStartX + titleW + boxGap;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(optionsX, ktStartY, optionsW, row1H);

    // Bubbles (A B C D) in pink
    var ktBubs = ['A', 'B', 'C', 'D'];
    var bubGap = optionsW / (ktBubs.length + 1);

    for (var k = 0; k < ktBubs.length; k++) {
      var bx = optionsX + bubGap * (k + 1);
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.3);
      doc.circle(bx, ktStartY + 4, 2, 'S');
      doc.setTextColor(pR, pG, pB);
      doc.setFontSize(5.5);
      setFont('bold');
      doc.text(ktBubs[k], bx, ktStartY + 4.65, { align: 'center' });
    }

    doc.setTextColor(0);
    doc.setDrawColor(0);

    // ============ REFINED LAYOUT (BELOW DİKKAT) ============
    var leftStartY = dikkatY + dikkatH + 5;
    // We moved Kitapçık Türü up, so we can shift answer columns up to give more space
    var rightStartY = leftStartY;
    drawRefinedTYTLayout(doc, leftStartY, rightStartY, pR, pG, pB, student);


  }

  function drawRefinedTYTLayout(doc, leftStartY, rightStartY, pR, pG, pB, student) {
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    var margin = 8;
    var pageWidth = 210;
    var contentW = pageWidth - 2 * margin;

    // LEFT BLOCK (Kitapçık, TC, Sınıf, Name) approx 33% to leave a larger gap
    // RIGHT BLOCK (Oturum, Answers) approx 56%
    var leftW = contentW * 0.33;
    var rightW = contentW * 0.56;
    var rightX = margin + contentW - rightW; // Right-align the right block

    // ================= LEFT BLOCK =================
    // Left block items (Kitapçık türü, TC, Sınıf vs) were moved up or removed entirely.

    // ================= RIGHT BLOCK =================

    // The OTURUM and TEPM headers were removed per user request.

    // 2. Answer Columns (4 Cols)
    var colsY = rightStartY; // Start columns at the top of the right block

    // Refined distribution: First column at left edge (margin), last column at right edge (wider content area)
    var leftMargin = 22; // Increased to pull columns closer to center
    var rightMargin = 8;
    var totalAvailableWidth = pageWidth - leftMargin - rightMargin; // 180mm

    // Adjust column width and gap
    var ansColW = 34; // slightly wider columns
    var colGap = (totalAvailableWidth - 4 * ansColW) / 3; // Reduced gap: (180 - 136) / 3 = 14.6mm
    var colsStartX = leftMargin;

    var colTitles = [
      { top: ['TÜRKÇE'], bot: ['TÜRK DİLİ VE EDEBİYATI', 'SOSYAL BİLİMLER 1'] },
      { top: ['SOSYAL', 'BİLİMLER'], bot: ['SOSYAL', 'BİLİMLER 2'] },
      { top: ['TEMEL', 'MATEMATİK'], bot: ['MATEMATİK'] },
      { top: ['FEN', 'BİLİMLERİ'], bot: ['FEN', 'BİLİMLERİ'] }
    ];

    for (var i = 0; i < 4; i++) {
      var cx = colsStartX + i * (ansColW + colGap);

      // We split the header into two taller boxes
      var boxH = 8; // Increased from 5.8 to 8 for more height
      var gapH = 0.4;
      var topY = colsY;
      var botY = colsY + boxH + gapH;

      // Draw Top Box
      doc.setFillColor(pR, pG, pB);
      doc.rect(cx, topY, ansColW, boxH, 'F');

      // Draw Bottom Box
      doc.rect(cx, botY, ansColW, boxH, 'F');

      doc.setTextColor(255);
      doc.setFontSize(7); // Matched with KİTAPÇIK TÜRÜ

      var data = colTitles[i];

      // Helper to draw text vertically centered in a taller box
      var drawCenteredLines = function (lines, boxStartY) {
        if (lines.length === 1) {
          doc.text(lines[0], cx + ansColW / 2, boxStartY + 5.5, { align: 'center' });
        } else if (lines.length === 2) {
          doc.text(lines[0], cx + ansColW / 2, boxStartY + 3.8, { align: 'center' });
          doc.text(lines[1], cx + ansColW / 2, boxStartY + 7.0, { align: 'center' });
        }
      };

      // Draw Top Texts
      drawCenteredLines(data.top, topY);

      // Draw Bottom Texts
      drawCenteredLines(data.bot, botY);

      // 1-40 Answer Grid
      var headerTotalH = boxH * 2 + gapH; // 16.4
      var gapBelowHeader = 1.0; // Reduced white space separating headers from the grid
      var borderY = colsY + headerTotalH + gapBelowHeader;
      var rGap = 4.6; // Reduced from 4.8 to fit 40 rows within A4 height
      var gridY = borderY + rGap / 2 + 1.0; // Added 1.0 units of top padding inside the grid

      for (var q = 1; q <= 40; q++) {
        var qy = gridY + (q - 1) * rGap;

        // Zebra striping for even rows to improve readability
        if (q % 2 === 0) {
          doc.setFillColor(250, 230, 240); // Very light pink
          // doc.rect(x, y, w, h, style)
          doc.rect(cx, qy - rGap / 2, ansColW, rGap, 'F');
        }

        // Q Num
        doc.setTextColor(pR, pG, pB); // Pink like bubbles
        doc.setFontSize(5.5); // Slightly larger
        doc.text(String(q), cx + 1.5, qy + 0.1, { baseline: 'middle' });

        // Bubbles A-E
        var opts = ['A', 'B', 'C', 'D', 'E'];
        // Reduced paddings to bring bubbles closer to number and border
        var optW = (ansColW - 5 - 1.5) / 5;
        for (var o = 0; o < 5; o++) {
          var obx = cx + 5 + o * optW + optW / 2;
          doc.setDrawColor(pR, pG, pB);
          doc.setLineWidth(0.3); // Standardized thickness
          doc.circle(obx, qy, 1.75, 'S'); // Make bubbles bigger (radius from 1.5 to 1.75)
          doc.setTextColor(pR, pG, pB);
          doc.setFontSize(5.5); // Bigger text inside bubbles (from 4.5 to 5.5)
          doc.text(opts[o], obx, qy + 0.1, { align: 'center', baseline: 'middle' }); // Automatically center
        }
      }

      // Draw outer border for the answer grid (distinct from headers)
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.3); // Standardized thickness
      var gridH = 40 * rGap + 2.0; // Accounts for the 1.0 top padding and adds 1.0 bottom padding
      doc.rect(cx, borderY, ansColW, gridH);
    }
  }

  function drawSimpleBubble(doc, x, y, r, txt, filled, pR, pG, pB) {
    if (filled) {
      doc.setFillColor(0);
      doc.circle(x, y, r, 'F');
      doc.setTextColor(255);
    } else {
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.3); // Standardized thickness
      doc.circle(x, y, r, 'S');
      doc.setTextColor(pR, pG, pB);
    }
    doc.setFontSize(4);
    doc.text(txt, x - doc.getTextWidth(txt) / 2, y + r / 3);
  }


  function drawVerticalBubbleGrid(doc, label, value, x, y) {
    var cols = value.length;
    var rows = 10;
    var bubbleSize = 3;
    var gapX = 4;
    var gapY = 4;

    doc.setFontSize(8);
    if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    doc.text(label, x + (cols * gapX) / 2, y + 4, { align: 'center' });

    for (var c = 0; c < cols; c++) {
      doc.text(value[c], x + c * gapX + 1.5, y + 9, { align: 'center' });
    }

    var gridY = y + 12;
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var bx = x + c * gapX;
        var by = gridY + r * gapY;
        var digit = r.toString();
        var isFilled = (value[c] === digit);
        drawBubble(doc, bx + 1.5, by + 1.5, bubbleSize, digit, isFilled);
      }
    }
  }

  function drawAnswerColumn(doc, title, count, x, y) {
    doc.setFontSize(9);
    if (window.fontRobotoRegular) doc.setFont('Roboto', 'bold');
    doc.text(title, x + 15, y - 4, { align: 'center' });

    var bubbleSize = 2.5;
    var gapY = 3.8;
    var opts = ['A', 'B', 'C', 'D'];

    if (state.opticalFormType === 'tyt') {
      opts.push('E');
    }

    var gapX = 4;

    for (var i = 1; i <= count; i++) {
      var rowY = y + (i - 1) * gapY;
      doc.setFontSize(7);
      doc.text(String(i), x, rowY + 1);
      for (var o = 0; o < opts.length; o++) {
        drawBubble(doc, x + 6 + (o * gapX), rowY, bubbleSize, opts[o], false);
      }
    }
  }

  function drawBubble(doc, cx, cy, r, text, filled) {
    if (filled) {
      doc.setFillColor(0);
      doc.circle(cx, cy, r, 'F');
      doc.setTextColor(255);
      doc.setFontSize(6);
      var txtW = doc.getTextWidth(text);
      doc.text(text, cx - (txtW / 2), cy + 1.2);
      doc.setTextColor(0);
    } else {
      doc.setDrawColor(0);
      doc.circle(cx, cy, r, 'S');
      doc.setFontSize(6);
      var txtW = doc.getTextWidth(text);
      doc.text(text, cx - (txtW / 2), cy + 1.2);
    }
  }


  function generateQRCodeDataURL(text) {
    if (typeof QRCode === 'undefined') return null;

    // Create a temporary, isolated container for this specific QR code
    var tempContainer = document.createElement('div');
    tempContainer.style.display = 'none';
    document.body.appendChild(tempContainer);

    var dataURL = null;

    try {
      // Create a simplified ASCII-only version of the text to prevent QRCode.js from failing on Special/Turkish characters
      var safeText = text
        .replace(/ı/g, 'i').replace(/İ/g, 'I')
        .replace(/ğ/g, 'g').replace(/Ğ/g, 'G')
        .replace(/ü/g, 'u').replace(/Ü/g, 'U')
        .replace(/ş/g, 's').replace(/Ş/g, 'S')
        .replace(/ö/g, 'o').replace(/Ö/g, 'O')
        .replace(/ç/g, 'c').replace(/Ç/g, 'C');

      // Ensure it's safe ASCII
      var encodedText = unescape(encodeURIComponent(safeText));

      var qr = new QRCode(tempContainer, {
        text: encodedText,
        width: 256,
        height: 256,
        correctLevel: QRCode.CorrectLevel.M // Lower error correction to fit more data if needed
      });

      // QRCode.js (when using canvas) draws synchronously. 
      var canvas = tempContainer.querySelector('canvas');
      if (canvas) {
        dataURL = canvas.toDataURL("image/png");
      } else {
        // Fallback for img (might be async if it uses a data URI, but usually sync for small data)
        var img = tempContainer.querySelector('img');
        if (img && img.src) {
          dataURL = img.src;
        }
      }
    } catch (e) {
      console.error("QR Code generation failed:", e);
    } finally {
      // Clean up the temporary container
      document.body.removeChild(tempContainer);
    }

    return dataURL;
  }

  document.addEventListener('DOMContentLoaded', init);

})();
