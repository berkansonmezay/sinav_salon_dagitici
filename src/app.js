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

  } // End setupEventListeners



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

  function drawOpticalForm(doc, student, room, type) {
    var pageWidth = 210;
    var pageHeight = 297;
    var margin = 8;

    // Pink/Magenta theme color matching reference
    var pinkR = 220, pinkG = 50, pinkB = 120;

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
    var rightBlockX = 90; // Start of the right block (Name grid etc)
    var rightBlockW = pageWidth - rightBlockX - margin;

    // ============ RIGHT HEADER ============
    // "İLKOKUL & ORTAOKUL CEVAP KAĞIDI"
    doc.setFillColor(pR, pG, pB);
    doc.rect(rightBlockX, margin, rightBlockW, 8, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10);
    setFont('bold');
    doc.text('İLKOKUL & ORTAOKUL CEVAP KAĞIDI', rightBlockX + rightBlockW - 2, margin + 5, { align: 'right' });
    doc.setTextColor(0);

    // ============ LEFT BLOCK (Student Info & Codes) ============
    var row1Y = margin + 12;

    // --- Student Info Box ---
    // Box dimensions
    var infoBoxW = 75;
    var infoBoxH = 35;

    // "GRUP NO" vertical strip attached to left of info box
    var grupNoW = 12;
    doc.setFillColor(pR, pG, pB);
    doc.rect(leftX, row1Y, grupNoW, 8, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(6);
    doc.text('GRUP', leftX + grupNoW / 2, row1Y + 3, { align: 'center' });
    doc.text('NO', leftX + grupNoW / 2, row1Y + 6, { align: 'center' });
    doc.setTextColor(0);

    // Grup No bubbles (Just a visual placeholder 0-9 single col? Or 2 cols? 
    // Image shows "GRUP NO" header, and below it simple vertical line. 
    // Let's standard 0-9 single column for simplicity unless user complained.)
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(leftX, row1Y + 8, grupNoW, infoBoxH - 8);
    // Draw 0-9 bubbles 
    // We can use drawPinkBubbleGrid for a single column 0-9
    // But let's check if we have data. '01' is typical. 
    // The image seems to show 2 columns? "0 0", "1 1". 
    // Let's do 2 columns.
    var grupNoVal = '01'; // Default
    drawPinkBubbleGrid(doc, grupNoVal, leftX + 1, row1Y + 9, 2, 4, 3.2, 1.4, pR, pG, pB);


    // Info Text Box
    var infoX = leftX + grupNoW + 2;
    var infoW = infoBoxW - grupNoW - 2;
    doc.setDrawColor(pR, pG, pB);
    doc.rect(infoX, row1Y, infoW, infoBoxH);

    var fields = [
      { label: 'Soyadı - Adı', value: (student.name || '').toLocaleUpperCase('tr-TR') },
      { label: 'Numarası', value: String(student.id || '') },
      { label: 'Sınıfı / Şubesi', value: String(student.classRef || '') },
      { label: 'Telefon Nu.', value: String(student.phone || '') },
      { label: 'Kurum Adı', value: (state.examInfo.institution || '').toLocaleUpperCase('tr-TR') }
    ];

    for (var fi = 0; fi < fields.length; fi++) {
      var fy = row1Y + 6 + fi * 6.5;
      doc.setFontSize(7);
      setFont('bold');
      doc.setTextColor(pR, pG, pB);
      doc.text(fields[fi].label, infoX + 2, fy);
      doc.setTextColor(0);
      setFont('normal');
      doc.text(':', infoX + 25, fy);
      doc.text(fields[fi].value, infoX + 27, fy);
      // Dotted line
      doc.setDrawColor(200);
      doc.setLineWidth(0.1);
      doc.line(infoX + 27, fy + 1, infoX + infoW - 2, fy + 1);
    }

    // --- Doğru Kodlama / Kitapçık Türü ---
    var row2Y = row1Y + infoBoxH + 3;
    var dkH = 12;

    // Doğru Kodlama
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(leftX, row2Y, 22, dkH);
    doc.setFontSize(5);
    setFont('bold');
    doc.text('Doğru Kodlama', leftX + 11, row2Y + 3, { align: 'center' });
    doc.text('Örneği', leftX + 11, row2Y + 5.5, { align: 'center' });
    // Hand icon? Skip. Just bubble.
    doc.setFillColor(0);
    doc.circle(leftX + 11, row2Y + 9, 1.5, 'F');

    // Kitapçık Türü
    var ktX = leftX + 24;
    var ktW = 50;
    doc.setFontSize(5.5);
    setFont('italic');
    doc.setTextColor(pR, pG, pB);
    doc.text('Kitapçık türünü kodlamayı unutmayınız.', ktX, row2Y - 1); // Text above
    doc.setTextColor(0);
    setFont('normal');

    // Pink Box
    doc.setFillColor(pR, pG, pB);
    doc.rect(ktX, row2Y, ktW, dkH, 'F');
    // White Box inside for bubbles
    doc.setFillColor(255);
    doc.rect(ktX + 22, row2Y + 1, ktW - 24, dkH - 2, 'F');

    doc.setTextColor(255);
    doc.setFontSize(7);
    setFont('bold');
    doc.text('KİTAPÇIK TÜRÜ', ktX + 11, row2Y + 7, { align: 'center' });
    doc.setTextColor(0);

    // Bubbles A B
    var ktBubs = ['A', 'B'];
    for (var k = 0; k < ktBubs.length; k++) {
      var kbx = ktX + 28 + k * 10;
      doc.setDrawColor(0);
      doc.circle(kbx, row2Y + 6, 2, 'S');
      doc.setFontSize(6);
      doc.text(ktBubs[k], kbx - 0.5, row2Y + 7);
    }


    // --- ÖĞRENCİ NO (Left Bottom) ---
    var row3Y = row2Y + dkH + 4;
    var stdNoW = 22;

    // Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(leftX, row3Y, stdNoW, 6, 'F');
    doc.setTextColor(255);
    doc.setFontSize(6);
    doc.text('ÖĞRENCİ NO.', leftX + stdNoW / 2, row3Y + 4, { align: 'center' });
    doc.setTextColor(0);

    // Grid
    doc.setDrawColor(pR, pG, pB);
    doc.rect(leftX, row3Y + 6, stdNoW, 35); // Box around grid

    var studentNo = String(student.id || '').replace(/[^0-9]/g, '');
    studentNo = studentNo.padStart(6, '0').substring(0, 6); // visual only shows 6 cols?
    // Visual shows "0 0 0 0" 4 cols? Or 6?
    // Let's assume 6 for standard.
    // Actually the reference image shows 4 columns of bubbles under ÖĞRENCİ NO?
    // Let's count... 1,2,3... looks like 4-5.
    // Let's stick to 6 to be safe for typical IDs, or 9.
    // But I will fit 6.
    drawPinkBubbleGrid(doc, studentNo, leftX + 1, row3Y + 7, 6, 3.4, 3.2, 1.3, pR, pG, pB);


    // --- T.C. KİMLİK NO / CEP (Right Bottom of Left Block) ---
    var tcX = leftX + 26;
    var tcW = infoBoxW - 26;

    // Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(tcX, row3Y, tcW, 6, 'F');
    doc.setTextColor(255);
    doc.setFontSize(6);
    doc.text('T.C. KİMLİK NO. / CEP TELEFONU', tcX + tcW / 2, row3Y + 4, { align: 'center' });
    doc.setTextColor(0);

    // Grid
    doc.setDrawColor(pR, pG, pB);
    doc.rect(tcX, row3Y + 6, tcW, 35);

    var tcNo = String(student.tc || '').replace(/[^0-9]/g, '');
    tcNo = tcNo.padStart(11, '0').substring(0, 11);
    drawPinkBubbleGrid(doc, tcNo, tcX + 1, row3Y + 7, 11, 4.0, 3.2, 1.3, pR, pG, pB);


    // ============ RIGHT BLOCK (SINIF/ŞUBE + SOYADI-ADI) ============
    // "SOYADI - ADI" Header Strip
    var nameHeaderY = margin + 9; // Below the top header
    doc.setFillColor(pR, pG, pB);
    doc.rect(rightBlockX, nameHeaderY, rightBlockW, 6, 'F');

    // "SINIF ŞUBE" part of header (small left part)
    var sinifSubeW = 12; // Width for SINIF ŞUBE column
    doc.setDrawColor(255);
    doc.setLineWidth(0.5);
    doc.line(rightBlockX + sinifSubeW, nameHeaderY, rightBlockX + sinifSubeW, nameHeaderY + 6);

    doc.setTextColor(255);
    doc.setFontSize(5);
    doc.text('SINIF', rightBlockX + sinifSubeW / 2, nameHeaderY + 2.5, { align: 'center' });
    doc.text('ŞUBE', rightBlockX + sinifSubeW / 2, nameHeaderY + 5.0, { align: 'center' });

    // "SOYADI - ADI" part
    doc.setFontSize(6);
    doc.text('SOYADI - ADI (Soyadı, adı arasına bir karakter boşluk bırakınız.)', rightBlockX + sinifSubeW + 2, nameHeaderY + 4);
    doc.setTextColor(0);

    // --- Grid Area ---
    var gridY = nameHeaderY + 6;
    var gridH = 100; // Go down deep

    // SINIF / ŞUBE Columns (Left of Grid)
    // 2 columns: left is 1-8?, right is A-Z?
    // Visual shows:
    // Left col: bubbles 1..4..8?
    // Right col: bubbles A..Z?
    // Let's implement that.

    // Class Col (1-8)
    var classX = rightBlockX + 1;
    for (var c = 1; c <= 8; c++) {
      var by = gridY + 2 + (c - 1) * 4; // GAP 4
      doc.setDrawColor(pR, pG, pB);
      doc.circle(classX + 2, by, 1.5, 'S');
      doc.setFontSize(5);
      doc.setTextColor(pR, pG, pB);
      doc.text(String(c), classX + 2 - 0.5, by + 0.5);
    }

    // Letter Col (A-Z)
    var branchX = rightBlockX + 7;
    var alphabet = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ";
    for (var l = 0; l < alphabet.length; l++) {
      var by = gridY + 2 + l * 3.3; // Tighter gap for letters
      doc.setDrawColor(pR, pG, pB);
      doc.circle(branchX + 2, by, 1.5, 'S');
      doc.setFontSize(5);
      doc.setTextColor(pR, pG, pB);
      doc.text(alphabet[l], branchX + 2 - 0.5, by + 0.5);
    }

    // NAME GRID (Right of SINIF/ŞUBE)
    var nameGridX = rightBlockX + sinifSubeW;
    var nameGridW = rightBlockW - sinifSubeW;
    // We need columns for Name char positions (e.g. 20 chars)
    // And rows for A-Z
    var nameLen = 20;
    var charW = nameGridW / nameLen;
    var charGapY = 3.3; // Same as letter col

    var nameStr = (student.name || '').toLocaleUpperCase('tr-TR');

    // Draw Bubbles
    for (var r = 0; r < alphabet.length; r++) {
      var rowY = gridY + 2 + r * charGapY;
      var char = alphabet[r];

      for (var c = 0; c < nameLen; c++) {
        var colX = nameGridX + c * charW + charW / 2;

        // Top Header Letters (Inside header? No, usually circles at top row?)
        // Visual has A..Z circles in rows.
        // Top of grid has empty boxes for writing name?
        // Usually yes. Let's add write boxes above.

        // Check if filled
        var targetChar = nameStr[c];
        var isFilled = (targetChar === char);

        if (isFilled) {
          doc.setFillColor(0);
          doc.circle(colX, rowY, 1.3, 'F');
          doc.setTextColor(255);
        } else {
          doc.setDrawColor(pR, pG, pB);
          doc.setLineWidth(0.1);
          doc.circle(colX, rowY, 1.3, 'S');
          doc.setTextColor(pR, pG, pB);
        }

        doc.setFontSize(4);
        var tw = doc.getTextWidth(char);
        doc.text(char, colX - tw / 2, rowY + 0.5);
        doc.setTextColor(0);
      }
    }

    // Draw Name Write Boxes (in the header strip or just below?)
    // Visual shows them just below the pink header, above the bubbles.
    // Since we started bubbles at gridY + 2, let's put boxes at gridY - ?
    // Actually the bubbles start immediately. Let's put boxes *in* the pink header?
    // No, visual: "SOYADI - ADI" pink header, then a row of white boxes for writing chars, then bubbles A..Z.

    // Let's adjust gridY down to make space for write boxes
    // Reset loops? No, just draw boxes at `gridY` and shift bubbles down.
    // Shift bubbles by 5mm.

    // Redo Bubbles with shift
    // (Consolidating visual code)
    // Actually, I'll just adding writing boxes at gridY, and push bubbles to gridY+5

    // Writing boxes
    for (var c = 0; c < nameLen; c++) {
      var colX = nameGridX + c * charW;
      doc.setDrawColor(pR, pG, pB);
      doc.rect(colX + 0.5, gridY, charW - 1, 4);
      // Write char if exists
      if (c < nameStr.length) {
        doc.setFontSize(6);
        doc.text(nameStr[c], colX + charW / 2, gridY + 3, { align: 'center' });
      }
    }

    // Actual Grid Bubbles loop
    var bubblesStartY = gridY + 5;
    for (var l = 0; l < alphabet.length; l++) {
      var by = bubblesStartY + 2 + l * 3.3;
      var char = alphabet[l];

      // Re-draw Branch column bubbles aligned
      doc.setDrawColor(pR, pG, pB);
      doc.circle(branchX + 2, by, 1.4, 'S');
      doc.setFontSize(4);
      doc.setTextColor(pR, pG, pB);
      doc.text(char, branchX + 2 - 0.5, by + 0.5);

      // Name Grid
      for (var c = 0; c < nameLen; c++) {
        var colX = nameGridX + c * charW + charW / 2;
        var targetChar = nameStr[c];
        var isFilled = (targetChar === char);

        if (isFilled) {
          doc.setFillColor(0);
          doc.circle(colX, by, 1.3, 'F');
          doc.setTextColor(255);
        } else {
          doc.setDrawColor(pR, pG, pB);
          doc.setLineWidth(0.1);
          doc.circle(colX, by, 1.3, 'S');
          doc.setTextColor(pR, pG, pB);
        }
        doc.text(char, colX - doc.getTextWidth(char) / 2, by + 0.5);
      }
    }
    doc.setTextColor(0);


    // ============ BOTTOM BLOCK (SÖZEL / SAYISAL) ============
    var bottomY = row3Y + 45; // Start below the student info blocks
    drawLGSLayout(doc, bottomY, pR, pG, pB);

    // ============ FOOTER ============
    doc.setFontSize(5.5);
    setFont('normal');
    doc.setTextColor(80, 80, 80);
    doc.text('Bu Optik Form Mobil Okumaya Uygun Bastırılmıştır.', pageWidth / 2, pageHeight - 5, { align: 'center' });
    doc.text('Form Kodu: LGS 20-20', pageWidth - margin, pageHeight - 5, { align: 'right' });
    doc.setTextColor(0);
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
    var pageMargin = 8;
    var pageWidth = 210;
    var contentWidth = pageWidth - 2 * pageMargin;

    // Split: SÖZEL (Left 4 cols) | SAYISAL (Right 2 cols)
    // Sözel 4 columns: 20, 10, 10, 10
    // Sayısal 2 columns: 20, 20

    var sozelW = contentWidth * 0.62;
    var sayisalW = contentWidth * 0.36;
    var gap = contentWidth - sozelW - sayisalW; // space between

    var sayisalX = pageMargin + sozelW + gap;

    // Headers
    var headerH = 7;
    // SÖZEL Header
    doc.setFillColor(pR, pG, pB);
    doc.roundedRect(pageMargin, startY, sozelW, headerH, 1, 1, 'F');
    doc.setTextColor(255);
    doc.setFontSize(10);
    setFont('bold');
    doc.text('SÖZEL BÖLÜM', pageMargin + sozelW / 2, startY + 5, { align: 'center' });

    // SAYISAL Header
    doc.roundedRect(sayisalX, startY, sayisalW, headerH, 1, 1, 'F');
    doc.text('SAYISAL BÖLÜM', sayisalX + sayisalW / 2, startY + 5, { align: 'center' });
    doc.setTextColor(0);

    // Columns
    var colY = startY + headerH + 2;
    var subHeadH = 10;

    // SÖZEL Columns
    var sozelCols = [
      { title: 'TÜRKÇE', q: 20 },
      { title: 'SOSYAL BİLGİLER\nİNKILAP TARİHİ VE\nATATÜRKÇÜLÜK', q: 10 },
      { title: 'DİN KÜLTÜRÜ\nVE\nAHLAK BİLGİSİ', q: 10 },
      { title: 'İNGİLİZCE', q: 10 }
    ];
    var sColW = sozelW / 4;

    for (var i = 0; i < 4; i++) {
      var cx = pageMargin + i * sColW;

      // SubHeader
      doc.setFillColor(pR, pG, pB);
      doc.rect(cx + 1, colY, sColW - 2, subHeadH, 'F');
      doc.setTextColor(255);
      doc.setFontSize(5);

      var lines = sozelCols[i].title.split('\n');
      for (var li = 0; li < lines.length; li++) {
        doc.text(lines[li], cx + sColW / 2, colY + 3 + li * 2.5, { align: 'center' });
      }

      // Answers
      drawPinkAnswerColumn(doc, sozelCols[i].q, cx, colY + subHeadH + 2, sColW, pR, pG, pB);
    }

    // SAYISAL Columns
    var sayisalColsVals = [
      { title: 'MATEMATİK', q: 20 },
      { title: 'FEN BİLİMLERİ', q: 20 }
    ];
    var mColW = sayisalW / 2;

    for (var j = 0; j < 2; j++) {
      var cx = sayisalX + j * mColW;

      // SubHeader
      doc.setFillColor(pR, pG, pB);
      doc.rect(cx + 1, colY, mColW - 2, subHeadH, 'F');
      doc.setTextColor(255);
      doc.setFontSize(6);
      doc.text(sayisalColsVals[j].title, cx + mColW / 2, colY + 6, { align: 'center' });

      // Answers
      drawPinkAnswerColumn(doc, sayisalColsVals[j].q, cx, colY + subHeadH + 2, mColW, pR, pG, pB);
    }

    // Outer Borders
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    var height = subHeadH + 2 + 20 * 4.5 + 2;
    doc.rect(pageMargin, colY, sozelW, height);
    doc.rect(sayisalX, colY, sayisalW, height);

    // Vertical dividers
    for (var l = 1; l < 4; l++) doc.line(pageMargin + l * sColW, colY, pageMargin + l * sColW, colY + height);
    doc.line(sayisalX + mColW, colY, sayisalX + mColW, colY + height);
  }

  function drawPinkAnswerColumn(doc, count, x, y, availWidth, pR, pG, pB) {
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    var bubbleR = 2.0;
    var gapY = 4.5;
    var opts = ['A', 'B', 'C', 'D'];
    var optGap = (availWidth - 6) / opts.length;

    for (var i = 1; i <= count; i++) {
      var rowY = y + (i - 1) * gapY;

      // Question number
      doc.setFontSize(7);
      setFont('bold');
      doc.setTextColor(pR, pG, pB);
      doc.text(String(i), x + 1, rowY + 1, { align: 'center' });
      doc.setTextColor(0);

      // Bubbles
      for (var o = 0; o < opts.length; o++) {
        var bx = x + 6 + o * optGap;
        doc.setDrawColor(pR, pG, pB);
        doc.setLineWidth(0.25);
        doc.circle(bx, rowY, bubbleR, 'S');
        doc.setFontSize(5.5);
        setFont('bold');
        doc.setTextColor(pR, pG, pB);
        var tw = doc.getTextWidth(opts[o]);
        doc.text(opts[o], bx - tw / 2, rowY + 1);
        doc.setTextColor(0);
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
    var boxY = infoY + 8;
    var boxH = 28;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.4);
    doc.rect(infoX, boxY, infoWidth, boxH);

    // 3. Info Fields (Soyadı - Adı, Sınıf - Şube as requested)
    // We keep others? Let's prioritize Name and Class visually.
    var infoFields = [
      { label: 'Soyadı - Adı', value: (student.name || '').toLocaleUpperCase('tr-TR') },
      { label: 'Sınıf - Şube', value: String(student.classRef || '') }
    ];

    // Add Kurum Adı as extra if space allows, but keep it tight to look like image first
    // Add Kurum Adı - REMOVED AS REQUESTED
    /* if (state.examInfo.institution) {
        infoFields.push({ label: 'Kurum Adı', value: state.examInfo.institution.toLocaleUpperCase('tr-TR') });
    } */

    var startTextY = boxY + 8;
    var lineGap = 7;

    infoFields.forEach(function (f, i) {
      var ty = startTextY + i * lineGap;

      // Label (Pink)
      doc.setFontSize(8);
      setFont('bold');
      doc.setTextColor(pR, pG, pB);
      doc.text(f.label + ' :', infoX + 3, ty);

      // Calculate start of value area
      var labelW = doc.getTextWidth(f.label + ' :');
      var valStartX = infoX + 3 + labelW + 2;
      var lineEndX = infoX + infoWidth - 3;

      // Dotted Line
      doc.setDrawColor(180, 180, 180); // Light grey dots
      doc.setLineWidth(0.2);
      doc.setLineDash([0.5, 0.8], 0);
      doc.line(valStartX, ty + 1, lineEndX, ty + 1);
      doc.setLineDash([], 0);

      // Value
      doc.setTextColor(0);
      setFont('normal'); // Value normal weight? Or bold?
      // Image doesn't show value style but assuming normal/bold black on dotted line.
      doc.setFontSize(8);
      doc.text(f.value, valStartX + 2, ty - 0.5);
    });

    // ============ NUMARANIZ BOX (Left of Info Box) ============
    // NEW DESIGN: Vertical bubble grid on the left side of the info block.
    // The visual shows a tall vertical box with "NUMARANIZ" header and bubbles 0-9 for 6 digits?
    // Actually the visual shows a grid of bubbles. 6 columns (digits). Rows 0-9.

    // Position: To the left of infoX. 
    // infoX is rightX - 25. 
    // We have space between left margin and infoX.
    // Let's place it aligned with the top of the LİSE GRUBU... header. or slightly below?
    // Image: Top aligned with Info box top? Or floating?
    // It seems to be ALIGNED with the top of the header box "LİSE GRUBU..."

    // Width: 6 cols * 4mm = ~24mm.
    var numW = 22;
    var numX = infoX - numW - 5; // Gap 5mm
    var numY = infoY; // Top aligned with header

    // Header "NUMARANIZ"
    doc.setFillColor(pR, pG, pB);
    doc.rect(numX, numY, numW, 6, 'F');
    doc.setTextColor(255);
    doc.setFontSize(6);
    setFont('bold');
    doc.text('NUMARANIZ', numX + numW / 2, numY + 4, { align: 'center' });
    doc.setTextColor(0);

    // White Box for bubbles
    // Height: 10 rows * gap + padding
    var numGridH = 40; // Approx
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(numX, numY + 6, numW, numGridH);

    // Bubbles 0-9 for 6 columns?
    // Student ID max 6 chars?
    var studentNo = String(student.id || '').replace(/[^0-9]/g, '');
    studentNo = studentNo.padStart(6, '0').substring(0, 6);

    var numBubbleGapY = 3.5;
    var numBubbleGapX = 3.5;
    var startNumY = numY + 8.5;
    var startNumX = numX + 2.5;

    for (var r = 0; r < 10; r++) {
      var by = startNumY + r * numBubbleGapY;
      for (var c = 0; c < 6; c++) {
        var bx = startNumX + c * numBubbleGapX;
        var isFilled = (studentNo[c] === String(r));
        drawSimpleBubble(doc, bx, by, 1.3, String(r), isFilled, pR, pG, pB);
      }
    }

    // Text below: "Kitapçık Türünü Kodlamayı Unutmayınız" (Seems to be here in image?)
    // Actually in the image, below the number grid, there is text.
    doc.setFontSize(4.5);
    doc.setTextColor(0);
    setFont('bold');
    doc.text('Kitapçık Türünü', numX + numW / 2, numY + 6 + numGridH + 3, { align: 'center' });
    doc.text('Kodlamayı Unutmayınız.', numX + numW / 2, numY + 6 + numGridH + 6, { align: 'center' });


    // ============ DİKKAT SECTION ============
    // Visual: Pink "DİKKAT" box left, then white box with "Yanlış kodlama" examples, then "Doğru kodlama" example, then text.
    var dikkatY = boxY + boxH; // Start exactly below the info box to connect them? Or slightly overlapping border?
    // Visual shows them connected. The Pink 'DİKKAT' box seems to be attached to the bottom left of the white info box?
    // Actually, looking at image: The "DİKKAT" row is attached to the bottom of the "Soyadı-Adı / Sınıf" box.
    // So boxY + boxH is correct.
    var dikkatH = 7;

    // 1. DİKKAT Box (Pink)
    var dikkatW = 16;
    doc.setFillColor(pR, pG, pB);
    doc.rect(infoX, dikkatY, dikkatW, dikkatH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(7); // Slightly larger
    setFont('bold');
    doc.text('DİKKAT', infoX + dikkatW / 2, dikkatY + 4.5, { align: 'center' });
    doc.setTextColor(0);

    // 2. Yanlış kodlama Box (White with Pink Border)
    var yanlisX = infoX + dikkatW;
    var yanlisW = 38;
    doc.setDrawColor(pR, pG, pB);
    doc.setLineWidth(0.3);
    doc.rect(yanlisX, dikkatY, yanlisW, dikkatH);

    doc.setTextColor(pR, pG, pB);
    doc.setFontSize(5.5);
    setFont('bold');
    doc.text('Yanlış kodlama', yanlisX + yanlisW / 2, dikkatY + 2.5, { align: 'center' });
    doc.setTextColor(0);

    // Examples: Circle with center dot, filled oval, tick, cross, scribble, dash
    // We'll draw 5-6 small circles
    var startCircX = yanlisX + 3.5;
    var circY = dikkatY + 5;
    var gap = 5.5;

    // 1. Dot in center
    doc.setDrawColor(0);
    doc.setLineWidth(0.15);
    doc.circle(startCircX, circY, 1.8, 'S');
    doc.circle(startCircX, circY, 0.5, 'F'); // Dot

    // 2. Vertical Oval / Bean (Simulated)
    doc.ellipse(startCircX + gap, circY, 1.0, 1.8, 'F'); // Filled ovalish

    // 3. Tick
    doc.circle(startCircX + gap * 2, circY, 1.8, 'S');
    doc.setFontSize(5);
    doc.text('✔', startCircX + gap * 2 - 1, circY + 1.2);
    // Or lines manually if char not avail? Checkmark usually fine.

    // 4. Cross
    doc.circle(startCircX + gap * 3, circY, 1.8, 'S');
    doc.text('X', startCircX + gap * 3 - 1, circY + 1.2);

    // 5. Scribble (Zigzag line)
    doc.circle(startCircX + gap * 4, circY, 1.8, 'S');
    doc.line(startCircX + gap * 4 - 1, circY, startCircX + gap * 4 + 1, circY);
    // Just a line through

    // 6. Dash
    doc.circle(startCircX + gap * 5, circY, 1.8, 'S');
    doc.text('-', startCircX + gap * 5 - 0.5, circY + 1);

    // 3. Doğru kodlama Box
    var dogruX = yanlisX + yanlisW;
    var dogruW = 22;
    doc.setDrawColor(pR, pG, pB);
    doc.rect(dogruX, dikkatY, dogruW, dikkatH);

    doc.setTextColor(pR, pG, pB);
    doc.setFontSize(5.5);
    // setFont('bold'); // already bold
    doc.text('Doğru kodlama', dogruX + dogruW / 2, dikkatY + 2.5, { align: 'center' });
    doc.setTextColor(0);

    // Correct bubble
    doc.setFillColor(0);
    doc.circle(dogruX + dogruW / 2, circY, 1.8, 'F');

    // 4. Text on Right
    var textX = dogruX + dogruW + 3;
    doc.setTextColor(pR, pG, pB);
    doc.setFontSize(5.5);
    setFont('bold');

    // "Kodlamamızı lütfen yumuşak kurşun kalem ile yapınız." 
    // Using standard correct text.
    doc.text('Kodlamalarınızı lütfen yumuşak', textX, dikkatY + 2.5);
    doc.text('kurşun kalem ile yapınız.', textX + 10, dikkatY + 5.5); // Indented slightly for center alignment visual
    doc.setTextColor(0);

    // ============ REFINED LAYOUT (BELOW DİKKAT) ============
    // Replaces correct/booklet type/answer sections with the new comprehensive layout
    var refinedStartY = dikkatY + dikkatH + 5;
    drawRefinedTYTLayout(doc, refinedStartY, pR, pG, pB, student);

    // ============ FOOTER ============
    doc.setFontSize(5);
    setFont('normal');
    doc.setTextColor(80, 80, 80);
    doc.text('Bu Optik Form Mobil Okumaya Uygun Bastırılmıştır.', pageWidth / 2, pageHeight - 5, { align: 'center' });
    doc.text('Form Kodu: ProNET YKS-1', pageWidth - margin, pageHeight - 5, { align: 'right' });
    doc.setTextColor(0);
  }

  function drawRefinedTYTLayout(doc, startY, pR, pG, pB, student) {
    var setFont = function (style) {
      if (window.fontRobotoRegular) doc.setFont('Roboto', style);
    };

    var margin = 8;
    var pageWidth = 210;
    var contentW = pageWidth - 2 * margin;

    // LEFT BLOCK (Kitapçık, TC, Sınıf, Name) approx 42%
    // RIGHT BLOCK (Oturum, Answers) approx 56%
    // GAP 2%
    var leftW = contentW * 0.42;
    var gap = contentW * 0.02;
    var rightW = contentW - leftW - gap;
    var rightX = margin + leftW + gap;

    // ================= LEFT BLOCK =================

    // 1. KİTAPÇIK TÜRÜ (Pink Header)
    var row1H = 6;
    doc.setFillColor(pR, pG, pB);
    doc.rect(margin, startY, leftW, row1H, 'F');
    doc.setTextColor(255);
    doc.setFontSize(6);
    setFont('bold');
    doc.text('KİTAPÇIK TÜRÜ', margin + 2, startY + 4);
    doc.setTextColor(0);

    // Bubbles (1 2 3 4 5) - Right aligned in header? 
    // Image shows white circles with numbers inside the pink header or just below?
    // Image: "KİTAPÇIK TÜRÜ" text left, bubbles right. Bubbles seem to be in white boxes or just circles.
    // Let's put white circles in the pink header for valid visual.
    var ktBubs = ['1', '2', '3', '4', '5'];
    for (var k = 0; k < ktBubs.length; k++) {
      var bx = margin + leftW - 35 + k * 6;
      doc.setFillColor(255);
      doc.circle(bx, startY + 3, 2, 'F');
      doc.setTextColor(0);
      doc.setFontSize(5);
      doc.text(ktBubs[k], bx - 0.6, startY + 3.8);
    }

    // 2. TC / CEP & SINIF / ŞUBE / GRUP (Side by Side)
    var row2Y = startY + row1H + 2;

    // Col 1: TC (11 cols)
    // Col 2: Sınıf(2)+Şube(1)+Grup(2) = 5 cols.
    // Total 16 cols. 
    var colW = leftW / 16;

    // Headers
    var header2H = 10;
    var tcWidth = colW * 11;
    var sgWidth = leftW - tcWidth - 2; // Gap 2mm

    // TC Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(margin, row2Y, tcWidth, header2H, 'F');
    doc.setTextColor(255);
    doc.setFontSize(5);
    doc.text('T.C. KİMLİK NO. / CEP TEL. NO.', margin + tcWidth / 2, row2Y + 6, { align: 'center' });

    // SG Header
    doc.setFillColor(pR, pG, pB);
    doc.rect(margin + tcWidth + 1, row2Y, sgWidth, header2H, 'F');

    // SG Sub-labels (SINIF | ŞUBE | GRUP NO)
    // Sınıf (2), Şube (1), Grup (2)
    var subColW = sgWidth / 5;
    doc.text('SINIF', margin + tcWidth + 1 + subColW, row2Y + 3, { align: 'center' });
    doc.text('ŞUBE', margin + tcWidth + 1 + subColW * 2.5, row2Y + 3, { align: 'center' });
    doc.text('GRUP NO', margin + tcWidth + 1 + subColW * 4, row2Y + 3, { align: 'center' });

    // White boxes for writing?
    // Visual shows boxes below header.
    doc.setDrawColor(pR, pG, pB);
    var writeBoxH = 4;
    var writeY = row2Y + header2H;

    // TC Write Boxes
    var tcVal = (student.tc || '').replace(/[^0-9]/g, '');
    for (var i = 0; i < 11; i++) {
      doc.rect(margin + i * colW, writeY, colW, writeBoxH);
      if (i < tcVal.length) {
        doc.setTextColor(0); doc.text(tcVal[i], margin + i * colW + colW / 2 - 0.5, writeY + 3);
      }
    }

    // SG Write Boxes
    var sgX = margin + tcWidth + 1;
    var sgSubW = sgWidth / 5;
    // Sınıf (2)
    for (var i = 0; i < 2; i++) doc.rect(sgX + i * sgSubW, writeY, sgSubW, writeBoxH);
    // Şube (1)
    doc.rect(sgX + 2 * sgSubW, writeY, sgSubW, writeBoxH);
    // Grup (2)
    for (var i = 0; i < 2; i++) doc.rect(sgX + (3 + i) * sgSubW, writeY, sgSubW, writeBoxH);


    // Bubbles
    var gridY = writeY + writeBoxH;
    var bubbleGap = 3.5;
    var rows = 10; // 0-9

    // TC Bubbles 0-9
    for (var r = 0; r < 10; r++) {
      var by = gridY + 2 + r * bubbleGap;
      for (var c = 0; c < 11; c++) {
        var isFilled = (tcVal[c] === String(r));
        drawSimpleBubble(doc, margin + c * colW + colW / 2, by, 1.4, String(r), isFilled, pR, pG, pB);
      }
    }

    // SG Bubbles
    // Sınıf (0-9)
    for (var r = 0; r < 10; r++) {
      var by = gridY + 2 + r * bubbleGap;
      for (var c = 0; c < 2; c++) {
        // Determine dummy class val?
        drawSimpleBubble(doc, sgX + c * sgSubW + sgSubW / 2, by, 1.4, String(r), false, pR, pG, pB);
      }
      // Grup (0-9) - Cols 3,4 (indices 3,4 of 5)
      for (var c = 3; c < 5; c++) {
        drawSimpleBubble(doc, sgX + c * sgSubW + sgSubW / 2, by, 1.4, String(r), false, pR, pG, pB);
      }
    }
    // Şube (A-Z) - Needs more vertical space!
    // The image shows the layout stretches down.
    // Şube is Col index 2.
    var alphabet = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ";
    for (var r = 0; r < alphabet.length; r++) {
      var by = gridY + 2 + r * 3.0; // tighter gap
      drawSimpleBubble(doc, sgX + 2 * sgSubW + sgSubW / 2, by, 1.4, alphabet[r], false, pR, pG, pB);
    }


    // 3. SOYADI - ADI GRID
    var nameY = gridY + 10 * bubbleGap + 5; // Start after TC grid
    var nameH = 6;
    doc.setFillColor(pR, pG, pB);
    doc.rect(margin, nameY, leftW, nameH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(5);
    doc.text('SOYADI - ADI (Soyadı, adı arasına bir karakter boşluk bırakınız.)', margin + 2, nameY + 4);

    var nameGridY = nameY + nameH;
    // 20 cols? Match width
    var nameCols = 20;
    var nameColW = leftW / nameCols;

    var nameStr = (student.name || '').toLocaleUpperCase('tr-TR');

    // Bubbles A-Z
    for (var r = 0; r < alphabet.length; r++) {
      var by = nameGridY + 2 + r * 3.0;
      var char = alphabet[r];
      for (var c = 0; c < nameCols; c++) {
        var isFilled = (nameStr[c] === char);
        var bx = margin + c * nameColW + nameColW / 2;

        // Draw bubble circle
        if (isFilled) {
          doc.setFillColor(0);
          doc.circle(bx, by, 1.3, 'F');
          doc.setTextColor(255);
        } else {
          doc.setDrawColor(pR, pG, pB);
          doc.setLineWidth(0.1);
          doc.circle(bx, by, 1.3, 'S');
          doc.setTextColor(pR, pG, pB);
        }
        doc.setFontSize(3.5);
        doc.text(char, bx - 0.4, by + 0.4);
      }
    }
    doc.setTextColor(0);


    // ================= RIGHT BLOCK =================

    // 1. Headers (OTURUM, TEPM)
    var headerH = 8;
    // OTURUM
    var oturumW = rightW * 0.4;
    doc.setFillColor(pR, pG, pB);
    doc.rect(rightX, startY, oturumW, headerH, 'F');
    doc.setTextColor(255);
    doc.setFontSize(6);
    doc.text('OTURUM', rightX + 2, startY + 5);

    // Bubbles 1. OTURUM, 2. OTURUM
    doc.circle(rightX + 25, startY + 4, 2, 'S'); doc.text('1. OTURUM', rightX + 20, startY + 7.5);
    doc.circle(rightX + 45, startY + 4, 2, 'S'); doc.text('2. OTURUM', rightX + 40, startY + 7.5);

    // TEPM (ALAN)
    var alanX = rightX + oturumW + 2;
    var alanW = rightW - oturumW - 2;
    doc.setFillColor(pR, pG, pB);
    doc.rect(alanX, startY, alanW, headerH, 'F');
    doc.text('TEPM', alanX + 2, startY + 5); // Label as TEPM or ALAN
    // 4 Bubbles
    var alans = ['SÖZEL', 'SAYISAL', 'E.A.', 'DİL'];
    for (var a = 0; a < alans.length; a++) {
      var abx = alanX + 10 + a * 12;
      doc.circle(abx, startY + 3, 2, 'S');
      doc.setFontSize(4);
      doc.text(alans[a], abx - 2, startY + 7);
    }


    // 2. Answer Columns (4 Cols)
    var colsY = startY + headerH + 2;
    var colGap = 2;
    var ansColW = (rightW - 3 * colGap) / 4;

    var colTitles = [
      ['TÜRKÇE', 'T. DİLİ VE EDEB.', 'SOSYAL BİL. 1'],
      ['SOSYAL', 'BİLİMLER', 'SOSYAL', 'BİLİMLER 2'],
      ['TEMEL', 'MATEMATİK', '', 'MATEMATİK'],
      ['FEN', 'BİLİMLERİ', '', 'FEN BİLİMLERİ']
    ];

    for (var i = 0; i < 4; i++) {
      var cx = rightX + i * (ansColW + colGap);

      // Header Box
      doc.setFillColor(pR, pG, pB);
      doc.rect(cx, colsY, ansColW, 12, 'F'); // Tall header
      doc.setTextColor(255);
      doc.setFontSize(4.5);

      // Titles (Multi line)
      var tlines = colTitles[i];
      tlines.forEach(function (l, idx) {
        doc.text(l, cx + ansColW / 2, colsY + 3 + idx * 2.5, { align: 'center' });
      });

      // 1-40 Answer Grid
      var gridY = colsY + 12 + 2;
      var rGap = 3.2;

      for (var q = 1; q <= 40; q++) {
        var qy = gridY + (q - 1) * rGap;

        // Zebra striping for 5s?
        if (Math.ceil(q / 5) % 2 === 0) {
          // doc.setFillColor(240, 240, 240);
          // doc.rect(cx, qy-2, ansColW, rGap, 'F');
        }

        // Q Num
        doc.setTextColor(0);
        doc.setFontSize(5);
        doc.text(String(q), cx + 2, qy + 1);

        // Bubbles A-E
        var opts = ['A', 'B', 'C', 'D', 'E'];
        var optW = (ansColW - 6) / 5;
        for (var o = 0; o < 5; o++) {
          var obx = cx + 6 + o * optW + optW / 2;
          doc.setDrawColor(pR, pG, pB);
          doc.setLineWidth(0.15); // Thinner
          doc.circle(obx, qy, 1.3, 'S'); // Circle
          doc.setTextColor(pR, pG, pB);
          doc.setFontSize(3.5);
          doc.text(opts[o], obx - 0.4, qy + 0.4);
        }
      }
    }
  }

  function drawSimpleBubble(doc, x, y, r, txt, filled, pR, pG, pB) {
    if (filled) {
      doc.setFillColor(0);
      doc.circle(x, y, r, 'F');
      doc.setTextColor(255);
    } else {
      doc.setDrawColor(pR, pG, pB);
      doc.setLineWidth(0.15);
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

  document.addEventListener('DOMContentLoaded', init);

})();
