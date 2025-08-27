// Global variables for data storage
let idToUni = {};
let idToStud = {};
let manualAssignments = [];
let forbiddenAssignments = [];
let selectedStudentForAssignment = null;

const FACULTY_MAPPING = {
  'BMS': ['Psychology', 'Communication Science', 'Industrial Engineering and Management',
          'International Business Administration', 'Management, Society & Technology', 'Educational Science and Technology'],
  'EEMCS': ['Applied Mathematics', 'Business & IT', 'Computer Science', 'Creative Technology',
            'Electrical Engineering', 'Business Information Technology', 'Embedded Systems',
            'Interaction Technology', 'Robotics', 'Systems and Control'],
  'ET': ['Civil Engineering', 'Industrial Design Engineering', 'Mechanical Engineering',
         'Civil Engineering and Management', 'Construction Management & Engineering',
         'Sustainable Energy Technology', 'Mechanical Engineering - Amsterdam (VU-UT)'],
  'ITC': ['Spatial Engineering', 'Geo-information Science and Earth Observation'],
  'ST': ['Advanced Technology', 'Biomedical Engineering', 'Chemical Science & Engineering', 'Health Sciences',
         'Technical Medicine', 'Applied Physics', 'Nanotechnology', 'Water Technology',
         'Materials Science & Engineering', 'Fluid Dynamics'],
  'UCT': ['Technology and Liberal Arts & Sciences (ATLAS)']
};

function getFacultyForStudyField(studyField) {
  studyField = studyField.replace(/\s*\(BSc\)|\s*\(MSc\)/g, '').trim();
  for (const [faculty, fields] of Object.entries(FACULTY_MAPPING)) {
    if (fields.includes(studyField)) {
      return faculty;
    }
  }
  console.error(`Unknown faculty for study field: ${studyField}`);
  return null;
}

function hasUniFacultyConstraints(uni) {
  return (uni.max_BMS || 0) > 0 || (uni.max_EEMCS || 0) > 0 || (uni.max_ET || 0) > 0 ||
         (uni.max_ITC || 0) > 0 || (uni.max_ST || 0) > 0 || (uni.max_UCT || 0) > 0;
}

function hasFacultyConstraints() {
  return Object.values(idToUni).some(uni => hasUniFacultyConstraints(uni));
}


// Check if manual assignments are feasible
function checkManualAssignmentsFeasibility() {
  const uniCounts = {};
  const errors = [];
  
  for (const assignment of manualAssignments) {
    const student = idToStud[assignment.studentId];
    const uni = idToUni[assignment.uniKey];
    
    if (!student || !uni) {
      errors.push(`Invalid assignment for student ${assignment.studentId}`);
      continue;
    }
    
    const semester = student.semester === '1st semester' || student.semester === 'Full academic year' ? 'f' : 's';
    
    if (!uniCounts[uni.key]) {
      uniCounts[uni.key] = { f: 0, s: 0, BSc_f: 0, MSc_f: 0, BSc_s: 0, MSc_s: 0 };
    }
    
    if (!uniCounts[uni.agreement_ID]) {
      uniCounts[uni.agreement_ID] = { f: 0, s: 0 };
    }
    
    uniCounts[uni.key][semester]++;
    uniCounts[uni.agreement_ID][semester]++;
    
    if (student.study_level === 'BSc') {
      uniCounts[uni.key][`BSc_${semester}`]++;
    } else if (student.study_level === 'MSc') {
      uniCounts[uni.key][`MSc_${semester}`]++;
    }
    
    const semKey = semester === 'f' ? 'total1' : 'total2';
    const levelKey = student.study_level === 'BSc' ? (semester === 'f' ? 'B1' : 'B2') : (semester === 'f' ? 'M1' : 'M2');
    
    if (uniCounts[uni.key][semester] > uni[semKey]) {
      errors.push(`${uni.key} capacity exceeded in ${semester === 'f' ? '1st' : '2nd'} semester (${uniCounts[uni.key][semester]}/${uni[semKey]})`);
    }
    
    if (uniCounts[uni.key][`${student.study_level}_${semester}`] > uni[levelKey]) {
      errors.push(`${uni.key} ${student.study_level} capacity exceeded in ${semester === 'f' ? '1st' : '2nd'} semester (${uniCounts[uni.key][`${student.study_level}_${semester}`]}/${uni[levelKey]})`);
    }
    
    const agreementKey = semester === 'f' ? 'total1' : 'total2';
    if (uniCounts[uni.agreement_ID][semester] > uni[agreementKey]) {
      errors.push(`Agreement ${uni.agreement_ID} capacity exceeded in ${semester === 'f' ? '1st' : '2nd'} semester (${uniCounts[uni.agreement_ID][semester]}/${uni[agreementKey]})`);
    }
  }
  
  if (hasFacultyConstraints()) {
    const universityFacultyCounts = {};
    
    for (const assignment of manualAssignments) {
      const student = idToStud[assignment.studentId];
      const uni = idToUni[assignment.uniKey];
      
      if (!student || !uni || !hasUniFacultyConstraints(uni)) continue;
      
      const studentFaculty = getFacultyForStudyField(student.study_field);
      if (!studentFaculty) continue;
      
      const semester = student.semester === '1st semester' || student.semester === 'Full academic year' ? 'f' : 's';
      
      if (!universityFacultyCounts[uni.key]) {
        universityFacultyCounts[uni.key] = {};
        ['BMS', 'EEMCS', 'ET', 'ITC', 'ST', 'UCT'].forEach(faculty => {
          universityFacultyCounts[uni.key][faculty] = { f: 0, s: 0 };
        });
      }
      
      universityFacultyCounts[uni.key][studentFaculty][semester]++;
    }
    
    Object.keys(universityFacultyCounts).forEach(uniKey => {
      const uni = idToUni[uniKey];
      if (!uni) return;
      
      ['BMS', 'EEMCS', 'ET', 'ITC', 'ST', 'UCT'].forEach(faculty => {
        const facultyMax = uni[`max_${faculty}`] || 0;
        
        if (facultyMax > 0) {
          const used1 = universityFacultyCounts[uniKey][faculty].f || 0;
          const used2 = universityFacultyCounts[uniKey][faculty].s || 0;
          
          if (used1 > facultyMax) {
            errors.push(`University ${uniKey} faculty ${faculty} capacity exceeded in 1st semester (${used1}/${facultyMax})`);
          }
          
          if (used2 > facultyMax) {
            errors.push(`University ${uniKey} faculty ${faculty} capacity exceeded in 2nd semester (${used2}/${facultyMax})`);
          }
        }
      });
    });
  }
  
  const assignedStudents = {};
  for (const assignment of manualAssignments) {
    if (assignedStudents[assignment.studentId]) {
      errors.push(`Student ${assignment.studentId} is manually assigned multiple times`);
    }
    assignedStudents[assignment.studentId] = true;
  }
  
  return { feasible: errors.length === 0, errors };
}

// Update the manual assignments list display
function updateManualAssignmentsList() {
  const tbody = document.querySelector('#manualAssignmentsTable tbody');
  tbody.innerHTML = '';
  
  if (manualAssignments.length === 0) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 3;
    cell.textContent = 'No manual assignments yet';
    cell.style.textAlign = 'center';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }
  
  manualAssignments.forEach((assignment, index) => {
    const row = document.createElement('tr');
    
    const studentCell = document.createElement('td');
    const student = idToStud[assignment.studentId];
    if (student && student.fullName) {
      studentCell.textContent = `${assignment.studentId} (${student.fullName})`;
    } else {
      studentCell.textContent = assignment.studentId;
    }
    
    const uniCell = document.createElement('td');
    const uni = idToUni[assignment.uniKey];
    if (uni && uni.partner_name) {
      uniCell.textContent = `${assignment.uniKey} (${uni.partner_name})`;
    } else {
      uniCell.textContent = assignment.uniKey;
    }
    
    const actionCell = document.createElement('td');
    const removeBtn = document.createElement('button');
    removeBtn.textContent = '✖';
    removeBtn.className = 'remove-assignment';
    removeBtn.addEventListener('click', () => {
      manualAssignments.splice(index, 1);
      updateManualAssignmentsList();
    });
    actionCell.appendChild(removeBtn);
    
    row.appendChild(studentCell);
    row.appendChild(uniCell);
    row.appendChild(actionCell);
    
    tbody.appendChild(row);
  });
}

  // Update the forbidden assignments list display
  function updateForbiddenAssignmentsList() {
    const tbody = document.querySelector('#forbiddenAssignmentsTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (forbiddenAssignments.length === 0) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 3;
      cell.textContent = 'No forbidden assignments yet';
      cell.style.textAlign = 'center';
      row.appendChild(cell);
      tbody.appendChild(row);
      return;
    }

    forbiddenAssignments.forEach((item, index) => {
      const row = document.createElement('tr');

      const studentCell = document.createElement('td');
      const student = idToStud[item.studentId];
      if (student && student.fullName) {
        studentCell.textContent = `${item.studentId} (${student.fullName})`;
      } else {
        studentCell.textContent = String(item.studentId);
      }
      row.appendChild(studentCell);

      const uniCell = document.createElement('td');
      const uni = idToUni[item.uniKey];
      if (uni && uni.partner_name) {
        uniCell.textContent = `${item.uniKey} (${uni.partner_name})`;
      } else {
        uniCell.textContent = item.uniKey;
      }
      row.appendChild(uniCell);

      const actionCell = document.createElement('td');
      const removeBtn = document.createElement('button');
      removeBtn.textContent = '✖';
      removeBtn.className = 'remove-assignment';
      removeBtn.addEventListener('click', () => {
        forbiddenAssignments.splice(index, 1);
        updateForbiddenAssignmentsList();
      });
      actionCell.appendChild(removeBtn);
      row.appendChild(actionCell);

      tbody.appendChild(row);
    });
  }

// Setup student search functionality
function setupStudentSearch() {
  const searchInput = document.getElementById('studentSearch');
  const resultsDiv = document.getElementById('studentResults');
  const assignmentContainer = document.querySelector('.assignment-container');
  const selectedStudentNameSpan = document.getElementById('selectedStudentName');
  const preferenceOptionsDiv = document.getElementById('preferenceOptions');
  const addBtn = document.getElementById('addAssignmentBtn');
  const cancelBtn = document.getElementById('cancelAssignmentBtn');
  const forbidBtn = document.getElementById('forbidAssignmentBtn');
  
  // Hide results when clicking outside
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.search-container')) {
      resultsDiv.style.display = 'none';
    }
  });
  
  // Handle search input
  searchInput.addEventListener('input', () => {
    const query = searchInput.value.toLowerCase();
    
    if (query.length < 2) {
      resultsDiv.style.display = 'none';
      return;
    }
    
    resultsDiv.innerHTML = '';
    resultsDiv.style.display = 'block';
    let matchCount = 0;
    
    // Search by ID and name
    Object.values(idToStud).forEach(student => {
      const id = student.ID.toString();
      const name = student.fullName || '';
      
      if ((id.includes(query) || name.toLowerCase().includes(query)) && matchCount < 10) {
        const div = document.createElement('div');
        div.className = 'search-result-item';
        
        if (name) {
          div.textContent = `${id} (${name})`;
        } else {
          div.textContent = id;
        }
        
        div.addEventListener('click', () => {
          selectedStudentForAssignment = student;
          showStudentPreferences(student);
          assignmentContainer.style.display = 'block';
          if (name) {
            selectedStudentNameSpan.textContent = `${id} (${name})`;
          } else {
            selectedStudentNameSpan.textContent = id;
          }
          resultsDiv.style.display = 'none';
          searchInput.value = '';
        });
        
        resultsDiv.appendChild(div);
        matchCount++;
      }
    });
    
    if (matchCount === 0) {
      const div = document.createElement('div');
      div.className = 'search-result-item';
      div.textContent = 'No students found';
      div.style.fontStyle = 'italic';
      resultsDiv.appendChild(div);
    }
  });
  
  // Handle cancel button
  cancelBtn.addEventListener('click', () => {
    assignmentContainer.style.display = 'none';
    selectedStudentForAssignment = null;
    preferenceOptionsDiv.innerHTML = '';
  });
  
  // Handle add assignment button
  addBtn.addEventListener('click', () => {
    const selectedPref = preferenceOptionsDiv.querySelector('.preference-option.selected');
    
    if (!selectedPref || !selectedStudentForAssignment) {
      alert('Please select a preference first');
      return;
    }
    
    const uniKey = selectedPref.dataset.unikey;
    
    // Check if student is already assigned
    const existingIndex = manualAssignments.findIndex(a => a.studentId === selectedStudentForAssignment.ID);
    if (existingIndex !== -1) {
      manualAssignments[existingIndex] = { 
        studentId: selectedStudentForAssignment.ID, 
        uniKey: uniKey 
      };
      alert(`Updated assignment for student ${selectedStudentForAssignment.ID}`);
    } else {
      manualAssignments.push({ 
        studentId: selectedStudentForAssignment.ID, 
        uniKey: uniKey 
      });
    }
    
    updateManualAssignmentsList();
    assignmentContainer.style.display = 'none';
    selectedStudentForAssignment = null;
    preferenceOptionsDiv.innerHTML = '';
  });

  // Handle forbid button
  if (forbidBtn) {
    forbidBtn.addEventListener('click', () => {
      const selectedPref = preferenceOptionsDiv.querySelector('.preference-option.selected');
      if (!selectedPref || !selectedStudentForAssignment) {
        alert('Please select a preference first');
        return;
      }
      const uniKey = selectedPref.dataset.unikey;
      // Avoid duplicates
      const exists = forbiddenAssignments.some(a => a.studentId === selectedStudentForAssignment.ID && a.uniKey === uniKey);
      if (!exists) {
        forbiddenAssignments.push({ studentId: selectedStudentForAssignment.ID, uniKey });
        updateForbiddenAssignmentsList();
      } else {
        alert('This forbid rule already exists');
      }
    });
  }
}

// Display a student's university preferences for assignment
function showStudentPreferences(student) {
  const preferenceOptionsDiv = document.getElementById('preferenceOptions');
  preferenceOptionsDiv.innerHTML = '';
  
  if (!student.all_preferences || student.all_preferences.length === 0) {
    const div = document.createElement('div');
    div.textContent = 'No preferences found for this student';
    div.style.fontStyle = 'italic';
    preferenceOptionsDiv.appendChild(div);
    return;
  }
  
  student.all_preferences.forEach((pref, index) => {
    if (!pref) return;
    
    const uni = idToUni[pref];
    if (!uni) return;
    
    const div = document.createElement('div');
    div.className = 'preference-option';
    div.dataset.unikey = pref;
    
    if (uni.partner_name) {
      div.textContent = `Choice ${index + 1}: ${pref} (${uni.partner_name})`;
    } else {
      div.textContent = `Choice ${index + 1}: ${pref}`;
    }
    
    div.addEventListener('click', () => {
      preferenceOptionsDiv.querySelectorAll('.preference-option').forEach(el => {
        el.classList.remove('selected');
      });
      div.classList.add('selected');
    });
    
    preferenceOptionsDiv.appendChild(div);
  });
}

// File loading state variables
let universitiesLoaded = false;
let applicantsLoaded = false;
let universitiesData = null;
let applicantsData = null;

// Process universities data into the format needed for the matching algorithm
function processUniversitiesData() {
  // Normalize data and only keep rows with IDs
  const uniData = universitiesData.filter(r => typeof r['ID'] === 'number');
  
  const unis = uniData.map(r => ({
    agreement_ID: r['ID'],
    country: r['Host country'],
    partner_name: r['Partner institution'],
    academic_year: r['Academic year'],
    ISCED: r['ISCED 2013 code'],
    study_abbr: r['Abbr. of study field'],
    study_field: r['Study field'],
    agreement_type: r['Agreement type'],
    total_places: r['Total #'],
    max_BCs: r['Max # BSc'],
    max_MCs: r['Max # MSc'],
    spots_first_s: r['Nr of agreed spots in 1st sem.'],
    spots_second_s: r['Nr of agreed spots in 2nd sem.'],
    students_assigned: r['Students assigned'],
    partner_department: r['Partner department or consortium'],
    comments_spot_dis: r['Comments on spot distribution'],
    comments_agreement: r['Comments regarding the agreement (in portal)'],
    max_BMS: r['The maximum number of spots/seats for BMS'] || 0,
    max_EEMCS: r['The maximum number of spots/seats for EEMCS'] || 0,
    max_ET: r['The maximum number of spots/seats for ET'] || 0,
    max_ITC: r['The maximum number of spots/seats for ITC'] || 0,
    max_ST: r['The maximum number of spots/seats for ST'] || 0,
    max_UCT: r['The maximum number of spots/seats for UCT'] || 0
  }));
  
  idToUni = {};
  const agreementGroups = {};
  
  unis.forEach(r => {
    const sem = computeSemesters(r.spots_first_s, r.spots_second_s, r.total_places, r.max_BCs, r.max_MCs);
    const uni = Object.assign({}, r, { ...sem });
    const key = r.agreement_ID + '_' + r.study_abbr;
    uni.key = key;
    idToUni[key] = uni;
    
    agreementGroups[r.agreement_ID] = agreementGroups[r.agreement_ID] || [];
    agreementGroups[r.agreement_ID].push(uni);
  });
  
  // Add special universities for unmatched students and special cases
  ['dummy', '1', '2', '3', '4', '5', '6'].forEach(id => {
    const key = id === 'dummy' ? 'dummy-dummy' : `${id}-fictional`;
    const uni = { 
      agreement_ID: id, 
      study_abbr: id, 
      key, 
      total1: 10000, 
      B1: 10000, 
      M1: 10000, 
      total2: 10000, 
      B2: 10000, 
      M2: 10000,
    };
    idToUni[key] = uni;

    if (id !== 'dummy') {
      agreementGroups[id] = agreementGroups[id] || [];
      agreementGroups[id].push(uni);
    }
  });
  
  return {unis, agreementGroups};
}

// Process applicants data into the format needed for the matching algorithm
function processApplicantsData() {
  // Normalize data and only keep rows with IDs
  const appData = applicantsData.filter(r => typeof r['ID of application'] === 'number');
  
  const students = appData.map(r => {
  const studentNumber = r['Student number'] == null ? '' : String(r['Student number']);
    const hasFirstName = Object.keys(r).some(key => key === 'First name');
    const hasLastName = Object.keys(r).some(key => key.startsWith('Last name'));
    const firstName = hasFirstName ? (r['First name'] || '') : '';
    const lastNameKey = hasLastName ? Object.keys(r).find(key => key.startsWith('Last name')) : null;
    const lastName = lastNameKey ? r[lastNameKey] || '' : '';
    const fullName = [firstName, lastName].filter(Boolean).join(' ');
    return {
      ID: r['ID of application'],
      student_number: studentNumber,
      firstName: firstName,
      lastName: lastName,
      fullName: fullName,
      abbreviation: r['Abbreviation of study field'],
      study_field: r['Study field'],
      study_level: r['Study level'],
      academic_year: r['Academic year'],
      semester: r['Semester'],
      preference_1: r['Agreement-ID 1st choice'],
      preference_2: r['Agreement-ID 2nd choice'],
      preference_3: r['Agreement-ID 3rd choice'],
      preference_4: r['Agreement-ID 4th choice'],
      preference_5: r['Agreement-ID 5th choice'],
      preference_6: r['Agreement-ID 6th choice']
    };
  });
  
  const studentPrefsMap = {};
  idToStud = {};
  
  students.forEach(r => {
    const prefs = [];
    [r.preference_1, r.preference_2, r.preference_3, r.preference_4, r.preference_5, r.preference_6].forEach((p, i) => {
      prefs.push(p ? p + '_' + r.abbreviation : null);
    });
    studentPrefsMap[r.ID] = prefs;
    
    const stud = Object.assign({}, r, { all_preferences: prefs });
    idToStud[r.ID] = stud;
  });
  
  window.studentPrefsMap = studentPrefsMap;
  
  return students;
}

// Process uploaded files and initialize the solver UI
function processFiles() {
  const resultsDiv = document.getElementById('results');
  const manualAssignmentDiv = document.getElementById('manualAssignment');
  const prefWeightsContainer = document.getElementById('prefWeightsContainer');
  
  // If both files aren't loaded yet, wait
  if (!universitiesLoaded || !applicantsLoaded) {
    return;
  }
  
  resultsDiv.innerHTML = '';
  
  // Clean up student preferences - remove non-existing options and duplicates, shift remaining up
  cleanupStudentPreferences();
  
  // Show manual assignment section and exchange I factor after files are loaded
  manualAssignmentDiv.style.display = 'block';
  if (prefWeightsContainer) prefWeightsContainer.style.display = 'block';
   
  updateManualAssignmentsList();
  updateForbiddenAssignmentsList();
  
  if (!document.querySelector('.actual-solve-btn')) {
    const actualSolveBtn = document.createElement('button');
    actualSolveBtn.innerHTML = 'Run Solver<span id="solverSpinner" class="spinner" style="display: none;"></span>';
    actualSolveBtn.className = 'actual-solve-btn';
    document.getElementById('manualAssignmentsList').appendChild(actualSolveBtn);
    
    actualSolveBtn.addEventListener('click', () => {
      const solverSpinner = document.getElementById('solverSpinner');
      solverSpinner.style.display = 'inline-block';
      actualSolveBtn.disabled = true;
      
      const feasibilityCheck = checkManualAssignmentsFeasibility();
      
      if (!feasibilityCheck.feasible) {
        alert('Manual assignments are not feasible:\n\n' + feasibilityCheck.errors.join('\n'));
        solverSpinner.style.display = 'none';
        actualSolveBtn.disabled = false;
        return;
      }
      
      // Use setTimeout to allow the spinner to render before computation starts
      setTimeout(() => {
        try {
          const weights = getPreferenceWeights();
          const model = buildModel(weights);
          
          manualAssignments.forEach(assignment => {
            const student = idToStud[assignment.studentId];
            if (!student) return;
            
            const prefix = student.semester === '1st semester' || student.semester === 'Full academic year' ? 'x_' : 'y_';
            
            const varName = prefix + assignment.studentId + '_' + assignment.uniKey;
            
            model.constraints['manual_' + assignment.studentId] = { equal: 1 };
            
            if (model.variables[varName]) {
              model.variables[varName]['manual_' + assignment.studentId] = 1;
            } else {
              console.error(`Variable ${varName} not found in the model`);
            }
          });

          // Apply forbidden assignment constraints: ensure variable = 0
          forbiddenAssignments.forEach(rule => {
            const student = idToStud[rule.studentId];
            if (!student) return;
            const prefix = (student.semester === '1st semester' || student.semester === 'Full academic year') ? 'x_' : 'y_';
            const varName = prefix + rule.studentId + '_' + rule.uniKey;
            const cName = 'forbid_' + rule.studentId + '_' + rule.uniKey;
            model.constraints[cName] = { equal: 0 };
            if (model.variables[varName]) {
              model.variables[varName][cName] = 1;
            }
          });
          
          const results = solver.Solve(model);
          
          if (results.feasible) {
            displayResults(results);
          } else {
            alert('The model is infeasible with the current manual assignments');
          }
        } catch(e) {
          alert('Error solving the model: ' + e.message);
          console.error(e);
        } finally {
          solverSpinner.style.display = 'none';
          actualSolveBtn.disabled = false;
        }
      }, 100);
    });
  }
}

function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx' || ext === 'xls') {
    return new Promise(resolve => {
      const reader = new FileReader();
      reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: null });
        resolve(json);
      };
      reader.readAsArrayBuffer(file);
    });
  }
  return parseCsvFile(file);
}

function parseCsvFile(file) {
  return new Promise(resolve => {
    Papa.parse(file, {
      header: true,
      dynamicTyping: true,
      complete: res => resolve(res.data)
    });
  });
}

// Build a set of all column names seen across rows
function getHeaderSet(rows) {
  const set = new Set();
  (rows || []).forEach(r => Object.keys(r || {}).forEach(k => set.add(k)));
  return set;
}

// Validate universities raw data for required headers and presence of data rows
function validateUniversitiesRawData(rows) {
  const errors = [];
  if (!Array.isArray(rows) || rows.length === 0) {
    errors.push('File appears empty or could not be parsed.');
    return { ok: false, errors };
  }
  const headers = getHeaderSet(rows);
  const required = ['ID', 'Abbr. of study field'];
  const missing = required.filter(h => !headers.has(h));
  const capacityAny = ['Total #', 'Nr of agreed spots in 1st sem.', 'Nr of agreed spots in 2nd sem.'];
  const hasCapacityHeader = capacityAny.some(h => headers.has(h));
  if (missing.length) {
    errors.push('Missing required column(s): ' + missing.join(', '));
  }
  if (!hasCapacityHeader) {
    errors.push('Need at least one capacity column: Total # or Nr of agreed spots in 1st/2nd sem.');
  }
  const validRows = rows.filter(r => typeof r['ID'] === 'number');
  if (validRows.length === 0) {
    errors.push("No data rows with a numeric 'ID' found.");
  }
  return { ok: errors.length === 0, errors };
}

// Validate applicants raw data for required headers and presence of data rows
function validateApplicantsRawData(rows) {
  const errors = [];
  if (!Array.isArray(rows) || rows.length === 0) {
    errors.push('File appears empty or could not be parsed.');
    return { ok: false, errors };
  }
  const headers = getHeaderSet(rows);
  const required = [
    'ID of application',
    'Abbreviation of study field',
    'Study field',
    'Study level',
    'Semester'
  ];
  const missing = required.filter(h => !headers.has(h));
  if (missing.length) {
    errors.push('Missing required column(s): ' + missing.join(', '));
  }
  const validRows = rows.filter(r => typeof r['ID of application'] === 'number');
  if (validRows.length === 0) {
    errors.push("No data rows with a numeric 'ID of application' found.");
  }
  return { ok: errors.length === 0, errors };
}

// Calculate semester capacities from provided data
function computeSemesters(spots1, spots2, total, maxB, maxM) {
  if (spots1 && !spots2) return { total1: spots1, B1: maxB, M1: maxM, total2: 0, B2: 0, M2: 0 };
  if (spots2 && !spots1) return { total1: 0, B1: 0, M1: 0, total2: spots2, B2: maxB, M2: maxM };
  if (!spots1 && !spots2) return { total1: total, B1: maxB, M1: maxM, total2: total, B2: maxB, M2: maxM };
  return { total1: spots1, B1: maxB, M1: maxM, total2: spots2, B2: maxB, M2: maxM };
}

// Read preference weights from inputs; fallback to defaults [1..6]
function getPreferenceWeights() {
  const defaults = [1,2,3,4,5,6];
  const arr = [];
  for (let i = 1; i <= 6; i++) {
    const el = document.getElementById(`weight_${i}`);
    let v = el ? parseInt(el.value, 10) : defaults[i-1];
    if (isNaN(v) || v < 0) v = defaults[i-1];
    arr.push(v);
  }
  return arr; // index 0 -> choice 1
}

// Build the linear programming model for student-university matching
function buildModel(prefWeights = [1,2,3,4,5,6]) {
  
  const agreementGroups = {};
  const studs1 = [];
  const studs2 = [];
  
  Object.values(idToStud).forEach(stud => {
    if (stud.semester === '1st semester' || stud.semester === 'Full academic year') {
      studs1.push(stud);
    } else {
      studs2.push(stud);
    }
  });
  
  Object.values(idToUni).forEach(uni => {
    if (!uni.agreement_ID) return;
    
    agreementGroups[uni.agreement_ID] = agreementGroups[uni.agreement_ID] || [];
    agreementGroups[uni.agreement_ID].push(uni);
  });
  
  const model = { optimize: 'obj', opType: 'min', constraints: {}, variables: {} };
  const totalStudents = studs1.length + studs2.length;
  const maxWeight = Math.max.apply(null, prefWeights);
  const dummyPenalty = totalStudents * maxWeight + 1;
  // Every student must be assigned to exactly one university
  studs1.forEach(s => { model.constraints['student_f_' + s.ID] = { equal: 1 }; });
  studs2.forEach(s => { model.constraints['student_s_' + s.ID] = { equal: 1 }; });
  // Every university must have its capacity respected
  Object.entries(agreementGroups).forEach(([aid, group]) => {
    const uni0 = group[0];
    model.constraints['cap_ag_f_' + aid] = { max: uni0.total1 };
    model.constraints['cap_ag_s_' + aid] = { max: uni0.total2 };
  });
  // Every university must have its capacity respected for each study level
  Object.values(idToUni).forEach(u => {
    model.constraints['cap_u_f_' + u.key] = { max: u.total1 };
    model.constraints['cap_u_s_' + u.key] = { max: u.total2 };
    model.constraints['cap_uBSc_f_' + u.key] = { max: u.B1 };
    model.constraints['cap_uBSc_s_' + u.key] = { max: u.B2 };
    model.constraints['cap_uMSc_f_' + u.key] = { max: u.M1 };
    model.constraints['cap_uMSc_s_' + u.key] = { max: u.M2 };
    
    // Add faculty constraints per university if this university has faculty constraints
    if (hasUniFacultyConstraints(u)) {
      const faculties = ['BMS', 'EEMCS', 'ET', 'ITC', 'ST', 'UCT'];
      faculties.forEach(faculty => {
        const facultyMax = u[`max_${faculty}`] || 0;        
        const constraintNameF = `cap_faculty_f_${faculty}_${u.key}`;
        model.constraints[constraintNameF] = { max: facultyMax };
        const constraintNameS = `cap_faculty_s_${faculty}_${u.key}`;
        model.constraints[constraintNameS] = { max: facultyMax };
      });
    }
  });
  
  // variables for each arc
  studs1.forEach(s => {
  for (let i=6;i>=1;i--){ 
      const pref = s.all_preferences[i-1];
      if (pref && idToUni[pref]){
        const u = idToUni[pref];
    const wt = prefWeights[i-1] ?? i;
        const name = 'x_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(wt, s.ID, u, 'f', s);
      } else if (pref==null) {
        const keyF = i + '-fictional';
        const u = idToUni[keyF];
        const name = 'x_' + s.ID + '_' + u.key;
    model.variables[name] = buildVar(prefWeights[i-1] ?? i, s.ID, u, 'f', s);
      }
    }
    // dummy, if no option can be assigned
    const u = idToUni['dummy-dummy'];
    const name = 'x_' + s.ID + '_' + u.key;
  model.variables[name] = buildVar(dummyPenalty, s.ID, u, 'f', s);
  });
  studs2.forEach(s => {
    for (let i=6;i>=1;i--){
      const pref = s.all_preferences[i-1];
      if (pref && idToUni[pref]){
        const u = idToUni[pref];
    const wt = prefWeights[i-1] ?? i;
        const name = 'y_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(wt, s.ID, u, 's', s);
      } else if (pref==null) {
        const keyF = i + '-fictional';
        const u = idToUni[keyF];
        const name = 'y_' + s.ID + '_' + u.key;
    model.variables[name] = buildVar(prefWeights[i-1] ?? i, s.ID, u, 's', s);
      }
    }
    // dummy, if no option can be assigned
    const u = idToUni['dummy-dummy'];
    const name = 'y_' + s.ID + '_' + u.key;
  model.variables[name] = buildVar(dummyPenalty, s.ID, u, 's', s);
  });
  
  return model;
}

function buildVar(weight, studID, uni, sem, stud) {
  const v = { obj: weight };
  v['student_' + sem + '_' + studID] = 1;
  v['cap_ag_' + sem + '_' + uni.agreement_ID] = 1;
  v['cap_u_' + sem + '_' + uni.key] = 1;
  v['cap_u' + stud.study_level + '_' + sem + '_' + uni.key] = 1;
  
  if (hasUniFacultyConstraints(uni)) {
    const studentFaculty = getFacultyForStudyField(stud.study_field);
    if (studentFaculty) {
      const constraintName = `cap_faculty_${sem}_${studentFaculty}_${uni.key}`;
      v[constraintName] = 1;
    }
  }
  
  return v;
}

function displayResults(results) {
  const div = document.getElementById('results');
  div.innerHTML = '';
  
  const choiceCounts = Array(6).fill(0);
  const assignments = [];
  
  const matchedStudentIds = new Set();
  
  Object.keys(results).forEach(k => {
    if ((k.startsWith('x_') || k.startsWith('y_')) && results[k] > 0.9) {
      const parts = k.split('_');
      const stud = parts[1];
      const uniKey = parts.slice(2).join('_');
      const prefs = window.studentPrefsMap[stud] || [];
      const idx = prefs.indexOf(uniKey);
      
      // Only count as "matched" if not assigned to dummy or fictional university
      if (!uniKey.includes('-dummy') && !uniKey.includes('-fictional')) {
        matchedStudentIds.add(stud);
      }
      
      if (idx >= 0 && idx < 6) {
        choiceCounts[idx]++;
        
        
        const student = idToStud[stud];
        const fullName = student && student.fullName ? student.fullName : '';
        
        const uni = idToUni[uniKey];
        const partnerName = uni && uni.partner_name ? uni.partner_name : '';
        
        assignments.push({ 
          stud, 
          uniKey, 
          pref: idx + 1,
          fullName: fullName,
          partnerName: partnerName,
          student: student,
          uni: uni
        });
      }
    }
  });

  assignments.sort((a, b) => parseInt(a.stud) - parseInt(b.stud));
  const summaryHeader = document.createElement('h2');
  summaryHeader.textContent = 'Summary';
  div.appendChild(summaryHeader);
  const ul = document.createElement('ul');
  ul.className = 'summary-list';
  choiceCounts.forEach((count, i) => {
    const li = document.createElement('li');
    li.textContent = `${i+1} choice: ${count}`;
    ul.appendChild(li);
  });
  
  
  const totalStudents = Object.keys(idToStud).length;
  const notMatchedStudents = totalStudents - matchedStudentIds.size;
  
  const unmatchedLi = document.createElement('li');
  unmatchedLi.innerHTML = `<strong>Not matched students: ${notMatchedStudents}</strong>`;
  ul.appendChild(unmatchedLi);
  
  const totalStudentsLi = document.createElement('li');
  totalStudentsLi.innerHTML = `<strong>Total students: ${totalStudents}</strong>`;
  ul.appendChild(totalStudentsLi);

  let totalPreferenceScore = 0;  
  assignments.forEach(assignment => {
    totalPreferenceScore += assignment.pref;
  });  
  const unmatchedStudents = Object.values(idToStud).filter(student => 
    !matchedStudentIds.has(student.ID.toString())
  );
  unmatchedStudents.forEach(student => {
    const validPrefs = (student.all_preferences || []).filter(pref => pref !== null).length;
    totalPreferenceScore += validPrefs + 1;
  });
  const averagePreference = totalStudents > 0 ? (totalPreferenceScore / totalStudents) : 0;
  const averageLi = document.createElement('li');
  averageLi.innerHTML = `<strong>Average preference achieved: ${averagePreference.toFixed(3)}</strong>`;
  ul.appendChild(averageLi);
  
  div.appendChild(ul);
  
  if (notMatchedStudents > 0) {
    const unmatchedStudentsList = document.createElement('div');
    unmatchedStudentsList.className = 'unmatched-students';
    unmatchedStudentsList.innerHTML = '<h3>Unmatched Students</h3>';
    
    const listElement = document.createElement('ul');
    listElement.className = 'unmatched-students-list';
    
    const unmatchedStudents = Object.values(idToStud).filter(student => 
      !matchedStudentIds.has(student.ID.toString())
    );
    
    unmatchedStudents.sort((a, b) => a.ID - b.ID);
    
    unmatchedStudents.forEach(student => {
      const li = document.createElement('li');
      
      const validPrefCount = (student.all_preferences || []).filter(pref => pref !== null).length; 
      
      if (student.fullName) {
        li.textContent = `${student.ID} (${student.fullName}) - ${validPrefCount} preferences`;
      } else {
        li.textContent = `${student.ID} - ${validPrefCount} preferences`;
      }
      listElement.appendChild(li);
    });
    
    unmatchedStudentsList.appendChild(listElement);
    div.appendChild(unmatchedStudentsList);
  }
  const assignHeader = document.createElement('h2');
  assignHeader.textContent = 'Assignments';
  div.appendChild(assignHeader);

  // Add export to CSV button
  const exportButton = document.createElement('button');
  exportButton.textContent = 'Export to CSV';
  exportButton.className = 'export-csv-btn';
  exportButton.style.marginBottom = '15px';
  exportButton.addEventListener('click', () => exportAssignmentsToCSV(assignments));
  div.appendChild(exportButton);

  const assignTable = document.createElement('table');
  // enable sorting by clicking column headers
  const headers = ['Student','University','Preference'];
  const headerRow = document.createElement('tr');
  headers.forEach(text => {
    const th = document.createElement('th');
    th.textContent = text;
    th.style.cursor = 'pointer';
    headerRow.appendChild(th);
  });
  
  Array.from(headerRow.children).forEach((th, colIndex) => {
    th.addEventListener('click', () => {
      const rows = Array.from(assignTable.querySelectorAll('tr')).slice(1);
      const asc = !th.asc;
      rows.sort((a, b) => {
        let va = a.cells[colIndex].textContent;
        let vb = b.cells[colIndex].textContent;
        const na = parseFloat(va);
        const nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) {
          return asc ? na - nb : nb - na;
        }
        return asc ? va.localeCompare(vb) : vb.localeCompare(va);
      });
      rows.forEach(r => assignTable.appendChild(r));
      th.asc = asc;
    });
  });
  assignTable.appendChild(headerRow);
  
  assignments.forEach(a => {
    const row = document.createElement('tr');
    
    const studentCell = document.createElement('td');
    if (a.fullName) {
      studentCell.textContent = `${a.stud} (${a.fullName})`;
    } else {
      studentCell.textContent = a.stud;
    }
    row.appendChild(studentCell);
    
    const uniCell = document.createElement('td');
    if (a.partnerName) {
      uniCell.textContent = `${a.uniKey} (${a.partnerName})`;
    } else {
      uniCell.textContent = a.uniKey;
    }
    row.appendChild(uniCell);
    
    const prefCell = document.createElement('td');
    prefCell.textContent = a.pref;
    row.appendChild(prefCell);
    
    assignTable.appendChild(row);
  });
  div.appendChild(assignTable);
}

function exportAssignmentsToCSV(assignments) {
  const csvRows = [];
  
  csvRows.push([
    'Application ID',
    'Student number',
    'Student name',
    'University', 
    'Agreement-ID',
    'Semester',
    'Preference',
    'Programme',
    'Study level',
    'Faculty',
    'Agreement-type'
  ]);
  
  assignments.forEach(assignment => {
    const student = assignment.student;
    const uni = assignment.uni;
    
    if (!student || !uni) return;
    
    const semesterNum = (student.semester === '1st semester' || student.semester === 'Full academic year') ? 1 : 2;    
    const faculty = getFacultyForStudyField(student.study_field) || '';

    const applicationId = student.ID || assignment.stud;
    const studentNumber = student.student_number || '';
    const studentName = student.fullName || '';

    csvRows.push([
      applicationId,
      studentNumber,
      studentName,
      uni.partner_name || assignment.uniKey,
      uni.agreement_ID || '',
      semesterNum,
      assignment.pref,
      student.study_field || '',
      student.study_level || '',
      faculty,
      uni.agreement_type || ''
    ]);
  });
  
  const csvContent = csvRows.map(row => 
    row.map(field => {
      const fieldStr = String(field || '');
      if (fieldStr.includes(',') || fieldStr.includes('"') || fieldStr.includes('\n')) {
        return '"' + fieldStr.replace(/"/g, '""') + '"';
      }
      return fieldStr;
    }).join(',')
  ).join('\n');
  
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  
  if (navigator.msSaveBlob) {
    // IE 10+
    navigator.msSaveBlob(blob, 'assignment_results.csv');
  } else {
    // Other browsers
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'assignment_results.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }
}

// Initialize the application when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', function() {
  const universitiesFileInput = document.getElementById('universitiesFile');
  const applicantsFileInput = document.getElementById('applicantsFile');
  const resetWeightsBtn = document.getElementById('resetWeightsBtn');

  
  setupStudentSearch();
  if (resetWeightsBtn) {
    resetWeightsBtn.addEventListener('click', () => {
      for (let i=1;i<=6;i++) {
        const el = document.getElementById(`weight_${i}`);
        if (el) el.value = String(i);
      }
    });
  }
  
  universitiesFileInput.addEventListener('change', function() {
    const file = this.files[0];
    if (file) {
      parseFile(file).then(data => {
        const v = validateUniversitiesRawData(data);
        if (!v.ok) {
          alert('Universities file issues:\n\n' + v.errors.map(e => '- ' + e).join('\n'));
          universitiesData = null;
          universitiesLoaded = false;
          return;
        }
        universitiesData = data;
        universitiesLoaded = true;
        processUniversitiesData();
        processFiles();
      }).catch(err => {
        console.error('Error loading universities file:', err);
        alert('Error loading universities file: ' + err.message);
      });
    }
  });
  
  applicantsFileInput.addEventListener('change', function() {
    const file = this.files[0];
    if (file) {
      parseFile(file).then(data => {
        const v = validateApplicantsRawData(data);
        if (!v.ok) {
          alert('Applicants file issues:\n\n' + v.errors.map(e => '- ' + e).join('\n'));
          applicantsData = null;
          applicantsLoaded = false;
          return;
        }
        applicantsData = data;
        applicantsLoaded = true;
        processApplicantsData();
        processFiles();
      }).catch(err => {
        console.error('Error loading applicants file:', err);
        alert('Error loading applicants file: ' + err.message);
      });
    }
  });
});

/**
 * Cleans up student preferences by:
 * - Removing preferences for non-existent universities
 * - Eliminating duplicate preferences
 * - Padding with nulls to maintain consistent array size
 */
function cleanupStudentPreferences() {
  Object.values(idToStud).forEach(student => {
    if (!student.all_preferences || !Array.isArray(student.all_preferences)) {
      return;
    }
    
    const validPrefs = [];
    const seenPrefs = new Set();
    
    student.all_preferences.forEach(pref => {
      if (!pref || seenPrefs.has(pref)) {
        return;
      }
      
      if (idToUni[pref]) {
        validPrefs.push(pref);
        seenPrefs.add(pref);
      }
    });
    
    while (validPrefs.length < 6) {
      validPrefs.push(null);
    }
    
    student.all_preferences = validPrefs;
    studentPrefsMap[student.ID] = validPrefs;
  });
}
