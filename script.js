// Global variables for data storage
let idToUni = {};
let idToStud = {};
let manualAssignments = [];
let selectedStudentForAssignment = null;

// Validate Exchange I Factor input
function validateInput(input) {
  let value = parseInt(input.value);
  
  if (isNaN(value) || value < 1) {
    input.value = 1;
  } else {
    input.value = Math.floor(value);
  }
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

// Setup student search functionality
function setupStudentSearch() {
  const searchInput = document.getElementById('studentSearch');
  const resultsDiv = document.getElementById('studentResults');
  const assignmentContainer = document.querySelector('.assignment-container');
  const selectedStudentNameSpan = document.getElementById('selectedStudentName');
  const preferenceOptionsDiv = document.getElementById('preferenceOptions');
  const addBtn = document.getElementById('addAssignmentBtn');
  const cancelBtn = document.getElementById('cancelAssignmentBtn');
  
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
    comments_agreement: r['Comments regarding the agreement (in portal)']
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
      total1: 1000, 
      B1: 1000, 
      M1: 1000, 
      total2: 1000, 
      B2: 1000, 
      M2: 1000 
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
    const hasFirstName = Object.keys(r).some(key => key === 'First name');
    const hasLastName = Object.keys(r).some(key => key.startsWith('Last name'));
    const firstName = hasFirstName ? (r['First name'] || '') : '';
    const lastNameKey = hasLastName ? Object.keys(r).find(key => key.startsWith('Last name')) : null;
    const lastName = lastNameKey ? r[lastNameKey] || '' : '';
    const fullName = [firstName, lastName].filter(Boolean).join(' ');
    return {
      ID: r['ID of application'],
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
  const exchangeIFactorContainer = document.getElementById('exchangeIFactorContainer');
  
  // If both files aren't loaded yet, wait
  if (!universitiesLoaded || !applicantsLoaded) {
    return;
  }
  
  resultsDiv.innerHTML = '';
  
  // Clean up student preferences - remove non-existing options and duplicates, shift remaining up
  cleanupStudentPreferences();
  
  // Show manual assignment section and exchange I factor after files are loaded
  manualAssignmentDiv.style.display = 'block';
  exchangeIFactorContainer.style.display = 'block';
  
  updateManualAssignmentsList();
  
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
          const exchangeI_factor = parseInt(document.getElementById('exchangeIFactor').value) || 1;
          
          const model = buildModel(exchangeI_factor);
          
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

// Calculate semester capacities from provided data
function computeSemesters(spots1, spots2, total, maxB, maxM) {
  if (spots1 && !spots2) return { total1: spots1, B1: maxB, M1: maxM, total2: 0, B2: 0, M2: 0 };
  if (spots2 && !spots1) return { total1: 0, B1: 0, M1: 0, total2: spots2, B2: maxB, M2: maxM };
  if (!spots1 && !spots2) return { total1: total, B1: maxB, M1: maxM, total2: total, B2: maxB, M2: maxM };
  return { total1: spots1, B1: maxB, M1: maxM, total2: spots2, B2: maxB, M2: maxM };
}

// Build the linear programming model for student-university matching
function buildModel(exchangeI_factor = 1) {
  exchangeI_factor = (!isNaN(exchangeI_factor) && exchangeI_factor >= 1) ? Math.floor(exchangeI_factor) : 1;
  
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
  });
  
  // variables for each arc
  studs1.forEach(s => {
    for (let i=6;i>=1;i--){ 
      const pref = s.all_preferences[i-1];
      if (pref && idToUni[pref]){
        const u = idToUni[pref];
        const wt = (u.agreement_type==='Exchange-I'?i:exchangeI_factor*i);
        const name = 'x_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(wt, s.ID, u, 'f', s);
      } else if (pref==null) {
        const keyF = i + '-fictional';
        const u = idToUni[keyF];
        const name = 'x_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(exchangeI_factor*i, s.ID, u, 'f', s);
      }
    }
    // dummy, if no option can be assigned
    const u = idToUni['dummy-dummy'];
    const name = 'x_' + s.ID + '_' + u.key;
    model.variables[name] = buildVar(1000*exchangeI_factor, s.ID, u, 'f', s);
  });
    studs2.forEach(s => {
    for (let i=6;i>=1;i--){
      const pref = s.all_preferences[i-1];
      if (pref && idToUni[pref]){
        const u = idToUni[pref];
        const wt = (u.agreement_type==='Exchange-I'?i:exchangeI_factor*i);
        const name = 'y_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(wt, s.ID, u, 's', s);
      } else if (pref==null) {
        const keyF = i + '-fictional';
        const u = idToUni[keyF];
        const name = 'y_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(exchangeI_factor*i, s.ID, u, 's', s);
      }
    }
    // dummy, if no option can be assigned
    const u = idToUni['dummy-dummy'];
    const name = 'y_' + s.ID + '_' + u.key;
    model.variables[name] = buildVar(1000*exchangeI_factor, s.ID, u, 's', s);
  });
  
  return model;
}

function buildVar(weight, studID, uni, sem, stud) {
  const v = { obj: weight };
  v['student_' + sem + '_' + studID] = 1;
  v['cap_ag_' + sem + '_' + uni.agreement_ID] = 1;
  v['cap_u_' + sem + '_' + uni.key] = 1;
  v['cap_u' + stud.study_level + '_' + sem + '_' + uni.key] = 1;
  return v;
}

function displayResults(results) {
  const div = document.getElementById('results');
  div.innerHTML = '';
  
  const choiceCounts = Array(6).fill(0);
  const assignments = [];
  let exchangeICount = 0;
  
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
        
        const isExchangeI = !uniKey.includes('-dummy') && !uniKey.includes('-fictional') && 
                            idToUni[uniKey] && idToUni[uniKey].agreement_type === 'Exchange-I';
        if (isExchangeI) {
          exchangeICount++;
        }
        
        const student = idToStud[stud];
        const fullName = student && student.fullName ? student.fullName : '';
        
        const uni = idToUni[uniKey];
        const partnerName = uni && uni.partner_name ? uni.partner_name : '';
        
        assignments.push({ 
          stud, 
          uniKey, 
          pref: idx + 1,
          fullName: fullName,
          partnerName: partnerName
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
  
  const exchangeILi = document.createElement('li');
  exchangeILi.innerHTML = `<strong>Exchange I slots filled: ${exchangeICount}</strong>`;
  ul.appendChild(exchangeILi);
  
  const totalStudents = Object.keys(idToStud).length;
  const notMatchedStudents = totalStudents - matchedStudentIds.size;
  
  const unmatchedLi = document.createElement('li');
  unmatchedLi.innerHTML = `<strong>Not matched students: ${notMatchedStudents}</strong>`;
  ul.appendChild(unmatchedLi);
  
  const totalStudentsLi = document.createElement('li');
  totalStudentsLi.innerHTML = `<strong>Total students: ${totalStudents}</strong>`;
  ul.appendChild(totalStudentsLi);
  
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

// Initialize the application when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', function() {
  const exchangeIFactorInput = document.getElementById('exchangeIFactor');
  const universitiesFileInput = document.getElementById('universitiesFile');
  const applicantsFileInput = document.getElementById('applicantsFile');
  
  validateInput(exchangeIFactorInput);
  
  exchangeIFactorInput.addEventListener('blur', function() {
    validateInput(this);
  });
  
  setupStudentSearch();
  
  universitiesFileInput.addEventListener('change', function() {
    const file = this.files[0];
    if (file) {
      parseFile(file).then(data => {
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
