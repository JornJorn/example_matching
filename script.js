document.getElementById('solveBtn').addEventListener('click', () => {
  const solveBtn = document.getElementById('solveBtn');
  const spinner = document.getElementById('spinner');
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  solveBtn.disabled = true;
  spinner.style.display = 'inline-block';
  const uniFile = document.getElementById('universitiesFile').files[0];
  const appFile = document.getElementById('applicantsFile').files[0];
  if (!uniFile || !appFile) {
    alert('Please select both CSV files.');
    solveBtn.disabled = false;
    spinner.style.display = 'none';
    return;
  }
  Promise.all([parseFile(uniFile), parseFile(appFile)])
    .then(([uniDataRaw, appDataRaw]) => {
      // remove footer summary rows added by Excel
      const uniData = uniDataRaw.filter(r => typeof r['ID'] === 'number');
      const appData = appDataRaw.filter(r => typeof r['ID of application'] === 'number');
      // normalize university data fields
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
      // normalize applicant data fields
      const students = appData.map(r => ({
        ID: r['ID of application'],
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
      }));
      // build a map of student IDs to their preference keys
      const studentPrefsMap = {};
      students.forEach(r => {
        const prefs = [];
        [r.preference_1, r.preference_2, r.preference_3, r.preference_4, r.preference_5, r.preference_6].forEach((p, i) => {
          prefs.push(p ? p + '_' + r.abbreviation : null);
        });
        studentPrefsMap[r.ID] = prefs;
      });
      window.studentPrefsMap = studentPrefsMap;
      
      const model = buildModel(unis, students);
      const results = solver.Solve(model);
      displayResults(results);
    })
    .catch(err => {
      console.error(err);
      alert('An error occurred: ' + err.message);
    })
    .finally(() => {
      solveBtn.disabled = false;
      spinner.style.display = 'none';
    });
});

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

function computeSemesters(spots1, spots2, total, maxB, maxM) {
  if (spots1 && !spots2) return { total1: spots1, B1: maxB, M1: maxM, total2: 0, B2: 0, M2: 0 };
  if (spots2 && !spots1) return { total1: 0, B1: 0, M1: 0, total2: spots2, B2: maxB, M2: maxM };
  if (!spots1 && !spots2) return { total1: total, B1: maxB, M1: maxM, total2: total, B2: maxB, M2: maxM };
  return { total1: spots1, B1: maxB, M1: maxM, total2: spots2, B2: maxB, M2: maxM };
}

function buildModel(unisRaw, studsRaw) {
  const exchangeI_factor = 1;
  // prepare universities
  const idToUni = {};
  const agreementGroups = {};
  unisRaw.forEach(r => {
    const sem = computeSemesters(r.spots_first_s, r.spots_second_s, r.total_places, r.max_BCs, r.max_MCs);
    const uni = Object.assign({}, r, { ...sem });
    const key = r.agreement_ID + '_' + r.study_abbr;
    uni.key = key;
    idToUni[key] = uni;
    // group by agreement
    agreementGroups[r.agreement_ID] = agreementGroups[r.agreement_ID] || [];
    agreementGroups[r.agreement_ID].push(uni);
  });
  // add dummy and fictional
  ['dummy'].concat(['1','2','3','4','5','6']).forEach(id => {
    const key = id + (id==='dummy'?'-dummy':'-fictional');
    const uni = { agreement_ID: id, study_abbr: id, key, total1:1000, B1:1000, M1:1000, total2:1000, B2:1000, M2:1000 };
    idToUni[key] = uni;
    if (id!=='dummy') agreementGroups[id] = agreementGroups[id]||[], agreementGroups[id].push(uni);
  });
  // prepare students
  const studs1 = [];
  const studs2 = [];
  const idToStud = {};
  studsRaw.forEach(r => {
    if (!r.ID) return;
    const prefs = [];
    [r.preference_1, r.preference_2, r.preference_3, r.preference_4, r.preference_5, r.preference_6].forEach((p,i) => {
      prefs.push(p ? p + '_' + r.abbreviation : null);
    });
    const stud = Object.assign({}, r, { all_preferences: prefs });
    idToStud[r.ID] = stud;
    if (r.semester === '1st semester' || r.semester === 'Full academic year') studs1.push(stud);
    else studs2.push(stud);
  });
  // model
  const model = { optimize: 'obj', opType: 'min', constraints: {}, variables: {} };
  // student constraints
  studs1.forEach(s => { model.constraints['student_f_' + s.ID] = { equal: 1 }; });
  studs2.forEach(s => { model.constraints['student_s_' + s.ID] = { equal: 1 }; });
  // group caps
  Object.entries(agreementGroups).forEach(([aid, group]) => {
    const uni0 = group[0];
    model.constraints['cap_ag_f_' + aid] = { max: uni0.total1 };
    model.constraints['cap_ag_s_' + aid] = { max: uni0.total2 };
  });
  // uni caps
  Object.values(idToUni).forEach(u => {
    model.constraints['cap_u_f_' + u.key] = { max: u.total1 };
    model.constraints['cap_u_s_' + u.key] = { max: u.total2 };
    model.constraints['cap_uBSc_f_' + u.key] = { max: u.B1 };
    model.constraints['cap_uBSc_s_' + u.key] = { max: u.B2 };
    model.constraints['cap_uMSc_f_' + u.key] = { max: u.M1 };
    model.constraints['cap_uMSc_s_' + u.key] = { max: u.M2 };
  });
  // variables for each arc
  // first sem
  studs1.forEach(s => {
    for (let i=1;i<=6;i++){
      const pref = s.all_preferences[i-1];
      if (pref && idToUni[pref]){
        const u = idToUni[pref];
        const wt = (u.agreement_type==='Exchange-I'?i:exchangeI_factor*i);
        const name = 'x_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(name, wt, s.ID, u, 'f', s);
      } else if (pref==null) {
        const keyF = i + '-fictional';
        const u = idToUni[keyF];
        const name = 'x_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(name, exchangeI_factor*i, s.ID, u, 'f', s);
      }
    }
    // dummy
    const u = idToUni['dummy-dummy'];
    const name = 'x_' + s.ID + '_' + u.key;
    model.variables[name] = buildVar(name,1000*exchangeI_factor, s.ID, u, 'f', s);
  });
  // second sem
  studs2.forEach(s => {
    for (let i=1;i<=6;i++){
      const pref = s.all_preferences[i-1];
      if (pref && idToUni[pref]){
        const u = idToUni[pref];
        const wt = (u.agreement_type==='Exchange-I'?i:exchangeI_factor*i);
        const name = 'y_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(name, wt, s.ID, u, 's', s);
      } else if (pref==null) {
        const keyF = i + '-fictional';
        const u = idToUni[keyF];
        const name = 'y_' + s.ID + '_' + u.key;
        model.variables[name] = buildVar(name, exchangeI_factor*i, s.ID, u, 's', s);
      }
    }
    const u = idToUni['dummy-dummy'];
    const name = 'y_' + s.ID + '_' + u.key;
    model.variables[name] = buildVar(name,1000*exchangeI_factor, s.ID, u, 's', s);
  });
  return model;
}

function buildVar(varName, weight, studID, uni, sem, stud) {
  const v = { obj: weight };
  v['student_' + sem + '_' + studID] = 1;
  v['cap_ag_' + sem + '_' + uni.agreement_ID] = 1;
  v['cap_u_' + sem + '_' + uni.key] = 1;
  v['cap_u' + stud.study_level + '_' + sem + '_' + uni.key] = 1;
  // also contribute to cap_uBSc or cap_uMSc
  return v;
}

function displayResults(results) {
  const div = document.getElementById('results');
  div.innerHTML = '';
  // compute preference counts and assignments
  const choiceCounts = Array(6).fill(0);
  const assignments = [];
  Object.keys(results).forEach(k => {
    if ((k.startsWith('x_') || k.startsWith('y_')) && results[k] > 0.9) {
      const parts = k.split('_');
      const stud = parts[1];
      const uniKey = parts.slice(2).join('_');
      const prefs = window.studentPrefsMap[stud] || [];
      const idx = prefs.indexOf(uniKey);
      if (idx >= 0 && idx < 6) {
        choiceCounts[idx]++;
        assignments.push({ stud, uniKey, pref: idx + 1 });
      }
    }
  });
  // display summary
  const summaryHeader = document.createElement('h2');
  summaryHeader.textContent = 'Summary';
  div.appendChild(summaryHeader);
  const ol = document.createElement('ol');
  choiceCounts.forEach((count, i) => {
    const li = document.createElement('li');
    li.textContent = `${i+1} choice: ${count}`;
    ol.appendChild(li);
  });
  div.appendChild(ol);
  // display assignments
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
  // attach click handlers for sorting
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
    [a.stud, a.uniKey, a.pref].forEach(v => { const td = document.createElement('td'); td.textContent = v; row.appendChild(td); });
    assignTable.appendChild(row);
  });
  div.appendChild(assignTable);
}
