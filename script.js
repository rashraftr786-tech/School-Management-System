let allStudents = [];
let currentEditIndex = null;

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processBtn').onclick = processExcel;
    document.getElementById('searchInput').onkeyup = filterData;
    document.getElementById('classFilter').onchange = filterData;
    document.getElementById('downloadFullPDF').onclick = () => generateDetailedReport();
    
    document.querySelector('.close').onclick = () => document.getElementById('photoModal').style.display = 'none';
    document.getElementById('confirmUpload').onclick = savePhoto;
});

function processExcel() {
    const fileInput = document.getElementById('excelFile');
    if (!fileInput.files.length) return alert("Please select a file!");

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        allStudents = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        populateClassFilter();
        renderTable(allStudents);
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    if (!tbody) return;
    
    tbody.innerHTML = data.map((s, index) => `
        <tr>
            <td><img src="https://via.placeholder.com/45" class="student-photo" id="img-${index}" alt="Student" style="width:45px;height:45px;"></td>
            <td>${s['Student Name'] || 'N/A'}</td>
            <td>${s["Father's Name"] || 'N/A'}</td>
            <td>${s['Roll No'] || 'N/A'}</td>
            <td>${s['Class'] || 'N/A'}</td>
            <td>${s['Overall Percentage'] || '0'}%</td>
            <td><span class="grade-badge">${s['Overall Grade'] || 'N/A'}</span></td>
            <td>${s['Position in Class'] || 'N/A'}</td>
            <td>
                <button onclick="openModal(${index})" class="btn-sm btn-primary">Photo</button>
                <button onclick="generateDetailedReport(${index})" class="btn-sm btn-secondary">Report</button>
            </td>
        </tr>
    `).join('');
}

function filterData() {
    const search = document.getElementById('searchInput').value.toLowerCase();
    const classVal = document.getElementById('classFilter').value;

    const filtered = allStudents.filter(s => {
        const matchesName = String(s['Student Name'] || '').toLowerCase().includes(search);
        const matchesClass = classVal === "" || String(s['Class'] || '') === classVal;
        return matchesName && matchesClass;
    });
    renderTable(filtered);
}

function populateClassFilter() {
    const select = document.getElementById('classFilter');
    const classes = [...new Set(allStudents.map(s => s['Class']).filter(c => c))];
    select.innerHTML = '<option value="">All Classes</option>' + 
        classes.map(c => `<option value="${c}">${c}</option>`).join('');
}

function openModal(index) {
    currentEditIndex = index;
    document.getElementById('photoModal').style.display = 'block';
}

function savePhoto() {
    const file = document.getElementById('photoFile').files[0];
    if (file && currentEditIndex !== null) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = document.getElementById(`img-${currentEditIndex}`);
            if (img) img.src = e.target.result;
            document.getElementById('photoModal').style.display = 'none';
        };
        reader.readAsDataURL(file);
    }
}

// --- REPLACE YOUR EXISTING FUNCTION WITH THIS ONE ---

async function generateDetailedReport(index) {
    if (!window.html2pdf) return alert("html2pdf library not loaded! See instructions.");
    
    const s = allStudents[index];
    if (!s) return alert("Error: Student data not found.");

    // Generate specific styling for the grade badge
    const getGradeClass = (grade) => {
        if (!grade || grade === 'N/A') return 'grade-na';
        const g = String(grade).toUpperCase();
        if (g.includes('A')) return 'grade-a';
        if (g.includes('B')) return 'grade-b';
        if (g === 'F') return 'grade-f';
        return 'grade-default';
    };

    const studentPhotoId = `img-${index}`;
    const studentPhotoSrc = document.getElementById(studentPhotoId)?.src || "https://via.placeholder.com/80?text=Photo";

    // --- Start of New HTML/CSS Template for the PDF ---
    const element = document.createElement('div');
    element.innerHTML = `
        <style>
            /* Reset for PDF rendering to ensure standard margins are removed */
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { background-color: #f7f9fc; }

            /* Main Page Layout */
            #report-card-container {
                width: 210mm; /* A4 Width */
                min-height: 297mm; /* A4 Height */
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: white;
                color: #333;
                padding: 15mm;
                line-height: 1.4;
                position: relative;
                overflow: hidden;
            }

            /* Main Colorful Decorative Header Bar */
            .main-header {
                display: flex;
                align-items: center;
                gap: 20px;
                background: linear-gradient(135deg, #6a11cb 0%, #2575fc 100%);
                color: white;
                padding: 25px 35px;
                border-radius: 15px;
                box-shadow: 0 10px 20px rgba(0,0,0,0.1);
                margin-bottom: 25px;
            }

            .school-logo {
                width: 80px;
                height: 80px;
                border-radius: 15px;
                background: white;
                padding: 10px;
                object-fit: cover;
                border: 3px solid #eee;
            }

            .header-text h1 { font-size: 26px; font-weight: 800; letter-spacing: 1px; }
            .header-text p { font-size: 14px; opacity: 0.9; }

            /* Student Profile Section */
            .profile-section {
                display: flex;
                background: white;
                padding: 20px;
                border-radius: 15px;
                border: 1px solid #e1e8f0;
                margin-bottom: 25px;
                gap: 25px;
                align-items: center;
            }

            .profile-photo {
                width: 90px;
                height: 90px;
                border-radius: 20px;
                object-fit: cover;
                border: 4px solid #2575fc;
                box-shadow: 0 5px 15px rgba(37,117,252,0.2);
            }

            .profile-details {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 10px;
                flex: 1;
            }

            .detail-row {
                font-size: 13px;
                color: #555;
            }

            .detail-row b {
                color: #111;
                width: 120px;
                display: inline-block;
            }

            /* Colorful Marks Table */
            .marks-section h2 {
                color: #2575fc;
                font-size: 18px;
                border-bottom: 3px solid #6a11cb;
                display: inline-block;
                padding-bottom: 5px;
                margin-bottom: 15px;
            }

            .marks-table {
                width: 100%;
                border-collapse: collapse;
                border-radius: 10px;
                overflow: hidden;
                box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            }

            .marks-table thead {
                background: linear-gradient(to right, #6a11cb, #2575fc);
                color: white;
            }

            .marks-table th {
                padding: 14px;
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: 1px;
                font-weight: 700;
            }

            .marks-table td {
                padding: 14px;
                border-bottom: 1px solid #f1f1f1;
                font-size: 13px;
                text-align: center;
            }

            .marks-table td:first-child {
                text-align: left;
                font-weight: bold;
                color: #2575fc;
                padding-left: 20px;
            }

            .marks-table tr:last-child td { border-bottom: none; }

            /* Style for the Grade Badges */
            .grade-badge {
                padding: 5px 12px;
                border-radius: 12px;
                color: white;
                font-weight: bold;
                font-size: 12px;
                text-transform: uppercase;
            }
            .grade-a { background: #28a745; }
            .grade-b { background: #ffc107; color: #333; }
            .grade-f { background: #dc3545; }
            .grade-na { background: #999; }
            .grade-default { background: #2575fc; }

            /* Results Footer Summary */
            .results-footer {
                display: flex;
                background: #f1f7ff;
                padding: 20px 30px;
                border-radius: 15px;
                border-left: 10px solid #2575fc;
                margin-top: 30px;
                gap: 40px;
            }

            .result-item {
                font-size: 16px;
                font-weight: 700;
                color: #111;
            }
            .result-item span {
                font-size: 14px;
                font-weight: 400;
                color: #555;
                display: block;
                margin-bottom: 3px;
            }
        </style>

        <div id="report-card-container">
            <header class="main-header">
                <img src="${studentPhotoSrc}" class="school-logo" alt="School Logo">
                <div class="header-text">
                    <h1>GOVERNMENT PRIMARY SCHOOL</h1>
                    <p>PETHGAM WAGOORA 193109</p>
                    <p>OFFICIAL STUDENT REPORT CARD</p>
                </div>
            </header>

            <section class="profile-section">
                <img src="${studentPhotoSrc}" class="profile-photo" alt="Student Photo">
                <div class="profile-details">
                    <div class="detail-row"><b>Student Name:</b> ${s['Student Name'] || 'N/A'}</div>
                    <div class="detail-row"><b>Adm No:</b> ${s['Adm No'] || 'N/A'}</div>
                    <div class="detail-row"><b>Father\'s Name:</b> ${s["Father's Name"] || 'N/A'}</div>
                    <div class="detail-row"><b>Roll No:</b> ${s['Roll No'] || 'N/A'}</div>
                    <div class="detail-row"><b>Mother\'s Name:</b> ${s['Mother\'s Name'] || 'N/A'}</div>
                    <div class="detail-row"><b>PEN No:</b> ${s['PEN No'] || 'N/A'}</div>
                    <div class="detail-row"><b>Date of Birth:</b> ${s['D.O.B'] || 'N/A'}</div>
                    <div class="detail-row"><b>Academic Session:</b> ${s['Academic Session'] || 'N/A'}</div>
                </div>
            </section>

            <section class="marks-section">
                <h2>ACADEMIC PERFORMANCE</h2>
                <table class="marks-table">
                    <thead>
                        <tr>
                            <th>Subject</th>
                            <th>FA1</th><th>FA2</th><th>FA3</th><th>FA4</th><th>FA5</th><th>FA6</th>
                            <th>CCA</th><th>SA</th><th>Total</th><th>Grade</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${['English', 'Math', 'EVS', 'Urdu', 'Kashmiri'].map(sub => `
                            <tr>
                                <td>${sub}</td>
                                <td>${s[sub + ' FA1'] || '0'}</td>
                                <td>${s[sub + ' FA2'] || '0'}</td>
                                <td>${s[sub + ' FA3'] || '0'}</td>
                                <td>${s[sub + ' FA4'] || '0'}</td>
                                <td>${s[sub + ' FA5'] || '0'}</td>
                                <td>${s[sub + ' FA6'] || '0'}</td>
                                <td>${s[sub + ' CCA'] || '0'}</td>
                                <td>${s[sub + ' SSA'] || '0'}</td> <td>${s[sub + ' Total'] || '0'}</td>
                                <td><span class="grade-badge ${getGradeClass(s[sub + ' Grade'])}">${s[sub + ' Grade'] || 'N/A'}</span></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </section>

            <footer class="results-footer">
                <div class="result-item"><span>Total Marks Obtained</span> ${s['Total Marks Obtained'] || '0'} / ${s['Max Marks'] || '0'}</div>
                <div class="result-item"><span>Final Percentage</span> ${s['Overall Percentage'] || '0'}%</div>
                <div class="result-item"><span>Overall Result</span> ${s['Result'] || 'N/A'}</div>
                <div class="result-item"><span>Class Rank</span> ${s['Class Rank'] || 'N/A'}</div>
            </footer>
        </div>
    `;
    // --- End of HTML Template ---

    // Configuration for the html2pdf.js library
    const opt = {
        margin:       0, // No margin for full background coverage
        filename:     `${s['Student Name']}_ReportCard.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, letterRendering: true }, // Scale 2 for high definition
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    // New way to render the PDF directly from the generated HTML
    html2pdf().set(opt).from(element).save();
}

    }
}
        
