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

async function generateDetailedReport(index) {
    if (!window.html2pdf) return alert("html2pdf library not loaded! See instructions.");
    
    const s = allStudents[index];
    if (!s) return alert("Error: Student data not found.");

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

    const element = document.createElement('div');
    element.innerHTML = `
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { background-color: #f7f9fc; }
            #report-card-container { width: 210mm; min-height: 297mm; font-family: sans-serif; background-color: white; padding: 15mm; }
            .main-header { display: flex; align-items: center; gap: 20px; background: linear-gradient(135deg, #6a11cb 0%, #2575fc 100%); color: white; padding: 25px 35px; border-radius: 15px; margin-bottom: 25px; }
            .school-logo { width: 80px; height: 80px; border-radius: 15px; background: white; padding: 10px; object-fit: cover; }
            .profile-section { display: flex; background: white; padding: 20px; border: 1px solid #e1e8f0; margin-bottom: 25px; gap: 25px; align-items: center; }
            .profile-photo { width: 90px; height: 90px; border-radius: 20px; border: 4px solid #2575fc; }
            .marks-table { width: 100%; border-collapse: collapse; }
            .marks-table th, .marks-table td { padding: 12px; border: 1px solid #ddd; text-align: center; }
            .grade-badge { padding: 5px 10px; border-radius: 10px; color: white; font-size: 12px; }
            .grade-a { background: #28a745; }
            .grade-b { background: #ffc107; }
            .grade-f { background: #dc3545; }
            .grade-na { background: #999; }
            .grade-default { background: #2575fc; }
            .results-footer { display: flex; background: #f1f7ff; padding: 20px; margin-top: 30px; gap: 20px; }
        </style>

        <div id="report-card-container">
            <header class="main-header">
                <img src="${studentPhotoSrc}" class="school-logo" alt="Logo">
                <div><h1>GOVERNMENT PRIMARY SCHOOL</h1><p>OFFICIAL STUDENT REPORT CARD</p></div>
            </header>
            <section class="profile-section">
                <img src="${studentPhotoSrc}" class="profile-photo" alt="Photo">
                <div>
                    <p><b>Student Name:</b> ${s['Student Name'] || 'N/A'}</p>
                    <p><b>Father\'s Name:</b> ${s["Father's Name"] || 'N/A'}</p>
                </div>
            </section>
            <section>
                <table class="marks-table">
                    <thead><tr><th>Subject</th><th>Total</th><th>Grade</th></tr></thead>
                    <tbody>
                        ${['English', 'Math', 'EVS'].map(sub => `
                            <tr>
                                <td>${sub}</td>
                                <td>${s[sub + ' Total'] || '0'}</td>
                                <td><span class="grade-badge ${getGradeClass(s[sub + ' Grade'])}">${s[sub + ' Grade'] || 'N/A'}</span></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </section>
            <footer class="results-footer">
                <div><b>Overall Percentage:</b> ${s['Overall Percentage'] || '0'}%</div>
            </footer>
        </div>
    `;

    const opt = {
        margin: 0,
        filename: `${s['Student Name']}_ReportCard.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    html2pdf().set(opt).from(element).save();
}





            
                
