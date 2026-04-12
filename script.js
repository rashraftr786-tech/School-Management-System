let allStudents = [];
let currentEditIndex = null;

document.addEventListener('DOMContentLoaded', () => {
    // Event Listeners
    document.getElementById('processBtn').onclick = processExcel;
    document.getElementById('searchInput').onkeyup = filterData;
    document.getElementById('classFilter').onchange = filterData;
    
    // Updated: matches the renamed function below
    document.getElementById('downloadFullPDF').onclick = () => generateDetailedReport();
    
    // Modal Close logic
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
            <td><img src="https://via.placeholder.com/45" class="student-photo" id="img-${index}" alt="Student"></td>
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
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    
    // Check if we are printing a single student report
    if (typeof index === 'number') {
        const s = allStudents[index];
        
        // 1. Header
        doc.setFontSize(18);
        doc.text("GOVERNMENT PRIMARY SCHOOL", 105, 15, { align: "center" });
        doc.text("PETHGAM WAGOORA 193109", 105, 22, { align: "center" });
        
        // 2. Student Details
        doc.setFontSize(10);
        doc.text(`Student Name: ${s['Student Name'] || 'N/A'}`, 20, 35);
        doc.text(`Father's Name: ${s["Father's Name"] || 'N/A'}`, 20, 40);
        doc.text(`Mother's Name: ${s['Mother\'s Name'] || 'N/A'}`, 20, 45);
        doc.text(`D.O.B: ${s['D.O.B'] || 'N/A'}`, 120, 35);
        doc.text(`Adm No: ${s['Adm No'] || 'N/A'}`, 120, 40);
        doc.text(`Roll No: ${s['Roll No'] || 'N/A'}`, 120, 45);
        doc.text(`Academic Session: ${s['Academic Session'] || 'N/A'}`, 20, 50);
        doc.text(`PEN No: ${s['PEN No'] || 'N/A'}`, 120, 50);

        // 3. Subject Table Headers
        doc.rect(20, 60, 170, 10);
        doc.text("Subject", 25, 67);
        doc.text("FA1", 60, 67); doc.text("FA2", 80, 67); doc.text("FA3", 100, 67);
        doc.text("FA4", 120, 67); doc.text("Total", 140, 67); doc.text("Grade", 160, 67);

        // 4. Subject Data
        const subjects = ['English', 'Math', 'EVS', 'Urdu', 'Kashmiri'];
        let y = 75;
        subjects.forEach(sub => {
            doc.text(sub, 25, y);
            doc.text(`${s[sub + ' FA1'] || '0'}`, 60, y);
            doc.text(`${s[sub + ' FA2'] || '0'}`, 80, y);
            doc.text(`${s[sub + ' FA3'] || '0'}`, 100, y);
            doc.text(`${s[sub + ' FA4'] || '0'}`, 120, y);
            doc.text(`${s[sub + ' Total'] || '0'}`, 140, y);
            doc.text(`${s[sub + ' Grade'] || 'N/A'}`, 160, y);
            y += 7;
        });

        // 5. Final Results Box
        doc.rect(20, y + 10, 170, 30);
        doc.text(`Total Marks: ${s['Total Marks Obtained'] || '0'} / ${s['Max Marks'] || '0'}`, 25, y + 20);
        doc.text(`Overall Percentage: ${s['Overall Percentage'] || '0'}%`, 25, y + 27);
        doc.text(`Result: ${s['Result'] || 'N/A'}`, 120, y + 20);
        doc.text(`Class Rank: ${s['Class Rank'] || 'N/A'}`, 120, y + 27);
        
        doc.text(`General Remarks: ${s['General Remarks'] || 'Well done!'}`, 20, y + 45);

        doc.save(`${s['Student Name']}_ReportCard.pdf`);
    } else {
        // Fallback for full table export
        const table = document.getElementById('resultsTable');
        if (table) {
            const canvas = await html2canvas(table, { scale: 2 });
            const imgData = canvas.toDataURL('image/png');
            doc.addImage(imgData, 'PNG', 10, 10, 190, 0);
            doc.save("Full_School_Report.pdf");
        }
    }
}
        
