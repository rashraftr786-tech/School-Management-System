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
    if (!window.jspdf) return alert("PDF library not loaded.");
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    
    if (typeof index === 'number') {
        const s = allStudents[index];
        
        doc.setFontSize(18);
        doc.text("GOVERNMENT PRIMARY SCHOOL", 105, 15, { align: "center" });
        doc.setFontSize(12);
        doc.text("PETHGAM WAGOORA 193109", 105, 22, { align: "center" });
        
        doc.setFontSize(10);
        doc.text(`Student Name: ${s['Student Name'] || 'N/A'}`, 20, 35);
        doc.text(`Roll No: ${s['Roll No'] || 'N/A'}`, 140, 35);
        doc.text(`Father's Name: ${s["Father's Name"] || 'N/A'}`, 20, 40);
        doc.text(`Academic Session: ${s['Academic Session'] || 'N/A'}`, 140, 40);

        // Adjusted layout to fit A4 width (max 210mm)
        doc.rect(10, 55, 190, 10);
        doc.text("Subject", 12, 62);
        doc.text("FA1", 50, 62); doc.text("FA2", 65, 62); doc.text("FA3", 80, 62);
        doc.text("FA4", 95, 62); doc.text("FA5", 110, 62); doc.text("FA6", 125, 62); 
        doc.text("CCA", 140, 62); doc.text("SA", 155, 62); doc.text("Total", 170, 62); 
        doc.text("Grade", 185, 62);

        const subjects = ['English', 'Math', 'EVS', 'Urdu', 'Kashmiri'];
        let y = 75;
        subjects.forEach(sub => {
            doc.text(sub, 12, y);
            doc.text(`${s[sub + ' FA1'] || '0'}`, 50, y);
            doc.text(`${s[sub + ' FA2'] || '0'}`, 65, y);
            doc.text(`${s[sub + ' FA3'] || '0'}`, 80, y);
            doc.text(`${s[sub + ' FA4'] || '0'}`, 95, y);
            doc.text(`${s[sub + ' FA5'] || '0'}`, 110, y);
            doc.text(`${s[sub + ' FA6'] || '0'}`, 125, y);
            doc.text(`${s[sub + ' CCA'] || '0'}`, 140, y);
            doc.text(`${s[sub + ' SSA'] || '0'}`, 155, y);
            doc.text(`${s[sub + ' Total'] || '0'}`, 170, y);
            doc.text(`${s[sub + ' Grade'] || 'N/A'}`, 185, y);
            y += 10;
        });

        doc.rect(10, y + 5, 190, 20);
        doc.text(`Total Marks: ${s['Total Marks Obtained'] || '0'} / ${s['Max Marks'] || '0'}`, 15, y + 12);
        doc.text(`Percentage: ${s['Overall Percentage'] || '0'}%`, 15, y + 19);
        doc.text(`Result: ${s['Result'] || 'N/A'}`, 120, y + 12);
        
        doc.save(`${s['Student Name']}_ReportCard.pdf`);
    } else {
        const table = document.getElementById('resultsTable');
        if (table && window.html2canvas) {
            const canvas = await html2canvas(table, { scale: 2 });
            const imgData = canvas.toDataURL('image/png');
            doc.addImage(imgData, 'PNG', 10, 10, 190, 0);
            doc.save("Full_School_Report.pdf");
        }
    }
}
        
