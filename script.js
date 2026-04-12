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
    const file = document.getElementById('excelFile').files[0];
    if (!file) return alert("Please select a file!");

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        allStudents = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        populateClassFilter();
        renderTable(allStudents);
    };
    reader.readAsArrayBuffer(file);
}

function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    // Safety check for empty data
    if (!tbody) return;
    
    tbody.innerHTML = data.map((s, index) => `
        <tr>
            <td><img src="https://via.placeholder.com/45" class="student-photo" id="img-${index}"></td>
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
    const classes = [...new Set(allStudents.map(s => s['Class']))];
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

    // Case 1: Individual Report
    if (typeof index === 'number') {
        const s = allStudents[index];
        doc.setFontSize(22);
        doc.text("GOVERNMENT PRIMARY SCHOOL PETHGAM WAGOORA", 105, 15, { align: "center" });
        doc.setFontSize(12);
        doc.text(`Student: ${s['Student Name'] || 'N/A'}`, 20, 30);
        doc.text(`Fathers Name: ${s["Father's Name"] || 'N/A'}`, 20, 40);
        doc.text(`Adm No: ${s['Adm No'] || 'N/A'} | Roll No: ${s['Roll No'] || 'N/A'}`, 20, 50);
        doc.rect(20, 70, 170, 40);
        doc.text(`Total Marks: ${s['Total Marks Obtained'] || '0'} / ${s['Max Marks'] || '0'}`, 25, 80);
        doc.save(`${s['Student Name']}_Report.pdf`);
    } 
    // Case 2: Full Table Export
    else {
        const table = document.getElementById('resultsTable');
        const canvas = await html2canvas(table, { scale: 2 });
        const imgData = canvas.toDataURL('image/png');
        doc.addImage(imgData, 'PNG', 10, 10, 190, 0);
        doc.save("Full_School_Report.pdf");
    }
}
