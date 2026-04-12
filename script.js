let allStudents = [];
let currentEditIndex = null;

document.addEventListener('DOMContentLoaded', () => {
    // Event Listeners
    document.getElementById('processBtn').onclick = processExcel;
    document.getElementById('searchInput').onkeyup = filterData;
    document.getElementById('classFilter').onchange = filterData;
    document.getElementById('downloadFullPDF').onclick = () => generatePDF();
    
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
    tbody.innerHTML = data.map((s, index) => `
        <tr>
            <td><img src="https://via.placeholder.com/45" class="student-photo" id="img-${index}"></td>
            <td>${s['Student Name'] || 'N/A'}</td>
            <td>${s["Father's Name"] || 'N/A'}</td>
            <td>${s['Roll No'] || 'N/A'}</td>
            <td>${s['Class'] || 'N/A'}</td>
            <td>${s['Total (%)'] || '0'}%</td>
            <td><span class="grade-badge grade-${(s['Grade'] || '').replace('+', 'plus')}">${s['Grade'] || 'N/A'}</span></td>
            <td>${s['Rank'] || 'N/A'}</td>
            <td>
                <button onclick="openModal(${index})" class="btn-sm btn-primary">Photo</button>
                <button onclick="generatePDF(${index})" class="btn-sm btn-secondary">PDF</button>
            </td>
        </tr>
    `).join('');
}

function filterData() {
    const search = document.getElementById('searchInput').value.toLowerCase();
    const classVal = document.getElementById('classFilter').value;

    const filtered = allStudents.filter(s => {
        const matchesName = String(s['Student Name']).toLowerCase().includes(search);
        const matchesClass = classVal === "" || String(s['Class']) === classVal;
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
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            document.getElementById(`img-${currentEditIndex}`).src = e.target.result;
            document.getElementById('photoModal').style.display = 'none';
        };
        reader.readAsDataURL(file);
    }
}

async function generatePDF(index) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    if (index !== undefined) {
        // Single Marksheet
        const s = allStudents[index];
        doc.setFontSize(20);
        doc.text("OFFICIAL REPORT CARD", 105, 20, { align: "center" });
        doc.setFontSize(12);
        doc.text(`Student Name: ${s['Student Name']}`, 20, 40);
        doc.text(`Roll No: ${s['Roll No']}`, 20, 50);
        doc.text(`Grade: ${s['Grade']}`, 20, 60);
        doc.save(`${s['Student Name']}_Result.pdf`);
    } else {
        // Full Table capture
        const table = document.getElementById('resultsTable');
        const canvas = await html2canvas(table);
        const imgData = canvas.toDataURL('image/png');
        doc.addImage(imgData, 'PNG', 10, 10, 190, 0);
        doc.save("Full_School_Report.pdf");
    }
}
    
