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
            <td>${s['Student Name']}</td>
            <td>${s["Father's Name"]}</td>
            <td>${s['Roll No']}</td>
            <td>${s['Class']}</td>
            <td>${s['Overall Percentage']}%</td>
            <td><span class="grade-badge">${s['Overall Grade']}</span></td>
            <td>${s['Class Rank']}</td>
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

async function generateDetailedReport(index) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    const s = allStudents[index];

    // Header Section
    doc.setFontSize(22);
    doc.text("GOVERNMENT PRIMARY SCHOOL PETHGAM WAGOORA", 105, 15, { align: "center" });
    
    // Personal Details Grid
    doc.setFontSize(12);
    doc.text(`Student: ${s['Student Name']}`, 20, 30);
    doc.text(`Fathers Name: ${s["Father's Name"]}`, 20, 40);
    doc.text(`Adm No: ${s['Adm No']} | Roll No: ${s['Roll No']}`, 20, 50);

    // Final Results Section
    doc.rect(20, 70, 170, 40); // Box for summary
    doc.text(`Total Marks: ${s['Total Marks Obtained']} / ${s['Max Marks']}`, 25, 80);
    doc.text(`Overall Percentage: ${s['Overall Percentage']}%`, 25, 90);
    doc.text(`Position: ${s['Position in Class']}`, 25, 100);

    doc.save(`${s['Student Name']}_Detailed_Report.pdf`);
}

}
    
