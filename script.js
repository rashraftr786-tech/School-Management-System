// Global state to keep track of uploaded students
let studentData = [];
let selectedStudentIndex = null;

document.addEventListener('DOMContentLoaded', () => {
    console.log("System Initialized");
});

// --- EXCEL PROCESSING ---
function processExcel() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    if (!file) return alert("Please select an Excel file.");

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            studentData = XLSX.utils.sheet_to_json(sheet);
            
            renderTable(studentData);
            populateClassFilter(studentData);
            alert(`Success! Loaded ${studentData.length} students.`);
        } catch (err) {
            console.error(err);
            alert("Error reading Excel. Check file format.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderTable(data) {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = data.map((student, index) => `
        <tr>
            <td><img src="https://via.placeholder.com/40" class="student-photo" id="img-${index}"></td>
            <td>${student['Student Name'] || 'N/A'}</td>
            <td>${student["Father's Name"] || 'N/A'}</td>
            <td>${student['Roll No'] || 'N/A'}</td>
            <td>${student['Class'] || 'N/A'}</td>
            <td>${student['Total (%)'] || '0'}%</td>
            <td><span class="grade-badge grade-${(student['Grade'] || 'NA').replace(/\+/g, 'plus')}">${student['Grade'] || 'N/A'}</span></td>
            <td><span class="rank-badge">${student['Rank'] || 'N/A'}</span></td>
            <td class="action-buttons">
                <button onclick="openPhotoModal(${index})" class="btn-small btn-photo">Photo</button>
                <button onclick="generateReport(${index})" class="btn-small btn-marksheet">PDF</button>
            </td>
        </tr>
    `).join('');
}

// --- FILTER LOGIC ---
function filterTable() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const classTerm = document.getElementById('classFilter').value;
    
    const filtered = studentData.filter(s => {
        const matchesName = (s['Student Name'] || '').toLowerCase().includes(searchTerm);
        const matchesClass = classTerm === "" || String(s['Class']) === classTerm;
        return matchesName && matchesClass;
    });
    
    renderTable(filtered);
}

function populateClassFilter(data) {
    const select = document.getElementById('classFilter');
    const classes = [...new Set(data.map(s => s['Class']))].sort();
    select.innerHTML = '<option value="">All Classes</option>' + 
        classes.map(c => `<option value="${c}">${c}</option>`).join('');
}

// --- MODAL & PHOTO LOGIC ---
function openPhotoModal(index) {
    selectedStudentIndex = index;
    document.getElementById('photoModal').style.display = 'block';
}

function closeModal() {
    document.getElementById('photoModal').style.display = 'none';
}

function performUpload() {
    const file = document.getElementById('photoFile').files[0];
    if (!file) return alert("Select a file.");

    const reader = new FileReader();
    reader.onload = (e) => {
        document.getElementById(`img-${selectedStudentIndex}`).src = e.target.result;
        closeModal();
    };
    reader.readAsDataURL(file);
}

// --- PDF EXPORT ---
async function generateReport(index) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    if (index !== undefined) {
        // Single Student Logic
        const s = studentData[index];
        doc.setFontSize(20);
        doc.text("OFFICIAL MARKSHEET", 105, 20, { align: "center" });
        doc.setFontSize(12);
        doc.text(`Name: ${s['Student Name']}`, 20, 40);
        doc.text(`Roll No: ${s['Roll No']}`, 20, 50);
        doc.text(`Grade: ${s['Grade']}`, 20, 60);
        doc.save(`${s['Student Name']}_Result.pdf`);
    } else {
        // Full Table Logic
        const element = document.getElementById('resultsTable');
        const canvas = await html2canvas(element);
        const imgData = canvas.toDataURL('image/png');
        doc.addImage(imgData, 'PNG', 10, 10, 190, 0);
        doc.save("Full_School_Report.pdf");
    }
}
