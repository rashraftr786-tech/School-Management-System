let studentData = [];

// Event Listener for file upload
document.getElementById('excelFile').addEventListener('change', async (e) => {
    studentData = await ExcelHelper.parse(e.target.files[0]);
    renderTable(studentData);
});

function renderTable(data) {
    // DOM manipulation logic
    console.log("Rendering", data.length, "students.");
}
