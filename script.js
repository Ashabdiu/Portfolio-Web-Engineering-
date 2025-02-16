function isValidEmail(email) {
    let emailPattern = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailPattern.test(email);
}

function saveToExcel() {
    let name = document.getElementById("name").value;
    let email = document.getElementById("email").value;
    let question = document.getElementById("question").value;
    
    if (!name || !email || !question) {
        alert("Please fill in all fields.");
        return;
    }

    if (!isValidEmail(email)) {
        alert("Please enter a valid email address.");
        return;
    }

    let storedData = localStorage.getItem("excelData");
    let data = storedData ? JSON.parse(storedData) : [["Name", "Email", "Question"]];

    data.push([name, email, question]);
    localStorage.setItem("excelData", JSON.stringify(data));

    let ws = XLSX.utils.aoa_to_sheet(data);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Questions");

    XLSX.writeFile(wb, "questions.xlsx");
    
    alert("Your question has been saved!");
    document.getElementById("contactForm").reset();
}
