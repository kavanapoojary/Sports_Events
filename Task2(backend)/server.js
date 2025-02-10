const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const app = express();
app.use(cors());

const filePath = path.join(__dirname, "Copy of Student Sports and Cultural Achievements  (Responses).xlsx");
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

function excelSerialToDate(serial) {
    if (!serial || isNaN(serial)) return "Not available"; 
    const excelStartDate = new Date(1899, 11, 30); 
    return new Date(excelStartDate.getTime() + serial * 86400000).toISOString().split("T")[0]; 
}

const cleanedData = data.reduce((acc, student) => {
    const usn = student["USN"]?.toString().trim() || "Not available";
    if (!acc[usn]) {
        acc[usn] = {
            Name: student["Name of the Student\n"]?.toString().trim() || "Not available",
            Email: student["Email Address"]?.toString().trim() || "Not available",
            USN: usn,
            AdmissionYear: student["Admission Year"]?.toString().trim() || "Not available",
            Department: student["Department"]?.toString().trim() || "Not available",
            Email: student["Email Address"]?.toString().trim() || "Not available",
            Events: [] 
        };
    }

    acc[usn].Events.push({
        EventName: student["Name of the Event as per your certificate"]?.toString().trim() || "Not available",
        EventDate: excelSerialToDate(student["Date of Event"]),
        EventType: student["Type of Event"]?.toString().trim() || "Not available",
        CertificatePhotos: student["Upload the Certificate/ Photos of the Event (You can upload upto 5 photos/pdf"]?.toString().trim() || "Not available",
        CertificateDetails: student["Few Lines as per the Certificate"]?.toString().trim() || "Not available"
    });

    return acc;
}, {});

app.get("/students/:usn", (req, res) => {
    const usn = req.params.usn.toUpperCase();
    const student = cleanedData[usn];

    if (student) {
        res.json(student); 
    } else {
        res.status(404).json({ message: "Student not found" });
    }
});
app.listen(5000, () => console.log("Server running on port 5000"));