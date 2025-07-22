require("dotenv").config();
const XLSX = require("xlsx");
const fs = require("fs");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");
const nodemailer = require("nodemailer");

// Load Excel file
const workbook = XLSX.readFile("testing.xlsx");
const sheetName = workbook.SheetNames[0];
const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
console.log("Total Teams:", data.length);

// Load certificate template
const loadPdfTemplate = async (filename) => {
  return fs.readFileSync(filename);
};

// Generate certificate PDF buffer for a single member
const generateCertificate = async (memberName, teamName, templatePath) => {
  const pdfTemplate = await loadPdfTemplate(templatePath);
  const pdfDoc = await PDFDocument.load(pdfTemplate);
  const page = pdfDoc.getPages()[0];
  const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const { width } = page.getSize();
  const fontSize = 15;

  // Combine name and team name into a single line
  const combinedText = `${memberName} - ${teamName}`;

  // Calculate the text width to center it
  const textWidth = font.widthOfTextAtSize(combinedText, fontSize);
  const centeredX = (width - textWidth) / 2;

  // Draw the combined line centered
  page.drawText(combinedText, {
    x: centeredX,
    y: 290, // Adjust Y as needed
    size: fontSize,
    font,
    color: rgb(0, 0, 0),
  });

  return pdfDoc.save(); // Return buffer
};

// Configure Nodemailer
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Send all team certificates in a single email
const sendCertificatesByEmail = async (teamName, email, attachments) => {
  if (!email || !attachments.length) return;

  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: `Certificates for Team: ${teamName}`,
    text: `Dear Team ${teamName},\n\nPlease find attached the participation certificates for your team members.\n\nBest regards,\nOrganizing Team`,
    attachments,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`📩 Email sent to: ${email}`);
  } catch (error) {
    console.error(`❌ Failed to send email to ${email}:`, error);
  }
};

// Process all teams
const processTeams = async () => {
  for (const team of data) {
    const teamName = team["Team Name"]?.trim();
    const email = team["Email"]?.trim();

    if (!teamName || !email) {
      console.warn("⚠️ Skipping team due to missing Team Name or Email");
      continue;
    }

    const memberFields = [
      "Team Leader Name",
      "MEMEBER 1 NAME",
      "MEMEBER 2 NAME",
      "MEMEBER 3 NAME",
      "MEMEBER 4 NAME",
      "MEMEBER 5 NAME",
    ];

    const attachments = [];

    for (const field of memberFields) {
      const memberName = team[field];

      if (!memberName || memberName.trim() === "") {
        console.log(`⏭️ Skipping empty member in team '${teamName}'`);
        continue;
      }

      const certBuffer = await generateCertificate(memberName.trim(), teamName, "certificate.pdf");

      attachments.push({
        filename: `${memberName.trim().replace(/[^a-z0-9]/gi, "_")}_Certificate.pdf`,
        content: certBuffer,
        contentType: "application/pdf",
      });
    }

    if (attachments.length) {
      await sendCertificatesByEmail(teamName, email, attachments);
    } else {
      console.log(`📭 No valid members to send for team '${teamName}'`);
    }
  }
};

processTeams();
