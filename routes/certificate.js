const router = require("express").Router();
const fs = require("fs");
const path = require("path");
const multer = require("multer");
const ExcelJS = require("exceljs");
const archiver = require("archiver");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");

const upload = multer({ dest: "uploads/" });

// Helper function to format date from YYYY-MM-DD to DD/MM/YYYY
function formatDate(dateString) {
    if (!dateString) return "";
    
    // If already in DD/MM/YYYY format, return as is
    if (dateString.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        return dateString;
    }
    
    // Convert from YYYY-MM-DD to DD/MM/YYYY
    const parts = dateString.split('-');
    if (parts.length === 3) {
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    
    // If it's already in some other format, try to parse it
    try {
        const date = new Date(dateString);
        if (!isNaN(date)) {
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        }
    } catch (e) {
        // If all fails, return original
    }
    
    return dateString;
}

/* =========================
INDIVIDUAL COORDINATES FOR EACH CERTIFICATE
========================= */

const CERTIFICATE_COORDINATES = {
    // VOLUNTEER
    volunteer: {
        name: { x: 0, y: 462, size: 40 },
        event: { x: 660, y: 409, size: 18 },
        date: { x: 298, y: 210, size: 20 }
    },

    // ALL 36 VDP CERTIFICATES
    vdp: {
        "Wafers Making": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Vermicompost": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Value-Added Products from Cashews": {
            name: { x: 0, y: 1500, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Spoken English Course": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Spice Making": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Soap Making and Agarbattii": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Poultry Training": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Personality Development": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Papad and Pickel Making": {
            name: { x: 0, y: 1498, size: 85 },
            date: { x: 1825, y: 900, size: 60 }
        },
        "Organic Fertilizer": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Nauvari Saree Stitching": {
            name: { x: 0, y: 1575, size: 85 },
            date: { x: 1825, y: 920, size: 60 }
        },
        "Nath Making Course": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Micron Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Mehandi Training": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Masala Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Karate Training": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Jute Bag Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Jewelry Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Jackfruit Processing": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Incense Stick Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Homemade Hygiene Products": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Gravy Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Fruit Processing": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Fast Food Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Fast food and Snacks Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Fashion Designing": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Fabric Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Dough Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Cloth Bag Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Cloth and Jute Bag Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Candle Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Cake Making": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Cake Decoration": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Beauty Parlour": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Candle And Incense Stick": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        },
        "Advance Beauty Parlour": {
            name: { x: 0, y: 1055, size: 60 },
            date: { x: 1213, y: 610, size: 40 }
        }
    },

    // APPRECIATION CERTIFICATES
    appreciation: {
        "Employee Appreciation": {
            name: { x: 0, y: 1010, size: 80 },
            date: { x: 1010, y: 480, size: 60 }
        },
        "Wish Tree Appreciation": {
            name: { x: 450, y: 895, size: 50 },
            date: { x: 2155, y: 385, size: 40 }
        },
        "Donor Appreciation": {
            name: { x: 0, y: 495, size: 41 },
            date: { x: 325, y: 182, size: 20 }
        }
    },

    // OTHER CERTIFICATES
    other: {
        "Employee of the Month": {
            name: { x: 0, y: 450, size: 40 },
            date: { x: 310, y: 220, size: 18 }
        },
        "Employee of the Year": {
            name: { x: 0, y: 455, size: 38 },
            date: { x: 315, y: 215, size: 18 }
        },
        "Team of the Year": {
            name: { x: 0, y: 445, size: 39 },
            date: { x: 308, y: 218, size: 18 }
        },
        "TTT Certificate": {
            name: { x: 0, y: 448, size: 40 },
            date: { x: 312, y: 222, size: 18 }
        },
        "Aari Work": {
            name: { x: 0, y: 452, size: 38 },
            date: { x: 310, y: 220, size: 18 }
        },
        "Tailoring Certificate": {
            name: { x: 0, y: 450, size: 39 },
            date: { x: 308, y: 218, size: 18 }
        },
        "Partner Trainer": {
            name: { x: 0, y: 446, size: 40 },
            date: { x: 314, y: 224, size: 18 }
        }
    }
};

// Helper function to get coordinates for a specific certificate
function getCertificateCoordinates(type, eventName) {
    if (type === "volunteer") {
        return CERTIFICATE_COORDINATES.volunteer;
    }

    const typeCoords = CERTIFICATE_COORDINATES[type];
    if (!typeCoords) {
        return {
            name: { x: 0, y: 450, size: 38 },
            date: { x: 310, y: 200, size: 18 }
        };
    }

    let coords = typeCoords[eventName];
    if (!coords) {
        const matchKey = Object.keys(typeCoords).find(key => 
            key.toLowerCase() === eventName.toLowerCase()
        );
        coords = matchKey ? typeCoords[matchKey] : null;
    }

    if (!coords) {
        const firstKey = Object.keys(typeCoords)[0];
        coords = firstKey ? typeCoords[firstKey] : {
            name: { x: 0, y: 450, size: 38 },
            date: { x: 310, y: 200, size: 18 }
        };
    }

    return coords;
}

// FIND ANY AVAILABLE TEMPLATE
function findAnyTemplate(projectRoot) {
    const searchPaths = [
        path.join(projectRoot, "assets", "volunteer-template.png"),
        path.join(projectRoot, "assets", "vdp"),
        path.join(projectRoot, "assets", "appreciation"),
        path.join(projectRoot, "assets", "other"),
        path.join(projectRoot, "assets"),
        path.join(projectRoot, "public", "assets")
    ];
    
    for (const searchPath of searchPaths) {
        if (fs.existsSync(searchPath)) {
            try {
                if (fs.statSync(searchPath).isDirectory()) {
                    const files = fs.readdirSync(searchPath);
                    const pngFiles = files.filter(f => f.endsWith('.png') || f.endsWith('.jpg'));
                    if (pngFiles.length > 0) {
                        const found = path.join(searchPath, pngFiles[0]);
                        console.log("Found template:", found);
                        return found;
                    }
                } else if (searchPath.endsWith('.png') || searchPath.endsWith('.jpg')) {
                    if (fs.existsSync(searchPath)) {
                        console.log("Found template:", searchPath);
                        return searchPath;
                    }
                }
            } catch (err) {
                // Skip if can't read directory
            }
        }
    }
    
    return null;
}

// Get template with correct assets path
function getTemplate(type, templatePath) {
    const projectRoot = path.resolve(__dirname, "..");
    
    if (type === "volunteer") {
        const defaultTemplate = path.join(projectRoot, "assets", "volunteer-template.png");
        console.log("Looking for volunteer template at:", defaultTemplate);
        
        if (fs.existsSync(defaultTemplate)) {
            return defaultTemplate;
        }
        
        console.log("Volunteer template not found, searching for any template...");
        const anyTemplate = findAnyTemplate(projectRoot);
        if (anyTemplate) {
            console.log("Using fallback template:", anyTemplate);
            return anyTemplate;
        }
        
        throw new Error(`No template found. Please place a template PNG in: ${path.join(projectRoot, "assets")}`);
    }

    if (templatePath) {
        let clean = String(templatePath);
        clean = clean.replace(/^https?:\/\/[^\/]+/, '');
        clean = clean.replace(/^\//, '');
        
        if (clean.startsWith("assets/")) {
            const rootAssetsPath = path.join(projectRoot, clean);
            const publicAssetsPath = path.join(projectRoot, "public", clean);
            
            console.log("Looking for template at:", rootAssetsPath);
            
            if (fs.existsSync(rootAssetsPath)) {
                return rootAssetsPath;
            }
            
            if (fs.existsSync(publicAssetsPath)) {
                console.log("Found template in public/assets:", publicAssetsPath);
                return publicAssetsPath;
            }
        } else {
            const fullPath = path.join(projectRoot, clean);
            if (fs.existsSync(fullPath)) {
                return fullPath;
            }
        }
        
        const fileName = path.basename(clean);
        const assetsDir = path.join(projectRoot, "assets");
        const subDirs = ['vdp', 'appreciation', 'other', 'others'];
        
        if (fs.existsSync(assetsDir)) {
            for (const dir of subDirs) {
                const testPath = path.join(assetsDir, dir, fileName);
                if (fs.existsSync(testPath)) {
                    console.log("Found template in subdirectory:", testPath);
                    return testPath;
                }
            }
            
            for (const dir of subDirs) {
                const dirPath = path.join(assetsDir, dir);
                if (fs.existsSync(dirPath)) {
                    try {
                        const files = fs.readdirSync(dirPath);
                        const searchName = fileName.replace('.png', '').replace('.jpg', '').toLowerCase();
                        const match = files.find(f => 
                            f.toLowerCase().includes(searchName)
                        );
                        if (match) {
                            const foundPath = path.join(dirPath, match);
                            console.log("Found matching template:", foundPath);
                            return foundPath;
                        }
                    } catch (err) {
                        // Skip if can't read directory
                    }
                }
            }
        }
        
        console.log("Template not found, searching for any template...");
        const anyTemplate = findAnyTemplate(projectRoot);
        if (anyTemplate) {
            console.log("Using fallback template:", anyTemplate);
            return anyTemplate;
        }
        
        throw new Error(`Template not found: ${clean}`);
    }

    const anyTemplate = findAnyTemplate(projectRoot);
    if (anyTemplate) {
        console.log("Using fallback template:", anyTemplate);
        return anyTemplate;
    }
    
    throw new Error("No template found. Please place a template PNG in assets/ folder");
}

async function generatePDF(type, templatePath, name, event, date, output) {
    try {
        const templateFilePath = getTemplate(type, templatePath);
        
        console.log("Loading template from:", templateFilePath);

        const imageBytes = fs.readFileSync(templateFilePath);
        const pdf = await PDFDocument.create();
        
        let image;
        if (templateFilePath.toLowerCase().endsWith(".jpg") || templateFilePath.toLowerCase().endsWith(".jpeg")) {
            image = await pdf.embedJpg(imageBytes);
        } else {
            image = await pdf.embedPng(imageBytes);
        }

        const page = pdf.addPage([image.width, image.height]);
        page.drawImage(image, {
            x: 0,
            y: 0,
            width: image.width,
            height: image.height
        });

        const font = await pdf.embedFont(StandardFonts.TimesRomanBold);
        const pos = getCertificateCoordinates(type, event);

        // Format the date before drawing
        const formattedDate = formatDate(date);

        // DRAW NAME
        let nameSize = pos.name.size;
        if (name.length > 20) nameSize -= 4;
        if (name.length > 30) nameSize -= 4;
        
        const nameWidth = font.widthOfTextAtSize(name.toUpperCase(), nameSize);
        page.drawText(name.toUpperCase(), {
            x: ((image.width - nameWidth) / 2) + (pos.name.x || 0),
            y: pos.name.y,
            size: nameSize,
            font,
            color: rgb(0, 0, 0)
        });

        // DRAW EVENT (only for volunteer type)
        if (type === "volunteer" && event) {
            const eventFont = await pdf.embedFont(StandardFonts.HelveticaBold);
            const volunteerCoords = CERTIFICATE_COORDINATES.volunteer;
            if (volunteerCoords.event) {
                const eventWidth = eventFont.widthOfTextAtSize(event, volunteerCoords.event.size);
                page.drawText(event, {
                    x: volunteerCoords.event.x - (eventWidth / 2),
                    y: volunteerCoords.event.y,
                    size: volunteerCoords.event.size,
                    font: eventFont
                });
            }
        }

        // DRAW DATE (formatted)
        const dateWidth = font.widthOfTextAtSize(formattedDate, pos.date.size);
        page.drawText(formattedDate, {
            x: pos.date.x - (dateWidth / 2),
            y: pos.date.y,
            size: pos.date.size,
            font
        });

        const finalPdf = await pdf.save();
        fs.writeFileSync(output, finalPdf);
        console.log("PDF generated successfully:", output);
        
    } catch (error) {
        console.error("Error in generatePDF:", error);
        throw error;
    }
}

// SINGLE CERTIFICATE
router.post("/single", async (req, res) => {
    try {
        const { type, template, name, event, date } = req.body;

        console.log("Received request:", { type, template, name, event, date });

        if (!type || !name || !date) {
            return res.status(400).send("Missing required fields: type, name, or date");
        }

        if (!fs.existsSync("generated")) {
            fs.mkdirSync("generated");
        }

        const file = path.join("generated", `${Date.now()}.pdf`);
        
        await generatePDF(type, template, name, event, date, file);
        
        res.sendFile(path.resolve(file));
    } catch (e) {
        console.error("Single generation error:", e);
        res.status(500).send(`Generation Error: ${e.message}`);
    }
});

// BULK CERTIFICATES
router.post("/bulk", upload.single("excel"), async (req, res) => {
    try {
        const { type, template } = req.body;

        if (!req.file) {
            return res.status(400).send("No Excel file uploaded");
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);

        const sheet = workbook.getWorksheet(1);
        if (!sheet) {
            return res.status(400).send("No worksheet found in Excel file");
        }

        if (!fs.existsSync("generated")) {
            fs.mkdirSync("generated");
        }

        const zip = path.join("generated", "certificates.zip");
        const output = fs.createWriteStream(zip);
        const archive = archiver("zip", { zlib: { level: 9 } });
        archive.pipe(output);

        for (let i = 2; i <= sheet.rowCount; i++) {
            const row = sheet.getRow(i);
            const name = String(row.getCell(1).value || "").trim();
            
            if (!name) continue;

            let event = "";
            let date = "";

            if (type === "volunteer") {
                event = String(row.getCell(2).value || "").trim();
                date = String(row.getCell(3).value || "").trim();
            } else {
                // For VDP, Appreciation, Other - Column 2 is the date
                date = String(row.getCell(2).value || "").trim();
            }

            if (!date) continue;

            // Format the date for PDF
            const formattedDate = formatDate(date);

            const file = path.join("generated", `${name.replace(/[^a-zA-Z0-9]/g, '_')}.pdf`);
            
            await generatePDF(type, template, name, event, formattedDate, file);
            archive.file(file, { name: `${name}.pdf` });
        }

        await archive.finalize();

        output.on("close", () => {
            res.download(zip);
        });

    } catch (e) {
        console.error("Bulk generation error:", e);
        res.status(500).send(`Bulk Error: ${e.message}`);
    }
});

module.exports = router;