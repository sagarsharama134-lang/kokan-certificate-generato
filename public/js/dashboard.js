/* ELEMENTS */
const preview = document.getElementById("previewImage");
const pageTitle = document.getElementById("pageTitle");
const eventField = document.getElementById("eventField");

const volunteerBtn = document.getElementById("volunteerBtn");
const vdpSelect = document.getElementById("vdpSelect");
const appSelect = document.getElementById("appSelect");
const otherSelect = document.getElementById("otherSelect");

const singleBtn = document.getElementById("singleBtn");
const bulkBtn = document.getElementById("bulkBtn");
const singleSection = document.getElementById("singleSection");
const bulkSection = document.getElementById("bulkSection");
const previewBtn = document.getElementById("previewBtn");
const editBtn = document.getElementById("editBtn");
const generateSingle = document.getElementById("generateSingle");
const logoutBtn = document.getElementById("logoutBtn");

const previewName = document.getElementById("previewName");
const previewEvent = document.getElementById("previewEvent");
const previewDate = document.getElementById("previewDate");
const generateBulk = document.getElementById("generateBulk");
const excelFile = document.getElementById("excelFile");
const bulkInfo = document.getElementById("bulkInfo");
const eventInput = document.getElementById("event");

// Default event text for volunteer
const DEFAULT_EVENT_TEXT = `We recognise your valuable contribution to the Kokan Summer Camp - 2026, organised by Kokan NGO India.
Your dedication, coordination, and commitment played a vital role in making this initiative impactful and meaningful.
Your efforts have truly contributed to saving lives and strengthening the spirit of humanity.`;

/* =======================
INDIVIDUAL PREVIEW COORDINATES FOR EACH CERTIFICATE
ALL 36 VDP EVENTS INCLUDED
======================= */

const PREVIEW_COORDINATES = {
    // VOLUNTEER
    volunteer: {
        name: { left: "50%", top: "40%", size: "34px" },
        event: { left: "50%", top: "47%", size: "14px" },
        date: { left: "28.7%", top: "72%", size: "18px" }
    },

    // ALL 36 VDP CERTIFICATES
    vdp: {
        "Wafers Making": {
            name: { left: "50%", top: "52.3%", size: "33px" },
            date: { left: "69.9%", top: "72.3%", size: "20px" }
        },
        "Vermicompost": {
            name: { left: "50%", top: "52.5%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Value-Added Products from Cashews": {
            name: { left: "50%", top: "55%", size: "32px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Spoken English Course": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Spice Making": {
            name: { left: "50%", top: "52.5%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Soap Making and Agarbattii": {
            name: { left: "50%", top: "52.3%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Poultry Training": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Personality Development": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Papad and Pickel Making": {
            name: { left: "50%", top: "54.9%", size: "33px" },
            date: { left: "69.9%", top: "72.8%", size: "20px" }
        },
        "Organic Fertilizer": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Nauvari Saree Stitching": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Nath Making Course": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Micron Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Mehandi Training": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Masala Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Karate Training": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Jute Bag Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Jewelry Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Jackfruit Processing": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Incense Stick Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Homemade Hygiene Products": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Gravy Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Fruit Processing": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Fast Food Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Fast food and Snacks Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Fashion Designing": {
            name: { left: "50%", top: "54.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Fabric Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Dough Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Cloth Bag Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Cloth and Jute Bag Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Candle Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Cake Making": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Cake Decoration": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Beauty Parlour": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Candle And Incense Stick": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        },
        "Advance Beauty Parlour": {
            name: { left: "50%", top: "52.4%", size: "33px" },
            date: { left: "69.9%", top: "72.2%", size: "20px" }
        }
    },

    // APPRECIATION CERTIFICATES
    appreciation: {
        "Employee Appreciation": {
            name: { left: "50%", top: "54.4%", size: "30px" },
            date: { left: "32%", top: "78%", size: "18px" }
        },
        "Wish Tree Appreciation": {
            name: { left: "65%", top: "45%", size: "19px" },
            date: { left: "81%", top: "76%", size: "14px" }
        },
        "Donor Appreciation": {
            name: { left: "50%", top: "52.8%", size: "31px" },
            date: { left: "27.5%", top: "75.5%", size: "18px" }
        }
    },

    // OTHER CERTIFICATES
    other: {
        "Employee of the Month": {
            name: { left: "50%", top: "55%", size: "35px" },
            date: { left: "25%", top: "80%", size: "18px" }
        },
        "Employee of the Year": {
            name: { left: "50%", top: "55.5%", size: "33px" },
            date: { left: "25.5%", top: "79.5%", size: "18px" }
        },
        "Team of the Year": {
            name: { left: "50%", top: "54.5%", size: "34px" },
            date: { left: "25%", top: "80%", size: "18px" }
        },
        "TTT Certificate": {
            name: { left: "50%", top: "55%", size: "35px" },
            date: { left: "25.5%", top: "80.5%", size: "18px" }
        },
        "Aari Work": {
            name: { left: "50%", top: "55.2%", size: "33px" },
            date: { left: "25%", top: "80%", size: "18px" }
        },
        "Tailoring Certificate": {
            name: { left: "50%", top: "55%", size: "34px" },
            date: { left: "25%", top: "79.5%", size: "18px" }
        },
        "Partner Trainer": {
            name: { left: "50%", top: "54.6%", size: "35px" },
            date: { left: "25.5%", top: "80.5%", size: "18px" }
        }
    }
};

// Helper function to format date from YYYY-MM-DD to DD/MM/YYYY
function formatDate(dateString) {
    if (!dateString) return "";
    
    if (dateString.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        return dateString;
    }
    
    const parts = dateString.split('-');
    if (parts.length === 3) {
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    
    try {
        const date = new Date(dateString);
        if (!isNaN(date)) {
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        }
    } catch (e) {}
    
    return dateString;
}

// Helper function to get preview coordinates
function getPreviewCoordinates(type, eventName) {
    if (type === "volunteer") {
        return PREVIEW_COORDINATES.volunteer;
    }

    const typeCoords = PREVIEW_COORDINATES[type];
    if (!typeCoords) {
        return {
            name: { left: "50%", top: "55%", size: "34px" },
            date: { left: "25%", top: "80%", size: "18px" }
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
            name: { left: "50%", top: "55%", size: "34px" },
            date: { left: "25%", top: "80%", size: "18px" }
        };
    }

    return coords;
}

function applyPreviewPosition() {
    const type = getType();
    const eventName = getEventName();
    const pos = getPreviewCoordinates(type, eventName);

    previewName.style.left = pos.name.left;
    previewName.style.top = pos.name.top;
    previewName.style.fontSize = pos.name.size;
    previewName.style.transform = "translateX(-50%)";

    if (type === "volunteer") {
        previewEvent.style.display = "block";
        previewEvent.style.left = pos.event.left;
        previewEvent.style.top = pos.event.top;
        previewEvent.style.fontSize = pos.event.size;
        previewEvent.style.transform = "translateX(-50%)";
        previewEvent.style.whiteSpace = "pre-wrap";
        previewEvent.style.width = "70%";
        previewEvent.style.textAlign = "center";
        previewEvent.style.lineHeight = "1.4";
        previewEvent.style.color = "#6e4737";
        previewEvent.style.fontFamily = "StandardFonts.HelveticaOblique";
        previewEvent.style.fontWeight = "900";
        previewEvent.style.fontStyle = "italic";
    } else {
        previewEvent.style.display = "none";
    }

    previewDate.style.left = pos.date.left;
    previewDate.style.top = pos.date.top;
    previewDate.style.fontSize = pos.date.size;
    previewDate.style.transform = "translateX(-50%)";
}

function resetPreview() {
    previewName.innerText = "";
    previewEvent.innerText = "";
    previewDate.innerText = "";
    previewBtn.style.display = "inline-block";
    editBtn.style.display = "none";
    generateSingle.style.display = "none";

    document.querySelectorAll(".form input").forEach(x => x.disabled = false);
    document.getElementById("name").value = "";
    document.getElementById("date").value = "";
}

function clearSelections() {
    vdpSelect.selectedIndex = 0;
    appSelect.selectedIndex = 0;
    otherSelect.selectedIndex = 0;
}

function getSelectedSrc() {
    if (getType() === "vdp") {
        return vdpSelect.selectedOptions[0]?.dataset.src;
    }
    if (getType() === "appreciation") {
        return appSelect.selectedOptions[0]?.dataset.src;
    }
    if (getType() === "other") {
        return otherSelect.selectedOptions[0]?.dataset.src;
    }
    return "/assets/volunteer-template.png";
}

function getEventName() {
    if (getType() === "volunteer") {
        return "volunteer";
    }
    if (getType() === "vdp") {
        return vdpSelect.value;
    }
    if (getType() === "appreciation") {
        return appSelect.value;
    }
    if (getType() === "other") {
        return otherSelect.value;
    }
    return "";
}

function setPreview(type, eventName = "") {
    resetPreview();
    applyPreviewPosition();

    if (type === "volunteer") {
        pageTitle.innerText = "Volunteer";
        eventField.classList.remove("hidden");

        const eventInput = document.getElementById("event");
        if (!eventInput.value) {
            eventInput.value = DEFAULT_EVENT_TEXT;
        }

        preview.src = "/assets/volunteer-template.png";
        fillPreview();
        updateBulkInfo(type);
        return;
    }

    eventField.classList.add("hidden");
    pageTitle.innerText = eventName || type.toUpperCase();
    
    const src = getSelectedSrc();
    if (src && src !== "undefined") {
        preview.src = src;
    } else {
        preview.src = "/assets/volunteer-template.png";
    }
    updateBulkInfo(type);
}

function getType() {
    if (!eventField.classList.contains("hidden")) {
        return "volunteer";
    }
    if (vdpSelect.value) return "vdp";
    if (appSelect.value) return "appreciation";
    if (otherSelect.value) return "other";
    return "volunteer";
}

function getEvent() {
    if (getType() === "volunteer") {
        return document.getElementById("event").value;
    }
    return getEventName();
}

function fillPreview() {
    applyPreviewPosition();
    const name = document.getElementById("name").value || "Name";
    previewName.innerText = name;
    
    const rawDate = document.getElementById("date").value || "";
    const formattedDate = formatDate(rawDate);
    previewDate.innerText = formattedDate || "Date";

    if (getType() === "volunteer") {
        const event = document.getElementById("event").value || "Event";
        previewEvent.innerText = event;
    }
}

// Update bulk info based on certificate type
function updateBulkInfo(type) {
    if (type === "volunteer") {
        bulkInfo.innerHTML = `
            <strong>Excel Format:</strong><br>
            Column 1: Name (required)<br>
            Column 2: Event Name (required)<br>
            Column 3: Date (required) - Format: DD/MM/YYYY
        `;
    } else {
        bulkInfo.innerHTML = `
            <strong>Excel Format:</strong><br>
            Column 1: Name (required)<br>
            Column 2: Date (required) - Format: DD/MM/YYYY
        `;
    }
}

// Initial setup
setPreview("volunteer");

// EVENT LISTENERS
volunteerBtn.onclick = () => {
    clearSelections();
    eventInput.value = DEFAULT_EVENT_TEXT;
    setPreview("volunteer");
    document.querySelectorAll('.category').forEach(btn => btn.classList.remove('active'));
    volunteerBtn.classList.add('active');
};

vdpSelect.onchange = function() {
    appSelect.selectedIndex = 0;
    otherSelect.selectedIndex = 0;
    const eventName = this.value || "VDP";
    setPreview("vdp", eventName);
    document.querySelectorAll('.category').forEach(btn => btn.classList.remove('active'));
};

appSelect.onchange = function() {
    vdpSelect.selectedIndex = 0;
    otherSelect.selectedIndex = 0;
    const eventName = this.value || "Appreciation";
    setPreview("appreciation", eventName);
    document.querySelectorAll('.category').forEach(btn => btn.classList.remove('active'));
};

otherSelect.onchange = function() {
    vdpSelect.selectedIndex = 0;
    appSelect.selectedIndex = 0;
    const eventName = this.value || "Other";
    setPreview("other", eventName);
    document.querySelectorAll('.category').forEach(btn => btn.classList.remove('active'));
};

// Single/Bulk toggle
singleBtn.onclick = () => {
    singleBtn.classList.add('active');
    bulkBtn.classList.remove('active');
    singleSection.style.display = "flex";
    bulkSection.style.display = "none";
};

bulkBtn.onclick = () => {
    bulkBtn.classList.add('active');
    singleBtn.classList.remove('active');
    singleSection.style.display = "none";
    bulkSection.style.display = "block";
    updateBulkInfo(getType());
};

// Preview button
previewBtn.onclick = () => {
    fillPreview();
    document.querySelectorAll(".form input").forEach(x => x.disabled = true);
    previewBtn.style.display = "none";
    editBtn.style.display = "inline-block";
    generateSingle.style.display = "inline-block";
};

// Edit button
editBtn.onclick = () => {
    document.querySelectorAll(".form input").forEach(x => x.disabled = false);
    previewBtn.style.display = "inline-block";
    editBtn.style.display = "none";
    generateSingle.style.display = "none";
};

// Generate Single PDF
generateSingle.onclick = async () => {
    const type = getType();
    const name = document.getElementById("name").value;
    const rawDate = document.getElementById("date").value;
    const event = getEvent();
    const template = getSelectedSrc();

    if (!name || !rawDate) {
        alert("Please fill in Name and Date");
        return;
    }

    if (type !== "volunteer" && !event) {
        alert("Please select a certificate type");
        return;
    }

    const date = formatDate(rawDate);

    const payload = { type, template, name, event, date };
    console.log("Sending:", payload);

    try {
        const response = await fetch("/certificate/single", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(errorText);
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `${name}_certificate.pdf`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error("Generation error:", error);
        alert(`Failed to generate certificate: ${error.message}`);
    }
};

// Generate Bulk PDF
generateBulk.onclick = async () => {
    if (!excelFile.files || !excelFile.files[0]) {
        alert("Please select an Excel file");
        return;
    }

    const type = getType();
    const template = getSelectedSrc();
    const formData = new FormData();
    formData.append("excel", excelFile.files[0]);
    formData.append("type", type);
    formData.append("template", template);

    try {
        const response = await fetch("/certificate/bulk", {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(errorText);
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "certificates.zip";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error("Bulk generation error:", error);
        alert(`Failed to generate certificates: ${error.message}`);
    }
};

// Logout
logoutBtn.onclick = () => {
    location.href = "/logout";
};

// Live preview updates
document.getElementById("name").addEventListener("input", fillPreview);
document.getElementById("date").addEventListener("input", fillPreview);
document.getElementById("event").addEventListener("input", fillPreview);
vdpSelect.addEventListener("change", fillPreview);
appSelect.addEventListener("change", fillPreview);
otherSelect.addEventListener("change", fillPreview);