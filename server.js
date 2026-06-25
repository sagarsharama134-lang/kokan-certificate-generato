const express = require("express");
const session = require("express-session");
const path = require("path");
const fs = require("fs");
const PORT = process.env.PORT || 3000;
const auth = require("./routes/auth");
const certificate = require("./routes/certificate");

const app = express();

app.use(express.json());
app.use(
    express.urlencoded({
        extended: true
    })
);

app.use(
    session({
        secret: "kokan",
        resave: false,
        saveUninitialized: false
    })
);

// Serve static files from public folder
app.use(
    express.static(
        path.join(__dirname, "public")
    )
);

// Serve assets from the root assets folder (outside public)
app.use(
    "/assets",
    express.static(
        path.join(
            __dirname,
            "assets"
        )
    )
);

// Also serve from public/assets as fallback (if needed)
app.use(
    "/assets",
    express.static(
        path.join(
            __dirname,
            "public",
            "assets"
        )
    )
);

// Create necessary directories if they don't exist
const directories = ['generated', 'uploads'];
directories.forEach(dir => {
    const dirPath = path.join(__dirname, dir);
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath);
        console.log(`Created directory: ${dir}`);
    }
});

app.use("/auth", auth);
app.use("/certificate", certificate);

/* LOGIN CHECK */
function requireLogin(req, res, next) {
    if (req.session && req.session.loggedIn) {
        return next();
    }
    res.redirect("/");
}

/* DASHBOARD (PROTECTED) */
app.get(
    "/dashboard",
    requireLogin,
    (req, res) => {
        res.sendFile(
            path.join(
                __dirname,
                "views",
                "dashboard.html"
            )
        );
    }
);

/* LOGIN PAGE */
app.get(
    "/",
    (req, res) => {
        res.sendFile(
            path.join(
                __dirname,
                "public",
                "login.html"
            )
        );
    }
);

/* LOGOUT */
app.get(
    "/logout",
    (req, res) => {
        req.session.destroy(
            () => {
                res.redirect("/");
            }
        );
    }
);

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});