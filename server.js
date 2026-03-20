const express = require("express");
const compression = require("compression");
const path = require("path");
const fs = require("fs");
const os = require("os");
const crypto = require("crypto");
const sqlite3 = require("sqlite3").verbose();
const stripeSecretKey = process.env.STRIPE_SECRET_KEY;
const stripe = stripeSecretKey ? require("stripe")(stripeSecretKey) : null;

const app = express();
const port = process.env.PORT || 3000;
const configuredDataDir = process.env.DATA_DIR;
const configuredDbPath = process.env.DB_PATH;
const runtimeDataDir = path.join(os.tmpdir(), "grensesnittmatrise-runtime");
const dataDir = configuredDataDir
    ? path.resolve(configuredDataDir)
    : runtimeDataDir;
const dbPath = configuredDbPath ? path.resolve(configuredDbPath) : path.join(dataDir, "projects.db");
const dbDir = path.dirname(dbPath);
const SESSION_COOKIE_NAME = "gm_session";

if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
}

const db = new sqlite3.Database(dbPath);

db.serialize(() => {
    db.run(`
        CREATE TABLE IF NOT EXISTS projects (
            id TEXT PRIMARY KEY,
            data TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    `);
    db.run(`
        CREATE TABLE IF NOT EXISTS project_revisions (
            revision_id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT NOT NULL,
            data TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    `);
    db.run(`
        CREATE TABLE IF NOT EXISTS users (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            password_salt TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    `);
    db.run(`
        CREATE TABLE IF NOT EXISTS user_sessions (
            session_id TEXT PRIMARY KEY,
            user_id TEXT NOT NULL,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL
        )
    `);
});

function runQuery(sql, params = []) {
    return new Promise((resolve, reject) => {
        db.run(sql, params, function onRun(error) {
            if (error) {
                reject(error);
                return;
            }

            resolve(this);
        });
    });
}

function getQuery(sql, params = []) {
    return new Promise((resolve, reject) => {
        db.get(sql, params, (error, row) => {
            if (error) {
                reject(error);
                return;
            }

            resolve(row);
        });
    });
}

function parseCookies(req) {
    const cookieHeader = req.headers.cookie;

    if (!cookieHeader) {
        return {};
    }

    return cookieHeader.split(";").reduce((cookies, part) => {
        const [rawName, ...rawValueParts] = part.trim().split("=");

        if (!rawName) {
            return cookies;
        }

        cookies[rawName] = decodeURIComponent(rawValueParts.join("=") || "");
        return cookies;
    }, {});
}

function hashPassword(password, salt) {
    return crypto.scryptSync(password, salt, 64).toString("hex");
}

function createSessionCookie(sessionId) {
    const isProduction = process.env.NODE_ENV === "production";
    const parts = [
        `${SESSION_COOKIE_NAME}=${encodeURIComponent(sessionId)}`,
        "Path=/",
        "HttpOnly",
        "SameSite=Lax",
        "Max-Age=2592000",
    ];

    if (isProduction) {
        parts.push("Secure");
    }

    return parts.join("; ");
}

function clearSessionCookie() {
    const isProduction = process.env.NODE_ENV === "production";
    const parts = [
        `${SESSION_COOKIE_NAME}=`,
        "Path=/",
        "HttpOnly",
        "SameSite=Lax",
        "Max-Age=0",
    ];

    if (isProduction) {
        parts.push("Secure");
    }

    return parts.join("; ");
}

// ── Security headers ──
app.use((_req, res, next) => {
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("X-Frame-Options", "SAMEORIGIN");
    res.setHeader("X-XSS-Protection", "1; mode=block");
    res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
    res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=()");
    if (process.env.NODE_ENV === "production") {
        res.setHeader("Strict-Transport-Security", "max-age=31536000; includeSubDomains");
    }
    next();
});

// ── Page routes ──
app.get("/", (req, res) => res.sendFile(path.join(__dirname, "landing.html")));
app.get("/app", (req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.get("/login", (req, res) => res.sendFile(path.join(__dirname, "login.html")));
app.get("/favicon.ico", (_req, res) => res.status(204).end());
app.get("/robots.txt", (_req, res) => {
    res.type("text/plain");
    res.send("User-agent: *\nAllow: /\n");
});

// ── Gzip compression ──
app.use(compression());

app.use(express.json({ limit: "2mb" }));

// ── Static files with caching ──
app.use(express.static(__dirname, {
    index: false,
    maxAge: process.env.NODE_ENV === "production" ? "1d" : 0,
    etag: true,
}));

// ── Rate limiting for API ──
const rateLimitMap = new Map();
const RATE_LIMIT_WINDOW = 60_000;
const RATE_LIMIT_MAX = 120;

function rateLimit(req, res, next) {
    const ip = req.ip || req.connection.remoteAddress;
    const now = Date.now();
    const entry = rateLimitMap.get(ip);

    if (!entry || now - entry.start > RATE_LIMIT_WINDOW) {
        rateLimitMap.set(ip, { start: now, count: 1 });
        return next();
    }

    entry.count++;
    if (entry.count > RATE_LIMIT_MAX) {
        res.setHeader("Retry-After", Math.ceil((entry.start + RATE_LIMIT_WINDOW - now) / 1000));
        return res.status(429).json({ error: "For mange forespørsler. Prøv igjen senere." });
    }

    next();
}

app.use("/api", rateLimit);

// ── Health check ──
app.get("/health", (_req, res) => {
    res.json({ status: "ok", uptime: process.uptime() });
});

// ── API routes ──

app.post("/auth/register", async (req, res) => {
    try {
        const name = String(req.body?.name || "").trim();
        const email = String(req.body?.email || "").trim().toLowerCase();
        const password = String(req.body?.password || "");

        if (!name || !email || !password) {
            return res.status(400).json({ error: "Fyll inn navn, e-post og passord." });
        }

        if (!email.includes("@")) {
            return res.status(400).json({ error: "Ugyldig e-postadresse." });
        }

        if (password.length < 8) {
            return res.status(400).json({ error: "Passordet må ha minst 8 tegn." });
        }

        const existingUser = await getQuery("SELECT id FROM users WHERE email = ?", [email]);

        if (existingUser) {
            return res.status(409).json({ error: "Det finnes allerede en konto med denne e-posten." });
        }

        const userId = crypto.randomUUID();
        const salt = crypto.randomBytes(16).toString("hex");
        const passwordHash = hashPassword(password, salt);
        const createdAt = new Date().toISOString();

        await runQuery(
            `
            INSERT INTO users (id, name, email, password_hash, password_salt, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            `,
            [userId, name, email, passwordHash, salt, createdAt]
        );

        res.status(201).json({
            ok: true,
            user: {
                id: userId,
                name,
                email,
            },
        });
    } catch (error) {
        console.error("Register error:", error);
        res.status(500).json({ error: "Kunne ikke opprette konto." });
    }
});

app.post("/auth/login", async (req, res) => {
    try {
        const email = String(req.body?.email || "").trim().toLowerCase();
        const password = String(req.body?.password || "");

        if (!email || !password) {
            return res.status(400).json({ error: "Fyll inn e-post og passord." });
        }

        const user = await getQuery(
            `
            SELECT id, name, email, password_hash, password_salt
            FROM users
            WHERE email = ?
            `,
            [email]
        );

        if (!user) {
            return res.status(401).json({ error: "Feil e-post eller passord." });
        }

        const attemptedHash = hashPassword(password, user.password_salt);

        if (attemptedHash !== user.password_hash) {
            return res.status(401).json({ error: "Feil e-post eller passord." });
        }

        const sessionId = crypto.randomUUID();
        const createdAt = new Date();
        const expiresAt = new Date(createdAt.getTime() + 30 * 24 * 60 * 60 * 1000).toISOString();

        await runQuery(
            `
            INSERT INTO user_sessions (session_id, user_id, created_at, expires_at)
            VALUES (?, ?, ?, ?)
            `,
            [sessionId, user.id, createdAt.toISOString(), expiresAt]
        );

        res.setHeader("Set-Cookie", createSessionCookie(sessionId));
        res.json({
            ok: true,
            user: {
                id: user.id,
                name: user.name,
                email: user.email,
            },
        });
    } catch (error) {
        console.error("Login error:", error);
        res.status(500).json({ error: "Kunne ikke logge inn." });
    }
});

app.get("/auth/me", async (req, res) => {
    try {
        const cookies = parseCookies(req);
        const sessionId = cookies[SESSION_COOKIE_NAME];

        if (!sessionId) {
            return res.status(401).json({ error: "Ikke logget inn." });
        }

        const session = await getQuery(
            `
            SELECT users.id, users.name, users.email, user_sessions.expires_at
            FROM user_sessions
            JOIN users ON users.id = user_sessions.user_id
            WHERE user_sessions.session_id = ?
            `,
            [sessionId]
        );

        if (!session || new Date(session.expires_at).getTime() < Date.now()) {
            res.setHeader("Set-Cookie", clearSessionCookie());
            return res.status(401).json({ error: "Økten er utløpt." });
        }

        res.json({
            user: {
                id: session.id,
                name: session.name,
                email: session.email,
            },
        });
    } catch (error) {
        console.error("Auth me error:", error);
        res.status(500).json({ error: "Kunne ikke hente bruker." });
    }
});

app.post("/auth/logout", async (req, res) => {
    try {
        const cookies = parseCookies(req);
        const sessionId = cookies[SESSION_COOKIE_NAME];

        if (sessionId) {
            await runQuery("DELETE FROM user_sessions WHERE session_id = ?", [sessionId]);
        }

        res.setHeader("Set-Cookie", clearSessionCookie());
        res.json({ ok: true });
    } catch (error) {
        console.error("Logout error:", error);
        res.status(500).json({ error: "Kunne ikke logge ut." });
    }
});

app.get("/api/projects", (_req, res) => {
    db.all(
        "SELECT id, updated_at FROM projects ORDER BY datetime(updated_at) DESC",
        [],
        (error, rows) => {
            if (error) {
                res.status(500).json({ error: "Kunne ikke hente prosjektlisten." });
                return;
            }

            res.json({
                projects: rows.map((row) => ({
                    id: row.id,
                    updatedAt: row.updated_at,
                })),
            });
        }
    );
});

app.get("/api/projects/:id", (req, res) => {
    db.get(
        "SELECT id, data, updated_at FROM projects WHERE id = ?",
        [req.params.id],
        (error, row) => {
            if (error) {
                res.status(500).json({ error: "Kunne ikke hente prosjektet." });
                return;
            }

            if (!row) {
                res.status(404).json({ error: "Prosjektet finnes ikke ennå." });
                return;
            }

            res.json({
                id: row.id,
                updatedAt: row.updated_at,
                data: JSON.parse(row.data),
            });
        }
    );
});

app.get("/api/projects/:id/revisions", (req, res) => {
    db.all(
        `
        SELECT revision_id, created_at
        FROM project_revisions
        WHERE project_id = ?
        ORDER BY revision_id DESC
        LIMIT 20
        `,
        [req.params.id],
        (error, rows) => {
            if (error) {
                res.status(500).json({ error: "Kunne ikke hente versjonshistorikken." });
                return;
            }

            res.json({
                revisions: rows.map((row) => ({
                    revisionId: row.revision_id,
                    createdAt: row.created_at,
                })),
            });
        }
    );
});

app.get("/api/projects/:id/revisions/:revisionId", (req, res) => {
    db.get(
        `
        SELECT revision_id, data, created_at
        FROM project_revisions
        WHERE project_id = ? AND revision_id = ?
        `,
        [req.params.id, req.params.revisionId],
        (error, row) => {
            if (error) {
                res.status(500).json({ error: "Kunne ikke hente versjonen." });
                return;
            }

            if (!row) {
                res.status(404).json({ error: "Versjonen finnes ikke." });
                return;
            }

            res.json({
                revisionId: row.revision_id,
                createdAt: row.created_at,
                data: JSON.parse(row.data),
            });
        }
    );
});

app.put("/api/projects/:id", (req, res) => {
    const payload = req.body;

    if (!payload || typeof payload !== "object") {
        res.status(400).json({ error: "Ugyldig prosjektdata." });
        return;
    }

    const updatedAt = new Date().toISOString();

    const serializedPayload = JSON.stringify(payload);

    db.serialize(() => {
        db.run(
            `
            INSERT INTO projects (id, data, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                data = excluded.data,
                updated_at = excluded.updated_at
            `,
            [req.params.id, serializedPayload, updatedAt],
            (error) => {
                if (error) {
                    res.status(500).json({ error: "Kunne ikke lagre prosjektet." });
                    return;
                }

                db.run(
                    `
                    INSERT INTO project_revisions (project_id, data, created_at)
                    VALUES (?, ?, ?)
                    `,
                    [req.params.id, serializedPayload, updatedAt],
                    function onRevisionInsert(revisionError) {
                        if (revisionError) {
                            res.status(500).json({ error: "Prosjektet ble lagret, men versjonshistorikken kunne ikke oppdateres." });
                            return;
                        }

                        res.json({
                            ok: true,
                            id: req.params.id,
                            updatedAt,
                            revisionId: this.lastID,
                        });
                    }
                );
            }
        );
    });
});

app.delete("/api/projects/:id", (req, res) => {
    db.serialize(() => {
        db.run(
            "DELETE FROM project_revisions WHERE project_id = ?",
            [req.params.id],
            (revisionError) => {
                if (revisionError) {
                    res.status(500).json({ error: "Kunne ikke slette prosjektets versjonshistorikk." });
                    return;
                }

                db.run(
                    "DELETE FROM projects WHERE id = ?",
                    [req.params.id],
                    function onDelete(error) {
                        if (error) {
                            res.status(500).json({ error: "Kunne ikke slette prosjektet." });
                            return;
                        }

                        if (this.changes === 0) {
                            res.status(404).json({ error: "Prosjektet finnes ikke." });
                            return;
                        }

                        res.json({
                            ok: true,
                            id: req.params.id,
                        });
                    }
                );
            }
        );
    });
});

// ── Stripe checkout routes ──

app.post("/api/create-checkout-session", async (req, res) => {
    if (!stripe) {
        return res.status(503).json({ error: "Betaling er ikke konfigurert ennÃ¥." });
    }

    try {
        const session = await stripe.checkout.sessions.create({
            mode: "subscription",
            line_items: [
                {
                    price_data: {
                        currency: "nok",
                        product_data: {
                            name: "Grensesnittmatrise – Bedrift",
                        },
                        unit_amount: 100000,
                        recurring: {
                            interval: "month",
                        },
                    },
                    quantity: 1,
                },
            ],
            success_url: (process.env.BASE_URL || "http://localhost:3000") + "/app?session_id={CHECKOUT_SESSION_ID}",
            cancel_url: (process.env.BASE_URL || "http://localhost:3000") + "/",
        });

        res.json({ url: session.url });
    } catch (error) {
        console.error("Stripe checkout error:", error);
        res.status(500).json({ error: "Kunne ikke opprette betalingsøkt." });
    }
});

app.get("/api/subscription-status", async (req, res) => {
    if (!stripe) {
        return res.status(503).json({ error: "Betaling er ikke konfigurert ennÃ¥." });
    }

    try {
        const { session_id } = req.query;

        if (!session_id) {
            return res.status(400).json({ error: "Mangler session_id." });
        }

        const session = await stripe.checkout.sessions.retrieve(session_id);

        res.json({
            status: session.payment_status,
            customer_email: session.customer_details?.email,
        });
    } catch (error) {
        console.error("Stripe status error:", error);
        res.status(500).json({ error: "Kunne ikke hente abonnementsstatus." });
    }
});

// ── Graceful shutdown ──
function shutdown() {
    console.log("Stenger ned...");
    db.close(() => {
        process.exit(0);
    });
}

process.on("SIGTERM", shutdown);
process.on("SIGINT", shutdown);

app.listen(port, () => {
    console.log(`Server kjører på http://localhost:${port}`);
});
