const express = require("express");
const compression = require("compression");
const path = require("path");
const fs = require("fs");
const os = require("os");
const crypto = require("crypto");
const sqlite3 = require("sqlite3").verbose();
const { Resend } = require("resend");
const stripeSecretKey = process.env.STRIPE_SECRET_KEY;
const stripe = stripeSecretKey ? require("stripe")(stripeSecretKey) : null;
const resendApiKey = process.env.RESEND_API_KEY;
const resend = resendApiKey ? new Resend(resendApiKey) : null;
const BASE_URL = process.env.BASE_URL || "http://localhost:3000";

const app = express();
const port = process.env.PORT || 3000;
const configuredDataDir = process.env.DATA_DIR;
const configuredDbPath = process.env.DB_PATH;
const runtimeDataDir = path.join(os.tmpdir(), "grensesnittmatrise-runtime");
const dataDir = configuredDataDir
    ? path.resolve(configuredDataDir)
    : runtimeDataDir;
const staticDir = path.resolve(__dirname);
const landingPagePath = path.join(staticDir, "landing.html");
const appPagePath = path.join(staticDir, "index.html");
const loginPagePath = path.join(staticDir, "login.html");
const dbPath = configuredDbPath ? path.resolve(configuredDbPath) : path.join(dataDir, "projects.db");
const dbDir = path.dirname(dbPath);
const SESSION_COOKIE_NAME = "gm_session";

function escapeHtml(text) {
    return String(text ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}

function buildVerificationEmailHtml(name, verifyUrl) {
    return `
        <div style="font-family:sans-serif;max-width:500px;margin:0 auto;padding:32px">
            <h2 style="color:#0c6e70">Velkommen, ${escapeHtml(name)}!</h2>
            <p>Takk for at du opprettet en konto på Grensesnittmatrise.no.</p>
            <p>Klikk knappen under for å bekrefte e-postadressen din:</p>
            <a href="${verifyUrl}" style="display:inline-block;padding:14px 28px;background:#0c6e70;color:#fff;text-decoration:none;border-radius:8px;font-weight:700;margin:20px 0">Bekreft e-post</a>
            <p style="color:#666;font-size:0.85rem">Eller kopier denne lenken: ${escapeHtml(verifyUrl)}</p>
            <hr style="border:none;border-top:1px solid #eee;margin:24px 0">
            <p style="color:#999;font-size:0.8rem">Hvis du ikke opprettet denne kontoen, kan du ignorere denne e-posten.</p>
        </div>
    `;
}

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
            email_verified INTEGER NOT NULL DEFAULT 0,
            verify_token TEXT,
            created_at TEXT NOT NULL
        )
    `);
    // Migrate existing users table if columns are missing
    db.run(`ALTER TABLE users ADD COLUMN email_verified INTEGER NOT NULL DEFAULT 0`, () => {});
    db.run(`ALTER TABLE users ADD COLUMN verify_token TEXT`, () => {});
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

function sendPage(res, filePath) {
    res.sendFile(filePath, (error) => {
        if (!error) {
            return;
        }

        console.error("Send file error:", filePath, error);

        if (!res.headersSent) {
            res.status(error.statusCode || 500).send("Kunne ikke laste siden.");
        }
    });
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
app.get("/", (_req, res) => sendPage(res, landingPagePath));
app.get("/app", (_req, res) => sendPage(res, appPagePath));
app.get("/login", (_req, res) => sendPage(res, loginPagePath));
app.get("/auth/verify", async (req, res) => {
    try {
        const token = String(req.query.token || "");
        if (!token) {
            return res.status(400).send("Ugyldig lenke.");
        }

        const user = await getQuery("SELECT id, name FROM users WHERE verify_token = ?", [token]);
        if (!user) {
            return res.status(400).send("Ugyldig eller utløpt verifiseringslenke.");
        }

        await runQuery("UPDATE users SET email_verified = 1, verify_token = NULL WHERE id = ?", [user.id]);

        res.send(`
            <!DOCTYPE html>
            <html lang="no"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
            <title>E-post bekreftet</title>
            <style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f3ede2;margin:0}
            .card{background:#fff;padding:40px;border-radius:20px;text-align:center;max-width:400px;box-shadow:0 10px 40px rgba(0,0,0,0.1)}
            h1{color:#0c6e70;margin:0 0 12px}p{color:#555;margin:0 0 24px}
            a{display:inline-block;padding:12px 28px;background:#0c6e70;color:#fff;text-decoration:none;border-radius:10px;font-weight:700}</style>
            </head><body><div class="card">
            <h1>E-post bekreftet!</h1>
            <p>Hei ${escapeHtml(user.name)}, kontoen din er nå aktivert.</p>
            <a href="/login">Logg inn</a>
            </div></body></html>
        `);
    } catch (error) {
        console.error("Verify error:", error);
        res.status(500).send("Noe gikk galt.");
    }
});
app.get("/favicon.ico", (_req, res) => res.status(204).end());
app.get("/robots.txt", (_req, res) => {
    res.type("text/plain");
    res.send("User-agent: *\nAllow: /\n");
});

// ── Gzip compression ──
app.use(compression());

app.use(express.json({ limit: "2mb" }));

// ── Static files with caching ──
app.use(express.static(staticDir, {
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

// Stricter rate limit for auth (10 attempts per minute)
const authRateLimitMap = new Map();
const AUTH_RATE_LIMIT_WINDOW = 60_000;
const AUTH_RATE_LIMIT_MAX = 10;

function authRateLimit(req, res, next) {
    const ip = req.ip || req.connection.remoteAddress;
    const now = Date.now();
    const entry = authRateLimitMap.get(ip);

    if (!entry || now - entry.start > AUTH_RATE_LIMIT_WINDOW) {
        authRateLimitMap.set(ip, { start: now, count: 1 });
        return next();
    }

    entry.count++;
    if (entry.count > AUTH_RATE_LIMIT_MAX) {
        res.setHeader("Retry-After", Math.ceil((entry.start + AUTH_RATE_LIMIT_WINDOW - now) / 1000));
        return res.status(429).json({ error: "For mange forsøk. Vent litt og prøv igjen." });
    }

    next();
}

app.use("/auth", authRateLimit);

// ── Health check ──
app.get("/health", (_req, res) => {
    res.json({ status: "ok", uptime: process.uptime() });
});

// ── Auth middleware for project routes ──
async function requireAuth(req, res, next) {
    try {
        const cookies = parseCookies(req);
        const sessionId = cookies[SESSION_COOKIE_NAME];

        if (!sessionId) {
            return res.status(401).json({ error: "Ikke logget inn." });
        }

        const session = await getQuery(
            `SELECT users.id, users.name, users.email, user_sessions.expires_at
             FROM user_sessions
             JOIN users ON users.id = user_sessions.user_id
             WHERE user_sessions.session_id = ?`,
            [sessionId]
        );

        if (!session || new Date(session.expires_at).getTime() < Date.now()) {
            res.setHeader("Set-Cookie", clearSessionCookie());
            return res.status(401).json({ error: "Økten er utløpt." });
        }

        req.user = { id: session.id, name: session.name, email: session.email };
        next();
    } catch (error) {
        console.error("Auth middleware error:", error);
        res.status(500).json({ error: "Autentiseringsfeil." });
    }
}

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
        const verifyToken = crypto.randomBytes(32).toString("hex");
        const createdAt = new Date().toISOString();

        await runQuery(
            `
            INSERT INTO users (id, name, email, password_hash, password_salt, email_verified, verify_token, created_at)
            VALUES (?, ?, ?, ?, ?, 0, ?, ?)
            `,
            [userId, name, email, passwordHash, salt, verifyToken, createdAt]
        );

        // Send verification email
        if (resend) {
            const verifyUrl = `${BASE_URL}/auth/verify?token=${verifyToken}`;
            await resend.emails.send({
                from: "Grensesnittmatrise <noreply@grensesnittmatrise.no>",
                to: [email],
                subject: "Bekreft din e-postadresse – Grensesnittmatrise.no",
                html: buildVerificationEmailHtml(name, verifyUrl),
            });
        }

        res.status(201).json({
            ok: true,
            needsVerification: !!resend,
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

app.post("/auth/resend-verification", async (req, res) => {
    try {
        const email = String(req.body?.email || "").trim().toLowerCase();
        if (!email) {
            return res.status(400).json({ error: "Fyll inn e-post." });
        }

        const user = await getQuery(
            "SELECT id, name, email, email_verified, verify_token FROM users WHERE email = ?",
            [email]
        );

        if (!user) {
            return res.json({ ok: true });
        }

        if (user.email_verified) {
            return res.json({ ok: true, alreadyVerified: true });
        }

        let token = user.verify_token;
        if (!token) {
            token = crypto.randomBytes(32).toString("hex");
            await runQuery("UPDATE users SET verify_token = ? WHERE id = ?", [token, user.id]);
        }

        if (resend) {
            const verifyUrl = `${BASE_URL}/auth/verify?token=${token}`;
            await resend.emails.send({
                from: "Grensesnittmatrise <noreply@grensesnittmatrise.no>",
                to: [email],
                subject: "Bekreft din e-postadresse – Grensesnittmatrise.no",
                html: buildVerificationEmailHtml(user.name, verifyUrl),
            });
        }

        res.json({ ok: true });
    } catch (error) {
        console.error("Resend verification error:", error);
        res.status(500).json({ error: "Kunne ikke sende verifiserings-e-post." });
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
            SELECT id, name, email, password_hash, password_salt, email_verified
            FROM users
            WHERE email = ?
            `,
            [email]
        );

        if (!user) {
            return res.status(401).json({ error: "Feil e-post eller passord." });
        }

        if (resend && !user.email_verified) {
            return res.status(403).json({ error: "Du må bekrefte e-postadressen din først. Sjekk innboksen." });
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

app.get("/api/projects", requireAuth, (_req, res) => {
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

app.get("/api/projects/:id", requireAuth, (req, res) => {
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

app.get("/api/projects/:id/revisions", requireAuth, (req, res) => {
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

app.get("/api/projects/:id/revisions/:revisionId", requireAuth, (req, res) => {
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

app.put("/api/projects/:id", requireAuth, (req, res) => {
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

app.delete("/api/projects/:id", requireAuth, (req, res) => {
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
        return res.status(503).json({ error: "Betaling er ikke konfigurert ennå." });
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
        return res.status(503).json({ error: "Betaling er ikke konfigurert ennå." });
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

// ── Periodic cleanup ──
setInterval(() => {
    const now = Date.now();
    for (const [ip, entry] of rateLimitMap) {
        if (now - entry.start > RATE_LIMIT_WINDOW) rateLimitMap.delete(ip);
    }
    for (const [ip, entry] of authRateLimitMap) {
        if (now - entry.start > AUTH_RATE_LIMIT_WINDOW) authRateLimitMap.delete(ip);
    }
}, 5 * 60_000);

// Clean expired sessions every hour
setInterval(() => {
    db.run("DELETE FROM user_sessions WHERE datetime(expires_at) < datetime('now')", (error) => {
        if (error) console.error("Session cleanup error:", error);
    });
}, 60 * 60_000);

// ── Graceful shutdown ──
function shutdown() {
    console.log("Stenger ned...");
    db.close(() => {
        process.exit(0);
    });
}

process.on("SIGTERM", shutdown);
process.on("SIGINT", shutdown);
process.on("uncaughtException", (error) => {
    console.error("Uncaught exception:", error);
});
process.on("unhandledRejection", (reason) => {
    console.error("Unhandled rejection:", reason);
});

app.use((error, _req, res, _next) => {
    console.error("Express error:", error);
    res.status(500).json({ error: "Uventet serverfeil." });
});

app.listen(port, () => {
    console.log(`Server kjører på http://localhost:${port}`);
});

