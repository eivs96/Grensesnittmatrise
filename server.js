const express = require("express");
const compression = require("compression");
const path = require("path");
const fs = require("fs");
const os = require("os");
const sqlite3 = require("sqlite3").verbose();

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
});

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

// ── Gzip compression ──
app.use(compression());

app.use(express.json({ limit: "2mb" }));

// ── Static files with caching ──
app.use(express.static(__dirname, {
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
