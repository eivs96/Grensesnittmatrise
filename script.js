const defaultDisciplines = [
    "BH",
    "Byggfag",
    "Lås og beslag",
    "Dørleverandør",
    "Rør",
    "Vent",
    "EL",
    "Aut",
    "SD",
];

const disciplines = [
    ...new Set([
        ...(Array.isArray(window.EXCEL_DISCIPLINES) ? window.EXCEL_DISCIPLINES : defaultDisciplines),
        "SD",
    ]),
].filter((discipline) => discipline !== "ADK");

const responsibilities = ["P", "L", "M", "K", "F", "I"];
const uploadedDocuments = [];
let lastComplexityResult = null;
const packageControlledDisciplines = ["EL", "Aut", "SD", "Lås og beslag"];

const matrixBody = document.getElementById("matrix-body");
const projectIdInput = document.getElementById("project-id");
const projectTypeSelect = document.getElementById("project-type");
const packageOptionInputs = Array.from(document.querySelectorAll("#package-options input[type='checkbox']"));
const tueCoreModelSelect = document.getElementById("tue-core-model");
const tueLocksModelSelect = document.getElementById("tue-locks-model");
const tueAdkModelSelect = document.getElementById("tue-adk-model");
const tueStandaloneBuilder = document.getElementById("tue-standalone-builder");
const tueCompositionSummary = document.getElementById("tue-composition-summary");
const tueCoreHelp = document.getElementById("tue-core-help");
const tueLocksHelp = document.getElementById("tue-locks-help");
const tueAdkHelp = document.getElementById("tue-adk-help");
const tueRecommendation = document.getElementById("tue-recommendation");
const applyPackagePresetButton = document.getElementById("apply-package-preset");
const bhUploadInput = document.getElementById("bh-upload");
const analyzeBhButton = document.getElementById("analyze-bh");
const bhAnalysisStatus = document.getElementById("bh-analysis-status");
const bhAnalysisInsights = document.getElementById("bh-analysis-insights");
const projectLogicStatus = document.getElementById("project-logic-status");
const contractSummary = document.getElementById("contract-summary");
const refreshSummaryButton = document.getElementById("refresh-summary");
const saveProjectButton = document.getElementById("save-project");
const loadProjectButton = document.getElementById("load-project");
const exportExcelButton = document.getElementById("export-excel");
const exportPdfButton = document.getElementById("export-pdf");
const persistenceStatus = document.getElementById("persistence-status");
const workspaceReadinessLabel = document.getElementById("workspace-readiness-label");
const workspaceNextAction = document.getElementById("workspace-next-action");
const workspaceBlockers = document.getElementById("workspace-blockers");
const autosaveStatus = document.getElementById("autosave-status");
const projectList = document.getElementById("project-list");
const revisionList = document.getElementById("revision-list");
const refreshProjectListButton = document.getElementById("refresh-project-list");
const projectSearchInput = document.getElementById("project-search");
const projectLibraryStats = document.getElementById("project-library-stats");
const newProjectButton = document.getElementById("new-project");
const deleteProjectButton = document.getElementById("delete-project");
const currentRowTfm = document.getElementById("current-row-tfm");
const currentRowDescription = document.getElementById("current-row-description");
const currentRowRisk = document.getElementById("current-row-risk");
const currentRowMissing = document.getElementById("current-row-missing");
const currentRowConfirm = document.getElementById("current-row-confirm");
const currentRowConfirmText = document.getElementById("current-row-confirm-text");
const currentRowComment = document.getElementById("current-row-comment");
const quickConfirmRowButton = document.getElementById("quick-confirm-row");
const quickNextUnresolvedButton = document.getElementById("quick-next-unresolved");
const quickClearCommentButton = document.getElementById("quick-clear-comment");
const moveRowUpButton = document.getElementById("move-row-up");
const moveRowDownButton = document.getElementById("move-row-down");
const matrixSearchInput = document.getElementById("matrix-search");
const showOpenOnlyInput = document.getElementById("show-open-only");
const addRowButton = document.getElementById("add-row");
const deleteRowButton = document.getElementById("delete-row");
const jumpUnresolvedButton = document.getElementById("jump-unresolved");
const toggleReviewModeButton = document.getElementById("toggle-review-mode");
const reviewFilterButtons = Array.from(document.querySelectorAll("[data-review-filter]"));
const matrixVisibleCount = document.getElementById("matrix-visible-count");
const matrixVisibleDetail = document.getElementById("matrix-visible-detail");
const matrixConfirmedCount = document.getElementById("matrix-confirmed-count");
const matrixConfirmedDetail = document.getElementById("matrix-confirmed-detail");
const matrixOpenCount = document.getElementById("matrix-open-count");
const matrixOpenDetail = document.getElementById("matrix-open-detail");
const matrixFilterCount = document.getElementById("matrix-filter-count");
const matrixFilterStatus = document.getElementById("matrix-filter-status");
const matrixSectionCards = document.getElementById("matrix-section-cards");
const matrixSectionResetButton = document.getElementById("matrix-section-reset");
const matrixSectionFocusEyebrow = document.getElementById("matrix-section-focus-eyebrow");
const matrixSectionFocusTitle = document.getElementById("matrix-section-focus-title");
const matrixSectionFocusSummary = document.getElementById("matrix-section-focus-summary");
const matrixSectionFocusKpis = document.getElementById("matrix-section-focus-kpis");
const matrixSectionFocusThemes = document.getElementById("matrix-section-focus-themes");
const matrixSectionFocusRisks = document.getElementById("matrix-section-focus-risks");
const matrixSectionFocusDeliverables = document.getElementById("matrix-section-focus-deliverables");
const matrixSectionFirstRowButton = document.getElementById("matrix-section-first-row");
const matrixSectionNextOpenButton = document.getElementById("matrix-section-next-open");
const currentRowInsightSummary = document.getElementById("current-row-insight-summary");
const currentRowInsightDisciplines = document.getElementById("current-row-insight-disciplines");
const currentRowInsightFocus = document.getElementById("current-row-insight-focus");
const currentRowInsightDeliverables = document.getElementById("current-row-insight-deliverables");
const matrixEmptyState = document.getElementById("matrix-empty-state");
const workflowStepStatus = document.getElementById("workflow-step-status");
const workflowStepButtons = Array.from(document.querySelectorAll("[data-step-target]"));
const workflowTabs = Array.from(document.querySelectorAll(".workflow-step"));
const workflowPanels = Array.from(document.querySelectorAll(".step-panel"));
const workflowProgressValue = document.getElementById("workflow-progress-value");
const workflowProgressText = document.getElementById("workflow-progress-text");
const step1State = document.getElementById("step-1-state");
const step1Hint = document.getElementById("step-1-hint");
const step2State = document.getElementById("step-2-state");
const step2Hint = document.getElementById("step-2-hint");
const step3State = document.getElementById("step-3-state");
const step3Hint = document.getElementById("step-3-hint");
const step4State = document.getElementById("step-4-state");
const step4Hint = document.getElementById("step-4-hint");
const step1Checklist = document.getElementById("step-1-checklist");
const step2Checklist = document.getElementById("step-2-checklist");
const step3Checklist = document.getElementById("step-3-checklist");
const step4Checklist = document.getElementById("step-4-checklist");
const cockpitProgressValue = document.getElementById("cockpit-progress-value");
const cockpitProgressText = document.getElementById("cockpit-progress-text");
const cockpitNextStep = document.getElementById("cockpit-next-step");
const cockpitNextStepDetail = document.getElementById("cockpit-next-step-detail");
const cockpitMatrixHealth = document.getElementById("cockpit-matrix-health");
const cockpitMatrixHealthDetail = document.getElementById("cockpit-matrix-health-detail");
const cockpitOfferHealth = document.getElementById("cockpit-offer-health");
const cockpitOfferHealthDetail = document.getElementById("cockpit-offer-health-detail");
const matrixQueueList = document.getElementById("matrix-queue-list");
const matrixCommentGapCount = document.getElementById("matrix-comment-gap-count");
const matrixConflictCount = document.getElementById("matrix-conflict-count");
const matrixReviewReadyCount = document.getElementById("matrix-review-ready-count");
const matrixCommandDetail = document.getElementById("matrix-command-detail");
const jumpConflictRowButton = document.getElementById("jump-conflict-row");
const jumpUncommentedRowButton = document.getElementById("jump-uncommented-row");
const focusOfferStepButton = document.getElementById("focus-offer-step");

const defaultRows = [
    {
        tfm: "300",
        description: "Generelt - Rørtekniske installasjoner",
        comments: "Delansvar EL (overordnet): Tegne inn el-komponenter utenfor teknisk rom på tegning og i BIM.",
        marks: {},
        section: true,
    },
    {
        tfm: "300",
        description: "Pumper",
        comments: "Hovedpumpe(r) til hovedstokken må bestilles med kalender-ur funksjon.",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "EL:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Pumper med integrert frekvensomformer",
        comments: "Delansvar RØR: Innregulering (hvis direkte på pumpe).",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Ekstern frekvensomformer for pumper",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Rør:F": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Akkumulatortanker",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Trykkgiver (rør)",
        comments: "Delansvar RØR: Tegne inn komp. på systemskjema. Rørlegger leverer følerlommer.",
        marks: {
            "Rør:P": "D",
            "Rør:I": "H",
            "EL:F": "H",
            "EL:I": "H",
            "Aut:F": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Temperaturgiver (rør)",
        comments: "Delansvar RØR: Tegne inn komp. på systemskjema. Rørlegger leverer følerlommer.",
        marks: {
            "Rør:P": "D",
            "Rør:L": "H",
            "Rør:I": "H",
            "EL:F": "H",
            "EL:I": "H",
            "Aut:F": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Termisk energimåler",
        comments:
            "Delansvar RØR: Tegne inn komp. på systemskjema. Leverandør må levere dokumentasjon på måler rettidig. Rørlegger leverer følerlommer.",
        marks: {
            "Rør:P": "D",
            "Rør:L": "H",
            "Rør:I": "H",
            "EL:F": "H",
            "EL:I": "H",
            "Aut:F": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "To- og treveisventil inkl motor",
        comments: "",
        marks: {
            "Rør:P": "D",
            "Rør:I": "H",
            "EL:F": "H",
            "EL:I": "H",
            "Aut:F": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Ventiler (automatisk styrt)",
        comments: "Magnetventiler etc.",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Ventiler (for manuell betjening)",
        comments: "Generelt for ventiler i rørnett (sanitær, varme og kjøling).",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
        },
    },
    {
        tfm: "300",
        description: "Servicebryter for pumpe/frekvensomformer",
        comments: "Krav iht. NEK400",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:K": "H",
            "Rør:F": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Sanitæranlegg",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:K": "H",
            "Rør:F": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Vannmåler (inntak)",
        comments: "Leveres med bus for kommunikasjon mot SD.",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Vannmåler (generelt)",
        comments:
            "Delansvar RØR: Tegne inn komp. på systemskjema. Alle KV og VV målere leveres av Rør. Alle målere skal kommunisere med BAS-anlegget med IP-kommunikasjon (ModBus eller Bacnet).",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Tappevannssentral",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Vannbehandling legionella",
        comments: "Mulig det blir hettvannspyling via FV-veksler.",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Sirkulasjonspumpe VV (VVC)",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Sluk/drenering luftinntak",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Blandebatteri kjøkkenbenk (ikke storkjøkken)",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Sanitærutstyr og blandebatterier i plassbygde bad og WC",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Utslagsvask inkl. batteri (tekniske rom, BK, etc.)",
        comments: "",
        marks: {
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "H",
            "Rør:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Taksluk",
        comments: "Rørlegger prosjekterer og leverer sluk. Taktekker monterer og rørlegger tilkobler. Arkitekt lager fallplan.",
        marks: {
            "Byggfag:F": "H",
            "Rør:P": "H",
            "Rør:L": "H",
            "Rør:M": "D",
            "Rør:K": "H",
            "Rør:I": "H",
        },
    },
    {
        tfm: "310",
        description: "Spillvannspumpekum, komplett system",
        comments:
            "H Grunneentreprenør leverer. Plasseres i bakken utenfor bygget. Rørlegger tilkobler spillvannspumpekum. Elektro trekker kabler og kobler til pumper. Integrasjon til SD.",
        marks: {
            "Byggfag:P": "D",
            "Byggfag:L": "H",
            "Byggfag:M": "H",
            "Byggfag:F": "D",
            "Rør:L": "H",
            "Rør:I": "H",
            "Aut:I": "H",
        },
    },
];

let rowIdCounter = 0;

const sectionDefinitions = {
    100: "100 Generelt",
    200: "200 Bygningsdeler",
    300: "300 Sanitær og VVS",
    400: "400 Elektrofag",
    500: "500 Tele og automatisering",
    600: "600 Andre installasjoner",
    700: "700 Utendørs",
    800: "800 BREEAM-NOR v6",
};

const sectionCatalog = {
    100: {
        shortTitle: "Generelt",
        summary: "Prosjektomfattende premisser, koordinering og grensesnitt som pavirker flere fag samtidig.",
        themes: ["Koordinering", "Fellestegninger", "Tverrfaglige leveranser", "Prosjektkrav"],
        risks: ["Ingen tydelig eier av samordning", "Manglende leveransegrenser", "Uavklarte BIM- og tegningsansvar"],
        deliverables: ["Grensesnittstrategi", "Overordnet ansvarsdeling", "Felles premisser for UE-er"],
    },
    200: {
        shortTitle: "Byggfag",
        summary: "Bygningsdeler, rom, dorer, sjakter, utsparinger, innfesting og andre fysiske avklaringer mot tekniske fag.",
        themes: ["Dorer", "Utsparinger", "Sjakter", "Innfesting", "Plassbehov"],
        risks: ["Tekniske behov kolliderer med byggleveranse", "Manglende spikerslag eller innkassinger", "Dormiljo uten tydelig grensesnittboks"],
        deliverables: ["Tegninger for innfesting", "Utsparingsunderlag", "Samordnet dør- og romavklaring"],
    },
    300: {
        shortTitle: "VVS",
        summary: "Sanitar, varme, kjoling og rortekniske installasjoner med tilhorende komponenter, signaler og driftsavklaringer.",
        themes: ["Pumper", "Ventiler", "Sensorer", "Varmekabler", "SD-signaler"],
        risks: ["Uklart ansvar for givere og motorer", "Manglende avklaring mellom ROR, EL og AUT", "Driftssignaler ikke beskrevet"],
        deliverables: ["Systemskjema", "Komponentlister", "Signal- og funksjonsoversikt"],
    },
    400: {
        shortTitle: "Elektro",
        summary: "Kraft, fordelinger, foringsveier, belysning og elektriske grensesnitt mot bygg og tekniske leveranser.",
        themes: ["Fordelinger", "Foringsveier", "Belysning", "Kraft til utstyr", "Tavleplass"],
        risks: ["Utstyr mangler spenningssetting", "For liten tavleplass", "Foringsveier samordnes for sent"],
        deliverables: ["Kraftbehovsliste", "Tavlereservasjon", "Koordinert foringsveisplan"],
    },
    500: {
        shortTitle: "Automasjon",
        summary: "Tele, sikkerhet, automasjon og integrasjoner der flere systemer ma snakke sammen.",
        themes: ["SD/BAS", "ADK", "AIA/ABA", "KNX", "Systemintegrasjon"],
        risks: ["Systemer snakker ikke sammen", "Uklart ansvar for grensesnittboks", "Signalpunkter beskrives ulikt per fag"],
        deliverables: ["IO-lister", "Integrasjonsbeskrivelse", "Koordinerte koblingsskjema"],
    },
    600: {
        shortTitle: "Heis og spesial",
        summary: "Heis og andre spesialinstallasjoner med behov for avklaringer rundt plass, forsyning, signaler og ansvar.",
        themes: ["Heissjakt", "Heisfordeling", "Kortleser", "Alarmoverforing", "Maskinrom"],
        risks: ["Kabling til heis blir glemt", "Heisleveranse og elektro har ulike forutsetninger", "Adgang og alarm er ikke koordinert"],
        deliverables: ["Heisgrensesnitt", "Forsyningsavklaringer", "Signal- og kablingsplan"],
    },
    700: {
        shortTitle: "Utendors",
        summary: "Utvendige anlegg, forsyninger i grunn, utendors lys, VA og tekniske grensesnitt utenfor bygget.",
        themes: ["Utendors VA", "Lavspent forsyning", "Lys i grunn", "Automatisering ute", "Pumpekummer"],
        risks: ["Grensesnitt mot grunnentreprise er uklart", "IP- og SD-integrasjon beskrives ikke", "Kabel og ror i grunn mangler koordinering"],
        deliverables: ["Utomhus grensesnittplan", "Koordinert grunnunderlag", "Avklart ansvar for utvendig drift"],
    },
};

const rowInsightRules = [
    {
        keywords: ["pumpe", "pump"],
        focus: ["Spenningssetting", "Signal til SD/automasjon", "Innregulering og funksjonstest"],
        deliverables: ["Systemskjema", "Kraftbehov", "IO-/signalliste"],
    },
    {
        keywords: ["ventil", "motorventil", "spjeld"],
        focus: ["Hvem leverer aktuator", "Kabling og styringssignal", "Funksjonsansvar ved test"],
        deliverables: ["Komponentliste", "Koblingsskjema", "Funksjonsbeskrivelse"],
    },
    {
        keywords: ["giver", "sensor", "termostat", "trykkgiver", "temperaturgiver", "fuktsensor"],
        focus: ["Plassering", "Folerlommer eller montasjegrunnlag", "Signalpunkt og integrasjon"],
        deliverables: ["Tegningsgrunnlag", "Signaloversikt", "Grensesnitt mot automasjon"],
    },
    {
        keywords: ["dor", "kortleser", "adgang", "lås", "beslag", "grensesnittboks"],
        focus: ["Dormiljo og fysisk plass", "Kabling mellom fag", "Koordinert koblingsskjema"],
        deliverables: ["Dortyper og prinsipper", "Koblingsskjema", "Avklart ansvar for AAK/AIA/ABA"],
    },
    {
        keywords: ["heis", "heissjakt", "heisstol", "heismaskin"],
        focus: ["Forsyning og reserver", "Adgang, alarm og kommunikasjon", "Kabling til heisleveranse"],
        deliverables: ["Heisgrensesnitt", "Kabelplan", "Avklaring mot heisleverandor"],
    },
    {
        keywords: ["utendors", "grunn", "kum", "va", "lavspent", "lys"],
        focus: ["Grensesnitt mot grunnentreprise", "Koordinering av ror og kabel i grunn", "Drifts- og integrasjonsbehov utendors"],
        deliverables: ["Utomhusplan", "Koordinert grunnunderlag", "Ansvarsdeling for utvendige installasjoner"],
    },
    {
        keywords: ["fordeling", "tavle", "kraft"],
        focus: ["Tavleplass", "Reserver og kapasitet", "Hvem som spenningssetter hva"],
        deliverables: ["Enlinjeskjema", "Kraftbehovsliste", "Plassavsetning i tavle"],
    },
    {
        keywords: ["varmekabel", "snosmelting"],
        focus: ["Styringsprinsipp", "Foletyper og plassering", "Samspill mellom EL og AUT"],
        deliverables: ["Styringsbeskrivelse", "Varmekabelplan", "Vær- eller bakkefølerstrategi"],
    },
];

function createRowId() {
    rowIdCounter += 1;
    return `row-${rowIdCounter}`;
}

function inferSectionCode(tfmValue) {
    const tfmText = String(tfmValue || "").replace(/\s+/g, " ").trim();
    const match = tfmText.match(/\d{2,4}/);

    if (!match) {
        return 100;
    }

    const numericCode = Number.parseInt(match[0], 10);

    if (numericCode < 200) {
        return 100;
    }

    if (numericCode >= 700) {
        return 700;
    }

    return Math.floor(numericCode / 100) * 100;
}

function getSectionDetails(sectionCode) {
    return sectionCatalog[sectionCode] || {
        shortTitle: sectionDefinitions[sectionCode] || String(sectionCode),
        summary: "Samlet kategori for relaterte grensesnitt i matrisen.",
    };
}

function getRowSectionCode(row) {
    return inferSectionCode(row?.tfm);
}

function normalizeTfmValue(tfmValue) {
    return String(tfmValue || "")
        .replace(/\s*\n+\s*/g, "/")
        .replace(/\s{2,}/g, " ")
        .trim();
}

function getNormalizedRowTfm(row) {
    const normalizedTfm = normalizeTfmValue(row.tfm);
    const normalizedDescription = String(row.description || "").trim().toLowerCase();

    if (normalizedDescription === "utsparinger i betong") {
        return "200";
    }

    if (normalizedDescription === "dører") {
        return "234/244";
    }

    return normalizedTfm;
}

function getPrimaryTfmCode(tfmValue) {
    const match = normalizeTfmValue(tfmValue).match(/\d{2,4}/);
    return match ? Number.parseInt(match[0], 10) : 0;
}

function normalizeRowsByTfm(inputRows) {
    const contentRows = inputRows
        .filter((row) => !row.section)
        .map((row) => ({
            ...row,
            uid: row.uid || createRowId(),
            tfm: getNormalizedRowTfm(row),
            comments: row.comments || "",
            marks: { ...(row.marks || {}) },
            section: false,
        }));
    const groupedRows = new Map();

    Object.keys(sectionDefinitions).forEach((key) => {
        groupedRows.set(Number(key), []);
    });

    contentRows.forEach((row) => {
        const sectionCode = inferSectionCode(row.tfm);
        if (!groupedRows.has(sectionCode)) {
            groupedRows.set(sectionCode, []);
        }
        groupedRows.get(sectionCode).push(row);
    });

    const normalizedRows = [];
    Array.from(groupedRows.keys()).sort((left, right) => left - right).forEach((sectionCode) => {
        const sectionRows = (groupedRows.get(sectionCode) || []).sort((left, right) => {
            const sortDifference = getPrimaryTfmCode(left.tfm) - getPrimaryTfmCode(right.tfm);
            if (sortDifference !== 0) {
                return sortDifference;
            }

            return left.description.localeCompare(right.description, "no");
        });
        if (!sectionRows.length) {
            return;
        }

        normalizedRows.push({
            uid: `section-${sectionCode}`,
            tfm: String(sectionCode),
            description: sectionDefinitions[sectionCode] || `${sectionCode}`,
            comments: "",
            marks: {},
            section: true,
            autogeneratedSection: true,
        });

        normalizedRows.push(...sectionRows);
    });

    return normalizedRows;
}

let rows = normalizeRowsByTfm(defaultRows);

const stateOrder = ["", "H", "D"];
const confirmationState = new Map();
const commentState = new Map();
let baseMarksByRow = rows.map((row) => ({ ...row.marks }));
const collapsedSections = new Map();
let uploadedBhText = "";
let focusedRowIndex = -1;
let autosaveTimer = null;
let isApplyingSavedState = false;
let isSavingProject = false;
const LAST_PROJECT_KEY = "grensesnittmatrise:last-project";
const REVIEW_MODE_KEY = "grensesnittmatrise:review-mode";
const REVIEW_FILTER_KEY = "grensesnittmatrise:review-filter";
let activeRowIndex = -1;
let cachedProjects = [];
let cachedRevisions = [];
let lastBhAnalysis = null;
let currentWorkflowStep = 1;
let matrixInitialized = false;
let matrixDataPromise = null;
let usingImportedBaseRows = false;
let hasProjectSpecificRows = false;
let matrixBuildInProgress = false;
let activeSectionFilter = "all";
let reviewModeEnabled = false;
let activeReviewFilter = "all";
const uploadedOfferDocuments = [];
let lastOfferAnalysis = null;

function getSectionKey(row) {
    return `${row.tfm}|${row.description}`;
}

const packageLabels = {
    sd: "SD separat",
    el: "EL separat",
    aut: "AUT separat",
    las: "Lås og beslag separat",
    el_aut: "EL + AUT",
    el_aut_sd: "EL + AUT + SD",
    totaltechnical: "Totalteknisk pakke",
};

function getTueConfig() {
    return {
        coreModel: tueCoreModelSelect?.value || "separate",
        locksModel: tueLocksModelSelect?.value || "separate",
        adkModel: tueAdkModelSelect?.value || "el",
        standaloneDisciplines: packageOptionInputs
            .filter((input) => input.checked)
            .map((input) => input.value),
    };
}

function describeTueConfig(config = getTueConfig()) {
    const coreDescriptions = {
        separate: "Separate tekniske UE-er",
        el_aut: "EL + AUT i felles pakke",
        el_aut_sd: "EL + AUT + SD i felles pakke",
        totaltechnical: "Totalteknisk pakke",
    };
    const parts = [coreDescriptions[config.coreModel] || coreDescriptions.separate];

    if (config.coreModel === "separate" && config.standaloneDisciplines.length) {
        parts.push(`Egne UE-er: ${config.standaloneDisciplines.map((key) => packageLabels[key]).join(", ")}`);
    }

    parts.push(
        config.locksModel === "separate"
            ? "Lås og beslag som egen UE"
            : "Lås og beslag integrert i dør-/byggleveranse"
    );
    parts.push(
        config.adkModel === "el"
            ? "ADK i elektrikerleveransen"
            : "ADK i lås og beslagsleveransen"
    );

    return parts.join(" • ");
}

function getTueGuidance(config = getTueConfig()) {
    const coreHelpText = {
        separate: "Best når EL, AUT og SD konkurranseutsettes eller styres hver for seg.",
        el_aut: "Passer når elektro og automasjon jobber tett og leveres som én teknisk pakke.",
        el_aut_sd: "Passer når styring, automasjon og SD skal samordnes i én leveranse.",
        totaltechnical: "Best når hele det tekniske omfanget skal styres som én samlet kontrakt.",
    };
    const recommendationText = {
        separate: "Velg separate UE-er når prosjektet trenger mest mulig fleksibilitet og tydelige faggrenser.",
        el_aut: "Velg EL + AUT når integrasjoner er viktige, men SD fortsatt ønskes som tydelig eget grensesnitt.",
        el_aut_sd: "Velg EL + AUT + SD når du vil minimere koordinering mellom tekniske styringsfag.",
        totaltechnical: "Velg totalteknisk når du vil redusere grensesnitt og legge helhetsansvar hos én aktør.",
    };

    return {
        coreHelp: coreHelpText[config.coreModel] || coreHelpText.separate,
        locksHelp:
            config.locksModel === "separate"
                ? "Brukes når lås og beslag kontraheres og følges opp som eget fag."
                : "Brukes når dørleveranse og beslag håndteres samlet i bygg- eller dørentreprisen.",
        adkHelp:
            config.adkModel === "el"
                ? "Velg dette når adgangskontroll prosjekteres og leveres sammen med elektro."
                : "Velg dette når adgangskontroll følger dørmiljø, beslag og låsleveranse.",
        recommendation: recommendationText[config.coreModel] || recommendationText.separate,
    };
}

function syncTueBuilderUI() {
    const config = getTueConfig();
    const guidance = getTueGuidance(config);

    if (tueStandaloneBuilder) {
        tueStandaloneBuilder.hidden = config.coreModel !== "separate";
    }

    packageOptionInputs.forEach((input) => {
        input.disabled = config.coreModel !== "separate";
        if (config.coreModel !== "separate") {
            input.checked = false;
        }
    });

    if (tueCompositionSummary) {
        tueCompositionSummary.textContent = describeTueConfig(getTueConfig());
    }
    if (tueCoreHelp) {
        tueCoreHelp.textContent = guidance.coreHelp;
    }
    if (tueLocksHelp) {
        tueLocksHelp.textContent = guidance.locksHelp;
    }
    if (tueAdkHelp) {
        tueAdkHelp.textContent = guidance.adkHelp;
    }
    if (tueRecommendation) {
        tueRecommendation.textContent = guidance.recommendation;
    }
}

function getSelectedPackages() {
    const config = getTueConfig();
    const packages = [];

    if (config.coreModel === "separate") {
        packages.push(...config.standaloneDisciplines);
    } else {
        packages.push(config.coreModel);
    }

    if (config.locksModel === "separate") {
        packages.push("las");
    }

    return packages;
}

function applyState(button, state) {
    button.dataset.state = state;
    button.classList.remove("active", "state-d");

    if (state === "H") {
        button.classList.add("active");
    }

    if (state === "D") {
        button.classList.add("state-d");
    }

    button.textContent = state;
    button.setAttribute("aria-pressed", state === "" ? "false" : "true");
    button.setAttribute("aria-label", `${button.dataset.row ? rows[Number(button.dataset.row)]?.description || "" : ""} - ${button.dataset.discipline || ""} - ${button.dataset.responsibility || ""}${state ? ` - ${state}` : " - tom"}`);
}

function nextState(currentState) {
    const currentIndex = stateOrder.indexOf(currentState);
    return stateOrder[(currentIndex + 1) % stateOrder.length];
}

function getResponsibilityState(rowIndex, responsibility) {
    const rowButtons = matrixBody.querySelectorAll(
        `button[data-row="${rowIndex}"][data-responsibility="${responsibility}"]`
    );

    const selectedButton = Array.from(rowButtons).find((button) => button.dataset.state !== "");
    return selectedButton ? selectedButton.dataset.state : "";
}

function getRiskState(rowIndex) {
    const missingResponsibilities = responsibilities.filter(
        (responsibility) => getResponsibilityState(rowIndex, responsibility) === ""
    );

    if (missingResponsibilities.length > 0) {
        return { level: "warning", icon: "🟠", title: "Uklart grensesnitt" };
    }

    if (!confirmationState.get(rowIndex)) {
        return { level: "warning", icon: "🟠", title: "UE ikke bekreftet" };
    }

    return { level: "ok", icon: "🟢", title: "OK" };
}

function getContentRowCount() {
    return rows.filter((row) => !row.section).length;
}

function getConfirmedRowCount() {
    return rows.filter((row, rowIndex) => !row.section && confirmationState.get(rowIndex)).length;
}

function getOpenRiskCount() {
    return rows.filter((row, rowIndex) => !row.section && getRiskState(rowIndex).level !== "ok").length;
}

function getExportRiskLabel(rowIndex) {
    const risk = getRiskState(rowIndex);
    return risk.title || "Uavklart";
}

function getMissingResponsibilities(rowIndex) {
    return responsibilities.filter((responsibility) => getResponsibilityState(rowIndex, responsibility) === "");
}

function buildExportHighlights() {
    const totalRows = getContentRowCount();
    const confirmedCount = getConfirmedRowCount();
    const openRiskCount = getOpenRiskCount();
    const commentedCount = rows.filter(
        (row, rowIndex) => !row.section && Boolean((commentState.get(rowIndex) ?? row.comments ?? "").trim())
    ).length;
    const completionRate = totalRows ? Math.round((confirmedCount / totalRows) * 100) : 0;

    return {
        totalRows,
        confirmedCount,
        openRiskCount,
        commentedCount,
        completionRate,
    };
}

function buildExportActionItems() {
    const items = [];
    const openRows = rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(({ row, rowIndex }) => !row.section && getRiskState(rowIndex).level !== "ok")
        .slice(0, 6);

    if (!openRows.length) {
        items.push("Alle registrerte rader er avklart og UE-bekreftet.");
        return items;
    }

    openRows.forEach(({ row, rowIndex }) => {
        items.push(`${row.tfm} ${row.description} - ${getExportRiskLabel(rowIndex)}`);
    });

    return items;
}

function updateMatrixOverview(visibleContentRows = null) {
    const totalContentRows = getContentRowCount();
    const confirmedCount = getConfirmedRowCount();
    const openRiskCount = getOpenRiskCount();
    const visibleCount = visibleContentRows ?? totalContentRows;
    const completionRate = totalContentRows ? Math.round((confirmedCount / totalContentRows) * 100) : 0;

    if (matrixVisibleCount) {
        matrixVisibleCount.textContent = String(totalContentRows);
    }

    if (matrixVisibleDetail) {
        matrixVisibleDetail.textContent = `${visibleCount} synlige`;
    }

    if (matrixConfirmedCount) {
        matrixConfirmedCount.textContent = String(confirmedCount);
    }

    if (matrixConfirmedDetail) {
        matrixConfirmedDetail.textContent = `${completionRate} % ferdig`;
    }

    if (matrixOpenCount) {
        matrixOpenCount.textContent = String(openRiskCount);
    }

    if (matrixOpenDetail) {
        matrixOpenDetail.textContent = openRiskCount === 0 ? "Ingen åpne punkter" : "Åpne punkter";
    }

    renderMatrixSectionCards();
    renderMatrixSectionFocusPanel();
    updateMatrixCommandCenter();
}

function updateMatrixCommandCenter() {
    const commentGaps = getRowsNeedingComment();
    const conflictRows = rows.filter((row) => !row.section && getOfferConflictRowIds().has(row.uid));
    const reviewReadyCount = getReviewReadyCount();
    const queueItems = [
        conflictRows[0] ? `${conflictRows[0].tfm} ${conflictRows[0].description} - tilbudskonflikt bor vurderes.` : "",
        commentGaps[0] ? `${commentGaps[0].row.tfm} ${commentGaps[0].row.description} - mangler kommentar pa apen avklaring.` : "",
        getOpenRiskCount() > 0 ? `${getOpenRiskCount()} rad(er) star fortsatt apne i matrisen.` : "Ingen apne avklaringer igjen i matrisen.",
    ].filter(Boolean);

    if (matrixCommentGapCount) {
        matrixCommentGapCount.textContent = String(commentGaps.length);
    }
    if (matrixConflictCount) {
        matrixConflictCount.textContent = String(conflictRows.length);
    }
    if (matrixReviewReadyCount) {
        matrixReviewReadyCount.textContent = String(reviewReadyCount);
    }
    if (matrixCommandDetail) {
        matrixCommandDetail.textContent = conflictRows.length
            ? "Tilbudskontrollen har funnet rader som bor gjennomgas direkte i matrisen."
            : "Bruk arbeidskoen til a lukke apne punkter og dokumentere vurderingene dine.";
    }
    if (matrixQueueList) {
        matrixQueueList.innerHTML = queueItems.map((item) => `<p>${escapeHtml(item)}</p>`).join("");
    }
    if (jumpConflictRowButton) {
        jumpConflictRowButton.disabled = !conflictRows.length;
    }
    if (jumpUncommentedRowButton) {
        jumpUncommentedRowButton.disabled = !commentGaps.length;
    }
}

function updateMatrixFilterFeedback(visibleCount, query, openOnly) {
    const filterParts = [];

    if (query) {
        filterParts.push(`sok: "${query}"`);
    }

    if (openOnly) {
        filterParts.push("kun apne");
    }

    if (activeSectionFilter !== "all") {
        filterParts.push(getSectionDetails(Number(activeSectionFilter)).shortTitle.toLowerCase());
    }

    if (activeReviewFilter !== "all") {
        filterParts.push(`review: ${getReviewFilterLabel().toLowerCase()}`);
    }

    if (matrixFilterCount) {
        matrixFilterCount.textContent = filterParts.length ? String(visibleCount) : "Alle";
    }

    if (matrixFilterStatus) {
        matrixFilterStatus.textContent = filterParts.length
            ? `${visibleCount} treff med ${filterParts.join(" + ")}`
            : "Ingen filter aktivt";
    }

    if (matrixEmptyState) {
        matrixEmptyState.hidden = visibleCount > 0;
        const emptyStateMessage = matrixEmptyState.querySelector("p");

        if (emptyStateMessage) {
            if (activeReviewFilter === "conflicts") {
                emptyStateMessage.textContent = "Ingen konfliktrader er synlige akkurat nå. Kjør tilbudskontroll eller bytt til en annen arbeidsvisning.";
            } else if (activeReviewFilter === "confirmed") {
                emptyStateMessage.textContent = "Ingen bekreftede rader matcher filtreringen akkurat nå. Bekreft rader eller bytt arbeidsvisning.";
            } else if (activeReviewFilter === "open") {
                emptyStateMessage.textContent = "Ingen åpne rader matcher filtreringen akkurat nå. Det kan bety at utvalget er ferdig gjennomgått.";
            } else if (activeSectionFilter !== "all") {
                emptyStateMessage.textContent = "Ingen rader matcher valgt kategori og aktiv filtrering akkurat nå. Prov a vise hele matrisen eller nullstille soket.";
            } else {
                emptyStateMessage.textContent = "Ingen rader matcher filtreringen akkurat nå. Tøm søket eller slå av filteret for åpne avklaringer.";
            }
        }
    }
}

function getSectionStats(sectionCode) {
    let total = 0;
    let confirmed = 0;
    let open = 0;

    rows.forEach((row, rowIndex) => {
        if (row.section || getRowSectionCode(row) !== sectionCode) {
            return;
        }

        total += 1;

        if (confirmationState.get(rowIndex)) {
            confirmed += 1;
        }

        if (getRiskState(rowIndex).level !== "ok") {
            open += 1;
        }
    });

    return { total, confirmed, open };
}

function renderTagList(container, items, fallbackText) {
    if (!container) {
        return;
    }

    const values = Array.isArray(items) ? items.filter(Boolean) : [];

    if (!values.length) {
        container.innerHTML = `<span>${escapeHtml(fallbackText)}</span>`;
        return;
    }

    container.innerHTML = values
        .map((item) => `<span>${escapeHtml(item)}</span>`)
        .join("");
}

function renderMatrixSectionFocusPanel() {
    if (!matrixSectionFocusTitle || !matrixSectionFocusKpis) {
        return;
    }

    if (activeSectionFilter === "all") {
        const totalContentRows = getContentRowCount();
        const confirmedCount = getConfirmedRowCount();
        const openCount = getOpenRiskCount();
        const completionRate = totalContentRows ? Math.round((confirmedCount / totalContentRows) * 100) : 0;

        if (matrixSectionFocusEyebrow) matrixSectionFocusEyebrow.textContent = "Fokusmodus";
        if (matrixSectionFocusTitle) matrixSectionFocusTitle.textContent = "Hele matrisen";
        if (matrixSectionFocusSummary) {
            matrixSectionFocusSummary.textContent = "Du ser alle fagomrader samlet. Velg en kategori over for a jobbe mer konsentrert med ett omrade og raskere lukke apne avklaringer.";
        }

        matrixSectionFocusKpis.innerHTML = `
            <div class="overview-card"><span class="overview-label">Totalt</span><strong>${totalContentRows}</strong><span class="overview-detail">Rader i prosjektet</span></div>
            <div class="overview-card"><span class="overview-label">Bekreftet</span><strong>${confirmedCount}</strong><span class="overview-detail">${completionRate} % ferdig</span></div>
            <div class="overview-card"><span class="overview-label">Apne</span><strong>${openCount}</strong><span class="overview-detail">Tverrfaglige avklaringer</span></div>
            <div class="overview-card"><span class="overview-label">Anbefaling</span><strong>Velg kategori</strong><span class="overview-detail">Jobb en del av bygget av gangen</span></div>
        `;

        renderTagList(matrixSectionFocusThemes, ["Start med storste apne seksjon", "Lukk grasoner fortlopende", "Bruk kommentarer for forbehold"], "Ingen tema valgt");
        renderTagList(matrixSectionFocusRisks, ["For bred arbeidsflate gir treg gjennomgang", "Apen matrisen kan skjule hvor risikoen ligger"], "Ingen risiko valgt");
        renderTagList(matrixSectionFocusDeliverables, ["Kategoriavklart matrise", "Eksportgrunnlag med tydelige UE-grenser", "Kort vei til neste apne punkt"], "Ingen leveranser valgt");
        return;
    }

    const sectionCode = Number(activeSectionFilter);
    const details = getSectionDetails(sectionCode);
    const stats = getSectionStats(sectionCode);
    const completionRate = stats.total ? Math.round((stats.confirmed / stats.total) * 100) : 0;
    const shareOfMatrix = getContentRowCount() ? Math.round((stats.total / getContentRowCount()) * 100) : 0;

    if (matrixSectionFocusEyebrow) matrixSectionFocusEyebrow.textContent = `Kategori ${sectionCode}`;
    if (matrixSectionFocusTitle) matrixSectionFocusTitle.textContent = `${sectionCode} ${details.shortTitle}`;
    if (matrixSectionFocusSummary) {
        matrixSectionFocusSummary.textContent = details.summary;
    }

    matrixSectionFocusKpis.innerHTML = `
        <div class="overview-card"><span class="overview-label">Rader</span><strong>${stats.total}</strong><span class="overview-detail">${shareOfMatrix} % av matrisen</span></div>
        <div class="overview-card"><span class="overview-label">Bekreftet</span><strong>${stats.confirmed}</strong><span class="overview-detail">${completionRate} % ferdig</span></div>
        <div class="overview-card"><span class="overview-label">Apne</span><strong>${stats.open}</strong><span class="overview-detail">${stats.open ? "Bør lukkes før eksport" : "Ingen apne punkt"}</span></div>
        <div class="overview-card"><span class="overview-label">Arbeidsmodus</span><strong>Fokus</strong><span class="overview-detail">Viser kun valgt kategori</span></div>
    `;

    renderTagList(matrixSectionFocusThemes, details.themes, "Legg til faglige tema for denne kategorien");
    renderTagList(matrixSectionFocusRisks, details.risks, "Legg til typiske grasoner for denne kategorien");
    renderTagList(matrixSectionFocusDeliverables, details.deliverables, "Legg til forventede leveranser for denne kategorien");
}

function getVisibleContentRowIndexes({ openOnly = false } = {}) {
    return rows
        .map((row, rowIndex) => ({ row, rowIndex, element: getRowElement(rowIndex) }))
        .filter(({ row, rowIndex, element }) => {
            if (row.section || !element || element.classList.contains("filtered-out")) {
                return false;
            }

            if (openOnly && getRiskState(rowIndex).level === "ok") {
                return false;
            }

            return true;
        })
        .map(({ rowIndex }) => rowIndex);
}

function focusFirstVisibleContentRow(options = {}) {
    const indexes = getVisibleContentRowIndexes(options);

    if (!indexes.length) {
        showToast("Ingen synlige rader a hoppe til i dette utvalget.", "info");
        return;
    }

    focusRow(indexes[0]);
}

function uniqueList(items) {
    return [...new Set(items.filter(Boolean))];
}

function getDisciplinesForRow(row) {
    return uniqueList(
        Object.keys(row?.marks || {})
            .map((key) => key.split(":")[0])
            .filter(Boolean)
    );
}

function matchInsightRules(text) {
    return rowInsightRules.filter((rule) => rule.keywords.some((keyword) => text.includes(keyword)));
}

function getRowInsightData(row, rowIndex) {
    const sectionDetails = getSectionDetails(getRowSectionCode(row));
    const text = `${row.tfm} ${row.description} ${commentState.get(rowIndex) ?? row.comments ?? ""}`.toLowerCase();
    const matchedRules = matchInsightRules(text);
    const disciplinesInRow = getDisciplinesForRow(row);
    const sectionFocus = {
        100: ["Overordnet ansvar", "Tverrfaglig koordinering", "Felles prosjektpremisser"],
        200: ["Fysisk plass og innfesting", "Utsparinger og sjakter", "Bygg mot tekniske fag"],
        300: ["ROR mot EL/AUT", "Signalpunkter", "Funksjon og innregulering"],
        400: ["Kraft og reserveplass", "Foringsveier", "Spenningssetting av teknisk utstyr"],
        500: ["Systemintegrasjon", "Koblingsskjema", "Ansvar mellom sikkerhet og automasjon"],
        600: ["Spesialleveranse", "Forsyning og kommunikasjon", "Koordinering mot ekstern leverandor"],
        700: ["Grunnentreprise", "Utvendig drift", "Koordinering i grunn"],
    };
    const sectionDeliverables = {
        100: ["Koordineringsnotat", "Ansvarsdeling", "Prosjektpremisser"],
        200: ["Utsparingsunderlag", "Innfestingsgrunnlag", "Samordnet dortegning"],
        300: ["Systemskjema", "Komponentliste", "Signaloversikt"],
        400: ["Kraftbehovsliste", "Tavleunderlag", "Foringsveisplan"],
        500: ["IO-liste", "Integrasjonsbeskrivelse", "Koblingsskjema"],
        600: ["Grensesnittnotat", "Kabelplan", "Leverandoravklaring"],
        700: ["Utomhusplan", "Koordinert grunnplan", "Ansvarsdeling ute"],
    };

    const focus = uniqueList([
        ...(sectionFocus[getRowSectionCode(row)] || []),
        ...matchedRules.flatMap((rule) => rule.focus || []),
        ...getMissingResponsibilities(rowIndex).slice(0, 3),
    ]).slice(0, 6);

    const deliverables = uniqueList([
        ...(sectionDeliverables[getRowSectionCode(row)] || []),
        ...matchedRules.flatMap((rule) => rule.deliverables || []),
    ]).slice(0, 6);

    const disciplineLabels = disciplinesInRow.length
        ? disciplinesInRow
        : [sectionDetails.shortTitle];

    let summary = `${row.description} ligger i ${sectionDetails.shortTitle.toLowerCase()} og bor avklares med tydelig ansvar mellom involverte fag.`;

    if (matchedRules.length) {
        summary = `${row.description} handler typisk om ${matchedRules[0].focus[0].toLowerCase()} og krever at leveranse, kobling og funksjon sees samlet.`;
    }

    return {
        summary,
        disciplines: disciplineLabels,
        focus,
        deliverables,
    };
}

function renderCurrentRowInsight(data) {
    if (currentRowInsightSummary) {
        currentRowInsightSummary.textContent = data.summary;
    }

    renderTagList(currentRowInsightDisciplines, data.disciplines, "Ingen fag valgt");
    renderTagList(currentRowInsightFocus, data.focus, "Ingen fokuspunkt valgt");
    renderTagList(currentRowInsightDeliverables, data.deliverables, "Ingen leveranser valgt");
}

function renderMatrixSectionCards() {
    if (!matrixSectionCards) {
        return;
    }

    matrixSectionCards.innerHTML = "";

    Object.keys(sectionDefinitions)
        .map((key) => Number(key))
        .filter((sectionCode) => sectionCode < 800)
        .forEach((sectionCode) => {
            const details = getSectionDetails(sectionCode);
            const stats = getSectionStats(sectionCode);
            const completionRate = stats.total ? Math.round((stats.confirmed / stats.total) * 100) : 0;
            const isActive = activeSectionFilter === sectionCode;
            const button = document.createElement("button");
            let stateLabel = "Tom kategori";
            let stateClass = "";

            if (stats.total) {
                if (stats.open > 0) {
                    stateLabel = `${stats.open} apne`;
                    stateClass = "state-warning";
                } else {
                    stateLabel = "Klar";
                    stateClass = "state-ok";
                }
            }

            button.type = "button";
            button.className = `matrix-section-card${isActive ? " is-active" : ""}${stats.total ? "" : " is-empty"}`;
            button.setAttribute("aria-pressed", isActive ? "true" : "false");
            button.innerHTML = `
                <div class="matrix-section-card-top">
                    <span class="matrix-section-code">${sectionCode}</span>
                    <span class="matrix-section-state ${stateClass}">${escapeHtml(stateLabel)}</span>
                </div>
                <div>
                    <p class="matrix-section-title">${escapeHtml(details.shortTitle)}</p>
                    <p class="matrix-section-copy">${escapeHtml(details.summary)}</p>
                </div>
                <div class="matrix-section-meta">
                    <span>${stats.total} rader</span>
                    <span>${completionRate} % bekreftet</span>
                </div>
            `;

            button.addEventListener("click", () => {
                setActiveSectionFilter(isActive ? "all" : sectionCode);
            });

            matrixSectionCards.appendChild(button);
        });
}

function getSectionFilterFromHash() {
    const hash = String(window.location.hash || "").replace(/^#/, "").trim().toLowerCase();
    const match = hash.match(/(?:kategori|section)-(\d{3})/);

    if (!match) {
        return "all";
    }

    return Number.parseInt(match[1], 10);
}

function setActiveSectionFilter(nextFilter, options = {}) {
    const { updateHash = true } = options;
    activeSectionFilter = nextFilter === "all" ? "all" : Number(nextFilter);

    if (updateHash) {
        const nextHash = activeSectionFilter === "all" ? "" : `kategori-${activeSectionFilter}`;
        const url = new URL(window.location.href);
        url.hash = nextHash;
        window.history.replaceState(null, "", url);
    }

    // Rebuild matrix with only visible rows for this chapter
    buildMatrixInBatches().then(() => {
        markHeaderGroups();
        filterMatrixRows();
        updateAllRiskCells();
    });
    renderMatrixSectionCards();
}

function focusAdjacentContentRow(direction) {
    const visibleRows = rows
        .map((row, rowIndex) => ({ row, rowIndex, element: getRowElement(rowIndex) }))
        .filter(({ row, element }) => !row.section && element && !element.classList.contains("filtered-out"));

    if (!visibleRows.length) {
        return;
    }

    const currentIndex = visibleRows.findIndex(({ rowIndex }) => rowIndex === activeRowIndex);
    const nextIndex = currentIndex < 0
        ? 0
        : Math.max(0, Math.min(visibleRows.length - 1, currentIndex + direction));

    focusRow(visibleRows[nextIndex].rowIndex);
}

function updateRowMetaPanel() {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        if (currentRowTfm) {
            currentRowTfm.value = "";
            currentRowTfm.disabled = true;
        }
        if (currentRowDescription) {
            currentRowDescription.value = "";
            currentRowDescription.disabled = true;
        }
        if (currentRowRisk) {
            currentRowRisk.textContent = "Ingen rad valgt";
        }
        if (currentRowMissing) {
            currentRowMissing.innerHTML = '<p class="helper-text">Velg en rad for å se hva som mangler.</p>';
        }
        if (currentRowConfirm) {
            currentRowConfirm.checked = false;
            currentRowConfirm.disabled = true;
        }
        if (currentRowConfirmText) {
            currentRowConfirmText.textContent = "Marker valgt rad som bekreftet";
        }
        if (currentRowComment) {
            currentRowComment.value = "";
            currentRowComment.disabled = true;
        }
        if (deleteRowButton) {
            deleteRowButton.disabled = true;
        }
        if (moveRowUpButton) {
            moveRowUpButton.disabled = true;
        }
        if (moveRowDownButton) {
            moveRowDownButton.disabled = true;
        }
        if (quickConfirmRowButton) {
            quickConfirmRowButton.disabled = true;
        }
        if (quickNextUnresolvedButton) {
            quickNextUnresolvedButton.disabled = false;
        }
        if (quickClearCommentButton) {
            quickClearCommentButton.disabled = true;
        }
        renderCurrentRowInsight({
            summary: "Velg en rad for a se hva dette grensesnittet normalt omfatter, hvilke fag som ofte ma med, og hva som bor avklares i leveransen.",
            disciplines: ["Ingen rad valgt"],
            focus: ["Velg en rad"],
            deliverables: ["Velg en rad"],
        });
        return;
    }

    const row = rows[activeRowIndex];
    const risk = getRiskState(activeRowIndex);
    const missingResponsibilities = getMissingResponsibilities(activeRowIndex);
    const insight = getRowInsightData(row, activeRowIndex);

    if (currentRowTfm) {
        currentRowTfm.disabled = false;
        currentRowTfm.value = row.tfm;
    }
    if (currentRowDescription) {
        currentRowDescription.disabled = false;
        currentRowDescription.value = row.description;
    }
    if (currentRowRisk) {
        currentRowRisk.textContent = `${risk.icon} ${risk.title}`;
    }
    if (currentRowMissing) {
        currentRowMissing.innerHTML = missingResponsibilities.length
            ? missingResponsibilities.map((item) => `<span class="missing-pill">${escapeHtml(item)}</span>`).join("")
            : '<p class="helper-text">Alle ansvarskolonner har fått en eier.</p>';
    }
    if (currentRowConfirm) {
        currentRowConfirm.disabled = false;
        currentRowConfirm.checked = Boolean(confirmationState.get(activeRowIndex));
    }
    if (currentRowConfirmText) {
        currentRowConfirmText.textContent = confirmationState.get(activeRowIndex)
            ? "Valgt rad er bekreftet"
            : "Marker valgt rad som bekreftet";
    }
    if (currentRowComment) {
        currentRowComment.disabled = false;
        currentRowComment.value = commentState.get(activeRowIndex) || "";
    }
    if (deleteRowButton) {
        deleteRowButton.disabled = false;
    }
    const contentRows = getContentRows();
    const currentContentIndex = contentRows.findIndex((item) => item.uid === row.uid);
    if (moveRowUpButton) {
        moveRowUpButton.disabled = currentContentIndex <= 0;
    }
    if (moveRowDownButton) {
        moveRowDownButton.disabled = currentContentIndex < 0 || currentContentIndex >= contentRows.length - 1;
    }
    if (quickConfirmRowButton) {
        quickConfirmRowButton.disabled = Boolean(confirmationState.get(activeRowIndex));
    }
    if (quickNextUnresolvedButton) {
        quickNextUnresolvedButton.disabled = false;
    }
    if (quickClearCommentButton) {
        quickClearCommentButton.disabled = !String(commentState.get(activeRowIndex) || "").trim();
    }

    renderCurrentRowInsight(insight);
}

function renderRowDescriptionContent(rowIndex) {
    const wrapper = document.createElement("div");
    wrapper.className = "row-description-stack";

    const title = document.createElement("span");
    title.className = "row-description-title";
    title.textContent = rows[rowIndex].description;
    wrapper.appendChild(title);

    const flags = getRowStatusFlags(rowIndex);
    if (flags.length) {
        const badgeRow = document.createElement("div");
        badgeRow.className = "row-status-badges";
        flags.forEach((flag) => {
            const badge = document.createElement("span");
            badge.className = `row-status-badge ${flag.className}`;
            badge.textContent = flag.label;
            badgeRow.appendChild(badge);
        });
        wrapper.appendChild(badgeRow);
    }

    return wrapper;
}

function refreshMatrixRowVisuals() {
    rows.forEach((row, rowIndex) => {
        if (row.section) {
            return;
        }

        const rowElement = getRowElement(rowIndex);
        if (!rowElement) {
            return;
        }

        const descriptionCell = rowElement.querySelector(".description-cell");
        if (descriptionCell) {
            descriptionCell.innerHTML = "";
            descriptionCell.appendChild(renderRowDescriptionContent(rowIndex));
        }

        const flags = getRowStatusFlags(rowIndex);
        rowElement.classList.toggle("row-has-conflict", flags.some((flag) => flag.className === "status-conflict"));
        rowElement.classList.toggle("row-is-confirmed", flags.some((flag) => flag.className === "status-confirmed"));
        rowElement.classList.toggle("row-is-open", flags.some((flag) => flag.className === "status-open"));
    });
}

function updateAllRiskCells() {
    updateRowMetaPanel();
    updateMatrixOverview();
    refreshMatrixRowVisuals();
}

function updateRowAfterMatrixEdit(rowIndex) {
    activeRowIndex = rowIndex;
    updateAllRiskCells();
    buildContractSummary();
    scheduleAutosave();
}

function moveMatrixButtonFocus(button, rowStep, responsibilityStep) {
    const rowIndex = Number(button.dataset.row);
    const responsibilityIndex = responsibilities.indexOf(button.dataset.responsibility);
    const discipline = button.dataset.discipline;

    if (rowIndex < 0 || responsibilityIndex < 0 || !discipline) {
        return;
    }

    const visibleContentRows = rows
        .map((row, index) => ({ row, index, element: getRowElement(index) }))
        .filter(({ row, element }) => !row.section && element && !element.classList.contains("filtered-out"));
    const visibleRowPosition = visibleContentRows.findIndex(({ index }) => index === rowIndex);

    if (visibleRowPosition < 0) {
        return;
    }

    const nextVisibleRow = visibleContentRows[Math.max(0, Math.min(visibleContentRows.length - 1, visibleRowPosition + rowStep))];
    const nextResponsibility = responsibilities[Math.max(0, Math.min(responsibilities.length - 1, responsibilityIndex + responsibilityStep))];
    const nextButton = matrixBody.querySelector(
        `button[data-row="${nextVisibleRow.index}"][data-discipline="${discipline}"][data-responsibility="${nextResponsibility}"]`
    );

    if (nextButton) {
        nextButton.focus();
        focusRow(nextVisibleRow.index);
    }
}

function setResponsibilityValue(rowIndex, discipline, responsibility, state) {
    const activityButtons = matrixBody.querySelectorAll(
        `button[data-row="${rowIndex}"][data-responsibility="${responsibility}"]`
    );
    const targetButton = matrixBody.querySelector(
        `button[data-row="${rowIndex}"][data-discipline="${discipline}"][data-responsibility="${responsibility}"]`
    );

    activityButtons.forEach((activityButton) => {
        applyState(activityButton, "");
    });

    if (targetButton && state) {
        applyState(targetButton, state);
    }

    updateRowAfterMatrixEdit(rowIndex);
}

function getRowElement(rowIndex) {
    return matrixBody.querySelector(`tr[data-row-index="${rowIndex}"]`);
}

function updateRowDisplay(rowIndex) {
    const rowElement = getRowElement(rowIndex);
    if (!rowElement) {
        return;
    }

    const tfmCell = rowElement.querySelector(".tfm-cell");
    const descriptionCell = rowElement.querySelector(".description-cell");

    if (tfmCell) {
        tfmCell.textContent = rows[rowIndex].tfm;
    }

    if (descriptionCell) {
        if (rows[rowIndex].section) {
            const label = descriptionCell.querySelector("span");
            const toggle = descriptionCell.querySelector(".section-toggle");
            if (label) {
                label.textContent = getSectionLabel(rowIndex);
            }
            if (toggle) {
                toggle.textContent = isSectionCollapsed(rowIndex) ? "+" : "-";
                toggle.setAttribute("aria-label", `${isSectionCollapsed(rowIndex) ? "Utvid" : "Skjul"} seksjon ${rows[rowIndex].description}`);
            }
        } else {
            descriptionCell.innerHTML = "";
            descriptionCell.appendChild(renderRowDescriptionContent(rowIndex));
        }
    }
}

function findSelectedButton(rowIndex, responsibility) {
    const rowButtons = matrixBody.querySelectorAll(
        `button[data-row="${rowIndex}"][data-responsibility="${responsibility}"]`
    );

    return Array.from(rowButtons).find((button) => button.dataset.state !== "");
}

function setCellState(rowIndex, discipline, responsibility, state) {
    const button = matrixBody.querySelector(
        `button[data-row="${rowIndex}"][data-discipline="${discipline}"][data-responsibility="${responsibility}"]`
    );

    if (button) {
        applyState(button, state);
    }
}

function clearResponsibility(rowIndex, responsibility) {
    const buttons = matrixBody.querySelectorAll(
        `button[data-row="${rowIndex}"][data-responsibility="${responsibility}"]`
    );
    buttons.forEach((button) => applyState(button, ""));
}

function setConfirmation(rowIndex, isConfirmed) {
    confirmationState.set(rowIndex, Boolean(isConfirmed));
    if (rowIndex === activeRowIndex) {
        updateRowMetaPanel();
    }
}

function resetMatrixToBaseMarks() {
    rows.forEach((row, rowIndex) => {
        disciplines.forEach((discipline) => {
            responsibilities.forEach((responsibility) => {
                const state = baseMarksByRow[rowIndex][`${discipline}:${responsibility}`] || "";
                setCellState(rowIndex, discipline, responsibility, state);
            });
        });
    });
}

function shouldOverrideResponsibility(rowIndex, responsibility) {
    const selectedButton = findSelectedButton(rowIndex, responsibility);
    if (!selectedButton) {
        return true;
    }

    return packageControlledDisciplines.includes(selectedButton.dataset.discipline);
}

function getPresetOwners(selectedPackages) {
    const owners = {};

    if (selectedPackages.includes("totaltechnical")) {
        owners.K = "EL";
        owners.F = "Aut";
        owners.I = "SD";
        owners.M = "Lås og beslag";
        return owners;
    }

    if (selectedPackages.includes("el_aut_sd")) {
        owners.K = "EL";
        owners.F = "Aut";
        owners.I = "SD";
    } else {
        if (selectedPackages.includes("el") || selectedPackages.includes("el_aut")) {
            owners.K = "EL";
        }

        if (selectedPackages.includes("aut") || selectedPackages.includes("el_aut")) {
            owners.F = "Aut";
        }

        if (selectedPackages.includes("sd")) {
            owners.I = "SD";
        } else if (selectedPackages.includes("aut") || selectedPackages.includes("el_aut")) {
            owners.I = "Aut";
        }
    }

    if (selectedPackages.includes("las")) {
        owners.M = "Lås og beslag";
    }

    return owners;
}

function applyPackagePreset() {
    const selectedPackages = getSelectedPackages();
    resetMatrixToBaseMarks();

    if (!selectedPackages.length) {
        updateAllRiskCells();
        buildContractSummary();
        return;
    }

    const presetOwners = getPresetOwners(selectedPackages);

    rows.forEach((row, rowIndex) => {
        if (row.section) {
            return;
        }

        Object.entries(presetOwners).forEach(([responsibility, ownerDiscipline]) => {
            if (!shouldOverrideResponsibility(rowIndex, responsibility)) {
                return;
            }

            clearResponsibility(rowIndex, responsibility);
            setCellState(rowIndex, ownerDiscipline, responsibility, "H");
        });
    });

    updateAllRiskCells();
    buildContractSummary();
}

function applyBhSuggestionsFromText(inputText) {
    const text = inputText.toLowerCase();
    const findings = [];
    const packageSignals = [];
    const integrationSignals = [];
    let suggestedProjectType = projectTypeSelect.value || "bolig";
    let suggestedCoreModel = tueCoreModelSelect?.value || "separate";
    const suggestedStandalone = new Set();

    const markStandalone = (value) => {
        const input = packageOptionInputs.find((option) => option.value === value);
        if (input) {
            input.checked = true;
        }
        suggestedStandalone.add(value);
    };

    if (text.includes("totalteknisk")) {
        if (tueCoreModelSelect) {
            tueCoreModelSelect.value = "totaltechnical";
        }
        suggestedCoreModel = "totaltechnical";
        packageSignals.push("Totalteknisk leveranse");
    } else if (text.includes("sd") && text.includes("aut") && text.includes("el")) {
        if (tueCoreModelSelect) {
            tueCoreModelSelect.value = "el_aut_sd";
        }
        suggestedCoreModel = "el_aut_sd";
        packageSignals.push("EL + AUT + SD");
    } else if (text.includes("automatikk") || text.includes("frekvensomformer")) {
        if (tueCoreModelSelect) {
            tueCoreModelSelect.value = "separate";
        }
        suggestedCoreModel = "separate";
        packageSignals.push("Separate tekniske UE-er");
    }

    if ((tueCoreModelSelect?.value || "separate") === "separate") {
        if (text.includes("bacnet") || text.includes("modbus") || text.includes("sd-anlegg") || text.includes("integrasjon")) {
            markStandalone("sd");
            integrationSignals.push("SD/BAS-integrasjon");
        }

        if (text.includes("frekvensomformer") || text.includes("automatikk") || text.includes("bus")) {
            markStandalone("aut");
            packageSignals.push("AUT som egen UE");
        }

        if (text.includes("kabel") || text.includes("elkraft") || text.includes("strøm")) {
            markStandalone("el");
            packageSignals.push("EL som egen UE");
        }
    }

    if (tueLocksModelSelect) {
        tueLocksModelSelect.value = text.includes("lås") || text.includes("beslag") ? "separate" : tueLocksModelSelect.value;
    }

    if (text.includes("lås") || text.includes("beslag")) {
        findings.push("Lås og beslag er nevnt og bør vurderes som egen avklaring.");
    }

    if (text.includes("sykehus")) {
        projectTypeSelect.value = "sykehus";
        suggestedProjectType = "sykehus";
    } else if (text.includes("skole")) {
        projectTypeSelect.value = "skole";
        suggestedProjectType = "skole";
    } else if (text.includes("barnehage")) {
        projectTypeSelect.value = "barnehage";
        suggestedProjectType = "barnehage";
    } else if (text.includes("hotell")) {
        projectTypeSelect.value = "hotell";
        suggestedProjectType = "hotell";
    } else if (text.includes("logistikk") || text.includes("lager")) {
        projectTypeSelect.value = "logistikk";
        suggestedProjectType = "logistikk";
    } else if (text.includes("datahall")) {
        projectTypeSelect.value = "datahall";
        suggestedProjectType = "datahall";
    } else if (text.includes("kontor")) {
        projectTypeSelect.value = "kontor";
        suggestedProjectType = "kontor";
    } else if (text.includes("industri")) {
        projectTypeSelect.value = "industri";
        suggestedProjectType = "industri";
    } else if (text.includes("rehab") || text.includes("ombygg")) {
        projectTypeSelect.value = "rehab";
        suggestedProjectType = "rehab";
    } else if (text.includes("bolig")) {
        projectTypeSelect.value = "bolig";
        suggestedProjectType = "bolig";
    }

    if (text.includes("bacnet")) {
        integrationSignals.push("BACnet");
    }

    if (text.includes("modbus")) {
        integrationSignals.push("Modbus");
    }

    if (text.includes("adgangskontroll") || text.includes("adk")) {
        findings.push("Adgangskontroll er nevnt og krever tydelig ansvar mellom EL og lås.");
    }

    if (text.includes("frekvensomformer")) {
        findings.push("Frekvensomformer er nevnt og peker mot behov for tydelig EL/AUT-grensesnitt.");
    }

    if (text.includes("sd-anlegg") || text.includes("sd")) {
        findings.push("Underlaget peker mot integrasjon mot SD-anlegg.");
    }

    syncTueBuilderUI();

    const summary = {
        projectType: suggestedProjectType,
        coreModel: suggestedCoreModel,
        standalone: Array.from(suggestedStandalone),
        findings: Array.from(new Set(findings)),
        packageSignals: Array.from(new Set(packageSignals)),
        integrationSignals: Array.from(new Set(integrationSignals)),
        keywordScore: findings.length + packageSignals.length + integrationSignals.length,
    };

    lastBhAnalysis = summary;
    renderBhAnalysisInsights(summary);
    return summary;
}

function setPersistenceMessage(message, isError = false) {
    if (!persistenceStatus) {
        return;
    }

    persistenceStatus.textContent = message;
    persistenceStatus.style.color = isError ? "#ab2220" : "";
}

function setAutosaveMessage(message, isError = false) {
    if (!autosaveStatus) {
        return;
    }

    autosaveStatus.textContent = message;
    autosaveStatus.style.color = isError ? "#ab2220" : "";
}

function getCurrentProjectId() {
    return (projectIdInput?.value || "").trim() || "default";
}

function rememberLastProject(projectId) {
    try {
        window.localStorage.setItem(LAST_PROJECT_KEY, projectId);
    } catch (_error) {
        // Ignore local storage issues and continue without persistence here.
    }
}

function loadExcelRowsData() {
    if (matrixDataPromise) {
        return matrixDataPromise;
    }

    matrixDataPromise = fetch("excel-data.json")
        .then((response) => {
            if (!response.ok) {
                throw new Error(`Kunne ikke hente matrisedata (HTTP ${response.status}).`);
            }

            return response.json();
        })
        .then((payload) => (Array.isArray(payload) && payload.length ? payload : defaultRows))
        .catch(() => defaultRows);

    return matrixDataPromise;
}

function initializeRows(rowSource) {
    rows = normalizeRowsByTfm(Array.isArray(rowSource) && rowSource.length ? rowSource : defaultRows);
    baseMarksByRow = rows.map((row) => ({ ...row.marks }));
}

function getRememberedProject() {
    try {
        return window.localStorage.getItem(LAST_PROJECT_KEY) || "";
    } catch (_error) {
        return "";
    }
}

function getSavedReviewMode() {
    try {
        return window.localStorage.getItem(REVIEW_MODE_KEY) === "true";
    } catch (_error) {
        return false;
    }
}

function getSavedReviewFilter() {
    try {
        const savedFilter = window.localStorage.getItem(REVIEW_FILTER_KEY) || "all";
        return ["all", "open", "conflicts", "confirmed"].includes(savedFilter) ? savedFilter : "all";
    } catch (_error) {
        return "all";
    }
}

function getReviewFilterLabel(filter = activeReviewFilter) {
    return {
        all: "Alle rader",
        open: "Åpne rader",
        conflicts: "Konflikter",
        confirmed: "Bekreftede",
    }[filter] || "Alle rader";
}

function getOfferConflictRowIds() {
    return new Set(
        (lastOfferAnalysis?.findings || [])
            .map((finding) => finding.rowUid)
            .filter(Boolean)
    );
}

function getRowStatusFlags(rowIndex) {
    if (rows[rowIndex]?.section) {
        return [];
    }

    const flags = [];
    const risk = getRiskState(rowIndex);
    const hasComment = Boolean(String(commentState.get(rowIndex) ?? rows[rowIndex]?.comments ?? "").trim());
    const isConfirmed = Boolean(confirmationState.get(rowIndex));
    const hasConflict = getOfferConflictRowIds().has(rows[rowIndex].uid);

    if (hasConflict) {
        flags.push({ label: "Konflikt", className: "status-conflict" });
    }
    if (risk.level !== "ok") {
        flags.push({ label: "Åpen", className: "status-open" });
    }
    if (isConfirmed) {
        flags.push({ label: "Bekreftet", className: "status-confirmed" });
    }
    if (hasComment) {
        flags.push({ label: "Kommentar", className: "status-comment" });
    }

    return flags;
}

function getRowsNeedingComment() {
    return rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(({ row, rowIndex }) => {
            if (row.section) {
                return false;
            }
            if (getRiskState(rowIndex).level === "ok") {
                return false;
            }
            return !String(commentState.get(rowIndex) ?? row.comments ?? "").trim();
        });
}

function getReviewReadyCount() {
    return rows.filter((row, rowIndex) => !row.section && confirmationState.get(rowIndex) && getRiskState(rowIndex).level === "ok").length;
}

function rowMatchesReviewFilter(row, rowIndex) {
    switch (activeReviewFilter) {
        case "open":
            return getRiskState(rowIndex).level !== "ok";
        case "confirmed":
            return Boolean(confirmationState.get(rowIndex));
        case "conflicts":
            return getOfferConflictRowIds().has(row.uid);
        case "all":
        default:
            return true;
    }
}

function updateReviewFilterButtons() {
    reviewFilterButtons.forEach((button) => {
        const isActive = button.dataset.reviewFilter === activeReviewFilter;
        button.setAttribute("aria-pressed", isActive ? "true" : "false");
        button.classList.toggle("is-active", isActive);
    });
}

function applyReviewFilter(nextFilter, options = {}) {
    const normalizedFilter = ["all", "open", "conflicts", "confirmed"].includes(nextFilter) ? nextFilter : "all";
    activeReviewFilter = normalizedFilter;
    updateReviewFilterButtons();

    try {
        window.localStorage.setItem(REVIEW_FILTER_KEY, activeReviewFilter);
    } catch (_error) {
        // Ignore storage issues.
    }

    if (!options.skipRefilter) {
        filterMatrixRows();
    }
}

function applyReviewMode(enabled) {
    reviewModeEnabled = Boolean(enabled);
    document.body.classList.toggle("review-mode", reviewModeEnabled);

    if (toggleReviewModeButton) {
        toggleReviewModeButton.setAttribute("aria-pressed", reviewModeEnabled ? "true" : "false");
        toggleReviewModeButton.textContent = reviewModeEnabled ? "Vanlig visning" : "Review mode";
    }

    try {
        window.localStorage.setItem(REVIEW_MODE_KEY, reviewModeEnabled ? "true" : "false");
    } catch (_error) {
        // Ignore storage issues.
    }

    updateReviewFilterButtons();
    filterMatrixRows();
}

function formatRelativeTime(dateValue) {
    if (!dateValue) {
        return "ukjent tidspunkt";
    }

    const timestamp = new Date(dateValue).getTime();

    if (Number.isNaN(timestamp)) {
        return "ukjent tidspunkt";
    }

    const diffMinutes = Math.round((Date.now() - timestamp) / 60000);

    if (diffMinutes <= 1) {
        return "akkurat nå";
    }

    if (diffMinutes < 60) {
        return `${diffMinutes} min siden`;
    }

    const diffHours = Math.round(diffMinutes / 60);

    if (diffHours < 24) {
        return `${diffHours} t siden`;
    }

    const diffDays = Math.round(diffHours / 24);
    return `${diffDays} d siden`;
}

function getWorkflowStepMeta(stepNumber) {
    return {
        1: { title: "Prosjektoppsett", description: "Steg 1 av 4: Prosjektoppsett" },
        2: { title: "BH-underlag", description: "Steg 2 av 4: BH-underlag" },
        3: { title: "Matrise", description: "Steg 3 av 4: Matrise" },
        4: { title: "Tilbudskontroll", description: "Steg 4 av 4: Tilbudskontroll" },
    }[stepNumber] || { title: "Arbeidsflyt", description: "Steg i arbeidsflyten" };
}

function createChecklistMarkup(items) {
    return items.map((item) => `
        <div class="checklist-item${item.done ? " is-done" : ""}">
            <span class="checkmark">${item.done ? "OK" : "!"}</span>
            <span><strong>${escapeHtml(item.label)}</strong><br>${escapeHtml(item.detail)}</span>
        </div>
    `).join("");
}

function getWorkflowHealth() {
    const hasProjectId = Boolean(getCurrentProjectId() && getCurrentProjectId() !== "default");
    const hasProjectType = Boolean(projectTypeSelect?.value);
    const hasTueSummary = Boolean((tueCompositionSummary?.textContent || "").trim());
    const hasDocuments = typeof uploadedDocuments !== "undefined" && uploadedDocuments.length > 0;
    const hasComplexityAnalysis = typeof lastComplexityResult !== "undefined" && lastComplexityResult !== null;
    const contentRows = getContentRowCount();
    const confirmedRows = getConfirmedRowCount();
    const openRows = getOpenRiskCount();
    const hasMatrixProgress = confirmedRows > 0;
    const hasOfferDocuments = uploadedOfferDocuments.length > 0;
    const hasOfferAnalysis = Boolean(lastOfferAnalysis);
    const offerConflicts = lastOfferAnalysis?.conflictCount || 0;

    const step1Checks = [
        {
            label: "Prosjektnavn gitt",
            detail: hasProjectId ? `Aktivt prosjekt: ${getCurrentProjectId()}` : "Gi prosjektet et navn slik at du kan finne det igjen.",
            done: hasProjectId,
        },
        {
            label: "Bygningstype valgt",
            detail: `Valgt type: ${getProjectTypeLabel(projectTypeSelect?.value || "bolig")}`,
            done: hasProjectType,
        },
        {
            label: "Teknisk pakkestruktur satt",
            detail: hasTueSummary ? tueCompositionSummary.textContent.trim() : "Velg hvordan de tekniske fagene skal organiseres.",
            done: hasTueSummary,
        },
    ];

    const docCountText = hasDocuments ? `${uploadedDocuments.length} dokument(er) lastet opp.` : "";
    const step2Checks = [
        {
            label: "Dokumenter lastet opp",
            detail: hasDocuments ? docCountText : "Last opp anbudsdokumentene fra byggherre.",
            done: hasDocuments,
        },
        {
            label: "Analyse kjørt",
            detail: hasComplexityAnalysis
                ? `Kompleksitet: ${lastComplexityResult.levelLabel} (${lastComplexityResult.score}/100)`
                : "Kjør analyse for å få forslag til oppsett og matriseomfang.",
            done: hasComplexityAnalysis,
        },
        {
            label: "Anbefalinger vurdert",
            detail: hasComplexityAnalysis
                ? "Se over TUE-anbefaling og matriseomfang, og bruk dem hvis de passer."
                : "Ingen anbefalinger tilgjengelig ennå.",
            done: hasComplexityAnalysis,
        },
    ];

    const completionRate = contentRows ? Math.round((confirmedRows / contentRows) * 100) : 0;
    const step3Checks = [
        {
            label: "Matrisearbeid startet",
            detail: hasMatrixProgress ? `${confirmedRows} rader er bekreftet.` : "Marker rader og jobb deg gjennom åpne avklaringer.",
            done: hasMatrixProgress,
        },
        {
            label: "Åpne avklaringer redusert",
            detail: `${openRows} åpne avklaringer gjenstår.`,
            done: openRows === 0 && contentRows > 0,
        },
        {
            label: "Eksportklar status",
            detail: contentRows ? `${completionRate} % av radene er bekreftet.` : "Ingen matriserader tilgjengelig.",
            done: contentRows > 0 && completionRate === 100 && openRows === 0,
        },
    ];

    const step4Checks = [
        {
            label: "Tilbud lastet opp",
            detail: hasOfferDocuments
                ? `${uploadedOfferDocuments.length} dokument(er) lastet opp.`
                : "Last opp ett eller flere UE-/TUE-tilbud for kontroll.",
            done: hasOfferDocuments,
        },
        {
            label: "Tilbud analysert",
            detail: hasOfferAnalysis
                ? `${lastOfferAnalysis.findings.length} funn registrert i kontrollen.`
                : "Kjør analyse mot matrisen for å se avvik og forbehold.",
            done: hasOfferAnalysis,
        },
        {
            label: "Konflikter vurdert",
            detail: hasOfferAnalysis
                ? (offerConflicts ? `${offerConflicts} konflikt(er) krever vurdering.` : "Ingen tydelige konflikter funnet i første kontroll.")
                : "Ingen vurdering gjort ennå.",
            done: hasOfferAnalysis && offerConflicts === 0,
        },
    ];

    return {
        step1Checks,
        step2Checks,
        step3Checks,
        step4Checks,
        completionRate,
    };
}

function updateWorkflowOverview() {
    const health = getWorkflowHealth();
    const steps = [
        { checks: health.step1Checks, stateNode: step1State, hintNode: step1Hint, title: "Prosjekt", fallback: "Fyll inn prosjekt og TUE" },
        { checks: health.step2Checks, stateNode: step2State, hintNode: step2Hint, title: "BH-underlag", fallback: "Importer BH-underlag" },
        { checks: health.step3Checks, stateNode: step3State, hintNode: step3Hint, title: "Matrise", fallback: "Bearbeid matrise og eksporter" },
        { checks: health.step4Checks, stateNode: step4State, hintNode: step4Hint, title: "Tilbud", fallback: "Kontroller mottatte tilbud" },
    ];
    const completedSteps = steps.filter(({ checks }) => checks.every((item) => item.done)).length;
    const progressPercent = Math.round((completedSteps / steps.length) * 100);
    const recommendedStep = getRecommendedWorkflowStep();
    const recommendedMeta = getWorkflowStepMeta(recommendedStep);

    if (workflowProgressValue) {
        workflowProgressValue.textContent = `${progressPercent} %`;
    }

    if (workflowProgressText) {
        workflowProgressText.textContent = completedSteps === steps.length
            ? "Alle steg er klare for eksport"
            : `Anbefalt neste fokus: ${recommendedMeta.title}`;
    }

    steps.forEach(({ checks, stateNode, hintNode, fallback }) => {
        const doneCount = checks.filter((item) => item.done).length;
        const totalCount = checks.length;
        const isComplete = doneCount === totalCount;
        const isStarted = doneCount > 0;

        if (stateNode) {
            stateNode.textContent = isComplete ? "Ferdig" : isStarted ? "Pågår" : "Venter";
        }

        if (hintNode) {
            hintNode.textContent = isComplete
                ? `${doneCount}/${totalCount} punkter fullført`
                : isStarted
                    ? `${doneCount}/${totalCount} punkter fullført`
                    : fallback;
        }
    });

    if (step1Checklist) {
        step1Checklist.innerHTML = createChecklistMarkup(health.step1Checks);
    }

    if (step2Checklist) {
        step2Checklist.innerHTML = createChecklistMarkup(health.step2Checks);
    }

    if (step3Checklist) {
        step3Checklist.innerHTML = createChecklistMarkup(health.step3Checks);
    }

    if (step4Checklist) {
        step4Checklist.innerHTML = createChecklistMarkup(health.step4Checks);
    }

    updateProductCockpit(health, progressPercent, recommendedMeta);
}

function updateProductCockpit(health, progressPercent, recommendedMeta) {
    const openRows = getOpenRiskCount();
    const commentGaps = getRowsNeedingComment().length;
    const offerConflicts = lastOfferAnalysis?.conflictCount || 0;
    const hasOfferAnalysis = Boolean(lastOfferAnalysis);

    if (cockpitProgressValue) {
        cockpitProgressValue.textContent = `${progressPercent} %`;
    }
    if (cockpitProgressText) {
        cockpitProgressText.textContent = progressPercent === 100
            ? "Alle hovedsteg er ferdig gjennomfort."
            : `${health.completionRate || 0} % av matriseradene er bekreftet.`;
    }
    if (cockpitNextStep) {
        cockpitNextStep.textContent = recommendedMeta.title;
    }
    if (cockpitNextStepDetail) {
        cockpitNextStepDetail.textContent = recommendedMeta.description;
    }
    if (cockpitMatrixHealth) {
        cockpitMatrixHealth.textContent = openRows === 0 ? "Kontrollert" : `${openRows} apne`;
    }
    if (cockpitMatrixHealthDetail) {
        cockpitMatrixHealthDetail.textContent = commentGaps
            ? `${commentGaps} apne rad(er) mangler kommentar eller vurdering.`
            : "Kommentarer og avklaringer ser ryddige ut i arbeidsflaten.";
    }
    if (cockpitOfferHealth) {
        cockpitOfferHealth.textContent = hasOfferAnalysis
            ? (offerConflicts ? `${offerConflicts} konflikter` : "Ingen konflikter")
            : "Ikke startet";
    }
    if (cockpitOfferHealthDetail) {
        cockpitOfferHealthDetail.textContent = hasOfferAnalysis
            ? `${lastOfferAnalysis.findingCount} funn er registrert i tilbudskontrollen.`
            : "Tilbudslaget blir synlig her nar UE-/TUE-tilbud er lastet opp.";
    }
}

function getRecommendedWorkflowStep() {
    const hasProjectId = Boolean(getCurrentProjectId());
    const hasDocContent = (typeof uploadedDocuments !== "undefined" && uploadedDocuments.length > 0)
        || Boolean(`${uploadedBhText}`.trim());
    const hasAnalysis = (typeof lastComplexityResult !== "undefined" && lastComplexityResult !== null) || lastBhAnalysis;
    const hasMatrixWork = getConfirmedRowCount() > 0;
    const hasOfferDocuments = uploadedOfferDocuments.length > 0;
    const hasOfferAnalysis = Boolean(lastOfferAnalysis);

    if (!hasProjectId) {
        return 1;
    }

    if (!hasDocContent && !hasAnalysis) {
        return 2;
    }

    if (hasOfferDocuments && !hasOfferAnalysis) {
        return 4;
    }

    if (hasOfferAnalysis) {
        return 4;
    }

    if (hasMatrixWork) {
        return 3;
    }

    return hasDocContent ? 3 : 2;
}

function setWorkflowStep(stepNumber, options = {}) {
    const nextStep = Math.max(1, Math.min(4, Number(stepNumber) || 1));
    currentWorkflowStep = nextStep;

    if (nextStep === 3) {
        void ensureMatrixInitialized({ focusFirstRow: true });
    }

    workflowPanels.forEach((panel) => {
        const isActive = Number(panel.dataset.stepPanel) === nextStep;
        panel.hidden = !isActive;
        panel.classList.toggle("active", isActive);
    });

    workflowTabs.forEach((tab) => {
        const isActive = Number(tab.dataset.stepTarget) === nextStep;
        tab.classList.toggle("active", isActive);
        tab.setAttribute("aria-selected", isActive ? "true" : "false");
        tab.tabIndex = isActive ? 0 : -1;
    });

    if (workflowStepStatus) {
        workflowStepStatus.textContent = getWorkflowStepMeta(nextStep).description;
    }

    updateWorkflowOverview();

    if (options.scroll !== false) {
        document.getElementById("top")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
}

async function ensureMatrixInitialized(options = {}) {
    if (matrixInitialized) {
        if (options.focusFirstRow && focusedRowIndex < 0) {
            const firstContentRowIndex = rows.findIndex((row) => !row.section);
            if (firstContentRowIndex >= 0) {
                focusRow(firstContentRowIndex);
            }
        }
        return;
    }

    if (matrixBuildInProgress) {
        return;
    }

    if (!usingImportedBaseRows && !hasProjectSpecificRows) {
        if (workflowProgressText) {
            workflowProgressText.textContent = "Laster matrisedata for steg 3...";
        }

        const loadedRows = await loadExcelRowsData();
        initializeRows(loadedRows);
        usingImportedBaseRows = true;
    }

    if (workflowProgressText) {
        workflowProgressText.textContent = "Bygger matrise...";
    }

    await buildMatrixInBatches();
    markHeaderGroups();
    matrixInitialized = true;
    updateAllRiskCells();
    filterMatrixRows();
    buildContractSummary();
    updateWorkflowOverview();

    if (options.focusFirstRow) {
        const firstContentRowIndex = rows.findIndex((row) => !row.section);
        if (firstContentRowIndex >= 0) {
            focusRow(firstContentRowIndex);
        }
    }
}

function scheduleMatrixInitialization() {
    if (currentWorkflowStep === 3) {
        void ensureMatrixInitialized();
    }
}

function renderProjectLibraryStats(projects = cachedProjects, revisions = cachedRevisions) {
    if (!projectLibraryStats) {
        return;
    }

    const activeProjectId = getCurrentProjectId();
    const totalProjects = projects.length;
    const activeProject = projects.find((project) => project.id === activeProjectId);
    const latestProject = [...projects]
        .sort((left, right) => new Date(right.updatedAt || 0).getTime() - new Date(left.updatedAt || 0).getTime())[0];

    projectLibraryStats.innerHTML = `
        <div class="library-stat">
            <strong>${totalProjects}</strong>
            <span>lagrede prosjekter</span>
        </div>
        <div class="library-stat">
            <strong>${revisions.length}</strong>
            <span>versjoner for aktivt prosjekt</span>
        </div>
        <div class="library-stat">
            <strong>${escapeHtml(activeProject?.id || activeProjectId)}</strong>
            <span>aktivt prosjekt</span>
        </div>
        <div class="library-stat">
            <strong>${escapeHtml(latestProject ? formatRelativeTime(latestProject.updatedAt) : "ingen data")}</strong>
            <span>siste oppdatering</span>
        </div>
    `;
    updateWorkflowOverview();
}

function renderBhAnalysisInsights(analysis = lastBhAnalysis) {
    if (!bhAnalysisInsights) {
        return;
    }

    if (!analysis || (!analysis.findings.length && !analysis.packageSignals.length && !analysis.integrationSignals.length)) {
        bhAnalysisInsights.innerHTML = `
            <div class="analysis-card">
                <p class="status-label">Analyse</p>
                <p class="status-value">Ingen tydelige signaler funnet ennå. Legg inn mer tekst eller last opp et mer detaljert underlag.</p>
            </div>
        `;
        updateWorkflowOverview();
        return;
    }

    const projectTypeLabel = getProjectTypeLabel(analysis.projectType);
    const packageText = analysis.packageSignals.length ? analysis.packageSignals.join(", ") : "Ingen tydelig pakkeindikasjon";
    const integrationText = analysis.integrationSignals.length ? analysis.integrationSignals.join(", ") : "Ingen tydelige integrasjonssignaler";
    const findingsMarkup = analysis.findings.length
        ? `<ul>${analysis.findings.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
        : "<p>Ingen spesifikke risikofunn registrert.</p>";

    bhAnalysisInsights.innerHTML = `
        <div class="analysis-grid">
            <article class="analysis-card">
                <p class="status-label">Foreslått prosjekt</p>
                <p class="status-value"><strong>${escapeHtml(projectTypeLabel)}</strong></p>
                <p class="helper-text">Basert på identifiserte ord og faglige signaler i underlaget.</p>
            </article>
            <article class="analysis-card">
                <p class="status-label">Foreslått TUE</p>
                <p class="status-value">${escapeHtml(packageText)}</p>
                <p class="helper-text">Separate fag markeres automatisk når teksten peker tydelig i den retningen.</p>
            </article>
            <article class="analysis-card">
                <p class="status-label">Integrasjonssignaler</p>
                <p class="status-value">${escapeHtml(integrationText)}</p>
                <p class="helper-text">Typiske indikatorer er SD, BACnet, Modbus, bus og frekvensomformer.</p>
            </article>
        </div>
        <article class="analysis-card analysis-findings">
            <p class="status-label">Funn som bør følges opp</p>
            ${findingsMarkup}
        </article>
    `;
    updateWorkflowOverview();
}

async function readResponsePayload(response) {
    const rawText = await response.text();

    if (!rawText) {
        return null;
    }

    try {
        return JSON.parse(rawText);
    } catch (_error) {
        return { error: rawText };
    }
}

function collectMatrixMarks() {
    const matrixMarks = {};

    rows.forEach((_, rowIndex) => {
        disciplines.forEach((discipline) => {
            responsibilities.forEach((responsibility) => {
                const button = matrixBody.querySelector(
                    `button[data-row="${rowIndex}"][data-discipline="${discipline}"][data-responsibility="${responsibility}"]`
                );

                if (button?.dataset.state) {
                    matrixMarks[`${rowIndex}:${discipline}:${responsibility}`] = button.dataset.state;
                }
            });
        });
    });

    return matrixMarks;
}

function collectComments() {
    return Object.fromEntries(commentState.entries());
}

function collectRowDefinitions() {
    return rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(({ row }) => !row.section)
        .map(({ row, rowIndex }) => ({
            uid: row.uid,
            tfm: row.tfm,
            description: row.description,
            comments: commentState.get(rowIndex) ?? row.comments ?? "",
            section: false,
            marks: { ...baseMarksByRow[rowIndex] },
        }));
}

function collectProjectState() {
    return {
        projectId: getCurrentProjectId(),
        projectType: projectTypeSelect.value,
        breeamLevel: getBreeamLevel(),
        tueConfig: getTueConfig(),
        selectedPackages: getSelectedPackages(),
        uploadedBhText,
        offerDocuments: uploadedOfferDocuments,
        offerAnalysis: lastOfferAnalysis,
        rowDefinitions: collectRowDefinitions(),
        matrixMarks: collectMatrixMarks(),
        comments: collectComments(),
        confirmations: Object.fromEntries(confirmationState.entries()),
        savedAt: new Date().toISOString(),
    };
}

function applySavedTueConfig(tueConfig, selectedPackages = []) {
    const fallbackConfig = {
        coreModel: selectedPackages.includes("totaltechnical")
            ? "totaltechnical"
            : selectedPackages.includes("el_aut_sd")
                ? "el_aut_sd"
                : selectedPackages.includes("el_aut")
                    ? "el_aut"
                    : "separate",
        locksModel: selectedPackages.includes("las") ? "separate" : "integrated",
        adkModel: "el",
        standaloneDisciplines: selectedPackages.filter((key) => ["el", "aut", "sd"].includes(key)),
    };
    const nextConfig = tueConfig || fallbackConfig;

    if (tueCoreModelSelect) {
        tueCoreModelSelect.value = nextConfig.coreModel || "separate";
    }

    if (tueLocksModelSelect) {
        tueLocksModelSelect.value = nextConfig.locksModel || "separate";
    }
    if (tueAdkModelSelect) {
        tueAdkModelSelect.value = nextConfig.adkModel || "el";
    }

    packageOptionInputs.forEach((input) => {
        input.checked = (nextConfig.standaloneDisciplines || []).includes(input.value);
    });

    syncTueBuilderUI();
}

function applySavedMatrix(matrixMarks = {}) {
    rows.forEach((_, rowIndex) => {
        disciplines.forEach((discipline) => {
            responsibilities.forEach((responsibility) => {
                const state = matrixMarks[`${rowIndex}:${discipline}:${responsibility}`] || "";
                setCellState(rowIndex, discipline, responsibility, state);
            });
        });
    });
}

function applySavedConfirmations(confirmations = {}) {
    rows.forEach((_, rowIndex) => {
        setConfirmation(rowIndex, Boolean(confirmations[rowIndex]));
    });
}

function applySavedComments(comments = {}) {
    rows.forEach((row, rowIndex) => {
        const nextComment = comments[rowIndex] ?? row.comments ?? "";
        commentState.set(rowIndex, nextComment);
    });

    updateRowMetaPanel();
}

function replaceRows(nextRows) {
    const normalizedRows = normalizeRowsByTfm(nextRows);
    rows.splice(0, rows.length, ...normalizedRows);
    baseMarksByRow.splice(0, baseMarksByRow.length, ...normalizedRows.map((row) => ({ ...(row.marks || {}) })));
    commentState.clear();
    confirmationState.clear();
    hasProjectSpecificRows = true;
}

function getContentRows() {
    return rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(({ row }) => !row.section)
        .map(({ row, rowIndex }) => ({
            uid: row.uid,
            tfm: row.tfm,
            description: row.description,
            comments: commentState.get(rowIndex) ?? row.comments ?? "",
            section: false,
            marks: { ...baseMarksByRow[rowIndex] },
        }));
}

function isSectionCollapsed(rowIndex) {
    const row = rows[rowIndex];
    if (!row?.section) {
        return false;
    }

    return Boolean(collapsedSections.get(getSectionKey(row)));
}

function setSectionCollapsed(rowIndex, isCollapsed) {
    const row = rows[rowIndex];
    if (!row?.section) {
        return;
    }

    collapsedSections.set(getSectionKey(row), Boolean(isCollapsed));
}

function toggleSection(rowIndex) {
    setSectionCollapsed(rowIndex, !isSectionCollapsed(rowIndex));
    filterMatrixRows();
}

function getSectionRowCount(sectionIndex) {
    let count = 0;

    for (let index = sectionIndex + 1; index < rows.length; index += 1) {
        if (rows[index].section) {
            break;
        }
        count += 1;
    }

    return count;
}

function getSectionLabel(rowIndex) {
    const row = rows[rowIndex];
    const count = getSectionRowCount(rowIndex);
    const suffix = count === 1 ? "rad" : "rader";
    return `${row.description} (${count} ${suffix})`;
}

function renderRevisionList(revisions) {
    if (!revisionList) {
        return;
    }

    cachedRevisions = Array.isArray(revisions) ? revisions : [];
    renderProjectLibraryStats(cachedProjects, cachedRevisions);

    if (!revisions.length) {
        revisionList.innerHTML = '<p class="project-list-empty">Ingen tidligere versjoner for dette prosjektet ennå.</p>';
        return;
    }

    revisionList.innerHTML = "";

    revisions.forEach((revision, index) => {
        const item = document.createElement("div");
        item.className = "revision-item";

        const meta = document.createElement("div");
        meta.className = "revision-meta";
        const title = document.createElement("strong");
        title.textContent = index === 0 ? "Siste lagring" : `Versjon ${revision.revisionId}`;
        const updated = document.createElement("span");
        updated.textContent = `${new Date(revision.createdAt).toLocaleString("no-NO")} (${formatRelativeTime(revision.createdAt)})`;
        meta.appendChild(title);
        meta.appendChild(updated);

        const button = document.createElement("button");
        button.type = "button";
        button.className = "secondary-button";
        button.textContent = "Gjenopprett";
        button.addEventListener("click", async () => {
            await loadProjectRevision(revision.revisionId);
        });

        item.appendChild(meta);
        item.appendChild(button);
        revisionList.appendChild(item);
    });
}

function renderProjectList(projects) {
    if (!projectList) {
        return;
    }

    cachedProjects = Array.isArray(projects) ? projects : [];
    renderProjectLibraryStats(cachedProjects, cachedRevisions);

    const query = (projectSearchInput?.value || "").trim().toLowerCase();
    const visibleProjects = cachedProjects.filter((project) => !query || String(project.id || "").toLowerCase().includes(query));

    if (!cachedProjects.length) {
        projectList.innerHTML = '<p class="project-list-empty">Ingen prosjekter er lagret ennå.</p>';
        return;
    }

    if (!visibleProjects.length) {
        projectList.innerHTML = '<p class="project-list-empty">Ingen prosjekter matcher søket ditt.</p>';
        return;
    }

    const activeProjectId = getCurrentProjectId();
    projectList.innerHTML = "";

    visibleProjects.forEach((project) => {
        const item = document.createElement("div");
        item.className = "project-item";

        if (project.id === activeProjectId) {
            item.classList.add("active");
        }

        const meta = document.createElement("div");
        meta.className = "project-meta";
        const title = document.createElement("strong");
        title.textContent = project.id;
        const badge = document.createElement("span");
        badge.className = "project-badge";
        badge.textContent = project.id === activeProjectId ? "Aktivt" : "Lagret";
        const updated = document.createElement("span");
        updated.textContent = `Sist oppdatert ${new Date(project.updatedAt).toLocaleString("no-NO")} (${formatRelativeTime(project.updatedAt)})`;
        meta.appendChild(title);
        meta.appendChild(badge);
        meta.appendChild(updated);

        const button = document.createElement("button");
        button.type = "button";
        button.className = "secondary-button";
        button.textContent = "Apne";
        button.addEventListener("click", async () => {
            if (projectIdInput) {
                projectIdInput.value = project.id;
            }

            rememberLastProject(project.id);
            await loadProject();
            await loadProjectList();
            await loadRevisionList(project.id);
        });

        item.appendChild(meta);
        item.appendChild(button);
        projectList.appendChild(item);
    });
}

async function loadProjectList() {
    try {
        const response = await fetch("/api/projects");
        const result = await readResponsePayload(response);

        if (!response.ok) {
            throw new Error(result?.error || `Kunne ikke hente prosjektlisten (HTTP ${response.status}).`);
        }

        renderProjectList(Array.isArray(result.projects) ? result.projects : []);
    } catch (_error) {
        if (projectList) {
            projectList.innerHTML = '<p class="project-list-empty">Prosjektlisten kunne ikke hentes akkurat naa.</p>';
        }
    }
}

async function loadRevisionList(projectId = getCurrentProjectId()) {
    if (!revisionList) {
        return;
    }

    revisionList.innerHTML = '<p class="project-list-empty">Laster versjonshistorikk...</p>';

    try {
        const response = await fetch(`/api/projects/${encodeURIComponent(projectId)}/revisions`);
        const result = await readResponsePayload(response);

        if (!response.ok) {
            throw new Error(result?.error || `Kunne ikke hente versjonshistorikken (HTTP ${response.status}).`);
        }

        renderRevisionList(Array.isArray(result.revisions) ? result.revisions : []);
    } catch (_error) {
        revisionList.innerHTML = '<p class="project-list-empty">Versjonshistorikken kunne ikke hentes akkurat nå.</p>';
    }
}

function scheduleAutosave() {
    if (isApplyingSavedState || isSavingProject) {
        return;
    }

    setAutosaveMessage("Endringer registrert. Autolagrer snart...");

    window.clearTimeout(autosaveTimer);
    autosaveTimer = window.setTimeout(() => {
        saveProject();
    }, 900);
}

function applyProjectState(data) {
    if (!data || typeof data !== "object") {
        return;
    }

    isApplyingSavedState = true;
    if (Array.isArray(data.rowDefinitions) && data.rowDefinitions.length) {
        replaceRows(data.rowDefinitions);
        if (matrixInitialized) {
            rebuildMatrix();
        }
    }
    projectTypeSelect.value = data.projectType || "bolig";
    if (breeamLevelSelect) breeamLevelSelect.value = data.breeamLevel || "none";
    uploadedBhText = data.uploadedBhText || "";
    uploadedOfferDocuments.splice(0, uploadedOfferDocuments.length, ...(Array.isArray(data.offerDocuments) ? data.offerDocuments : []));
    lastOfferAnalysis = data.offerAnalysis || null;
    renderOfferDocumentList();
    renderOfferAnalysis();
    applySavedTueConfig(data.tueConfig, Array.isArray(data.selectedPackages) ? data.selectedPackages : []);
    applySavedMatrix(data.matrixMarks || {});
    applySavedComments(data.comments || {});
    applySavedConfirmations(data.confirmations || {});
    updateAllRiskCells();
    lastBhAnalysis = null;
    renderBhAnalysisInsights();
    applyProjectLogic();
    setWorkflowStep(getRecommendedWorkflowStep(), { scroll: false });
    rememberLastProject(getCurrentProjectId());
    isApplyingSavedState = false;
}

async function saveProject() {
    const projectId = getCurrentProjectId();
    const payload = collectProjectState();

    isSavingProject = true;
    setPersistenceMessage("Lagrer prosjekt...");

    try {
        const response = await fetch(`/api/projects/${encodeURIComponent(projectId)}`, {
            method: "PUT",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify(payload),
        });

        const result = await readResponsePayload(response);

        if (!response.ok) {
            throw new Error(result?.error || `Lagring feilet (HTTP ${response.status}).`);
        }

        setPersistenceMessage(`Prosjekt "${projectId}" lagret ${new Date(result.updatedAt).toLocaleString("no-NO")}.`);
        setAutosaveMessage(`Sist autolagret ${new Date(result.updatedAt).toLocaleTimeString("no-NO")}.`);
        rememberLastProject(projectId);
        await loadProjectList();
        await loadRevisionList(projectId);
    } catch (error) {
        setPersistenceMessage(`Kunne ikke lagre prosjektet. ${error.message}`, true);
        setAutosaveMessage("Autolagring feilet.", true);
    } finally {
        isSavingProject = false;
    }
}

async function loadProject() {
    const projectId = getCurrentProjectId();

    setPersistenceMessage("Henter lagret prosjekt...");

    try {
        const response = await fetch(`/api/projects/${encodeURIComponent(projectId)}`);
        const result = await readResponsePayload(response);

        if (!response.ok) {
            throw new Error(result?.error || `Henting feilet (HTTP ${response.status}).`);
        }

        applyProjectState(result.data);
        setPersistenceMessage(`Prosjekt "${projectId}" hentet ${new Date(result.updatedAt).toLocaleString("no-NO")}.`);
        setAutosaveMessage(`Arbeider i prosjekt "${projectId}".`);
        rememberLastProject(projectId);
        await loadProjectList();
        await loadRevisionList(projectId);
    } catch (error) {
        setPersistenceMessage(`Kunne ikke hente prosjektet. ${error.message}`, true);
    }
}

async function loadProjectRevision(revisionId) {
    const projectId = getCurrentProjectId();

    setPersistenceMessage(`Henter versjon ${revisionId}...`);

    try {
        const response = await fetch(`/api/projects/${encodeURIComponent(projectId)}/revisions/${encodeURIComponent(revisionId)}`);
        const result = await readResponsePayload(response);

        if (!response.ok) {
            throw new Error(result?.error || `Kunne ikke hente versjonen (HTTP ${response.status}).`);
        }

        applyProjectState(result.data);
        setPersistenceMessage(`Versjon ${revisionId} fra ${new Date(result.createdAt).toLocaleString("no-NO")} er lastet inn.`);
        setAutosaveMessage(`Viser versjon ${revisionId}. Lagre for å opprette en ny aktiv versjon.`);
        await loadRevisionList(projectId);
    } catch (error) {
        setPersistenceMessage(`Kunne ikke hente versjonen. ${error.message}`, true);
    }
}

function resetProjectState() {
    isApplyingSavedState = true;
    hasProjectSpecificRows = false;
    uploadedBhText = "";
    uploadedOfferDocuments.length = 0;
    lastOfferAnalysis = null;
    if (projectTypeSelect) {
        projectTypeSelect.value = "bolig";
    }
    if (breeamLevelSelect) {
        breeamLevelSelect.value = "none";
    }
    if (tueCoreModelSelect) {
        tueCoreModelSelect.value = "separate";
    }
    if (tueLocksModelSelect) {
        tueLocksModelSelect.value = "separate";
    }
    if (tueAdkModelSelect) {
        tueAdkModelSelect.value = "el";
    }
    if (bhUploadInput) {
        bhUploadInput.value = "";
    }
    if (offerUploadInput) {
        offerUploadInput.value = "";
    }
    if (bhAnalysisStatus) {
        bhAnalysisStatus.textContent = "Første versjon bruker regelbasert analyse av nøkkelord. AI-tolkning kan bygges på senere.";
    }
    if (offerAnalysisStatus) {
        offerAnalysisStatus.textContent = "Last opp ett eller flere tilbud og sammenlign dem mot gjeldende grensesnittmatrise.";
    }
    packageOptionInputs.forEach((input) => {
        input.checked = false;
    });
    syncTueBuilderUI();
    rows.forEach((_, rowIndex) => {
        disciplines.forEach((discipline) => {
            responsibilities.forEach((responsibility) => {
                const state = baseMarksByRow[rowIndex][`${discipline}:${responsibility}`] || "";
                setCellState(rowIndex, discipline, responsibility, state);
            });
        });
        commentState.set(rowIndex, rows[rowIndex].comments || "");
        setConfirmation(rowIndex, false);
    });
    updateAllRiskCells();
    applyProjectLogic();
    renderOfferDocumentList();
    renderOfferAnalysis();
    isApplyingSavedState = false;
}

async function deleteCurrentProject() {
    const projectId = getCurrentProjectId();
    const shouldDelete = window.confirm(`Er du sikker på at du vil slette prosjekt "${projectId}"? Dette kan ikke angres.`);

    if (!shouldDelete) {
        setPersistenceMessage(`Sletting av prosjekt "${projectId}" ble avbrutt.`);
        return;
    }

    setPersistenceMessage(`Sletter prosjekt "${projectId}"...`);

    try {
        const response = await fetch(`/api/projects/${encodeURIComponent(projectId)}`, {
            method: "DELETE",
        });
        const result = await readResponsePayload(response);

        if (!response.ok) {
            throw new Error(result?.error || `Sletting feilet (HTTP ${response.status}).`);
        }

        const fallbackProjectId = "default";
        if (projectIdInput) {
            projectIdInput.value = fallbackProjectId;
        }
        resetProjectState();
        rememberLastProject(fallbackProjectId);
        setPersistenceMessage(`Prosjekt "${projectId}" er slettet.`);
        setAutosaveMessage(`Byttet til prosjekt "${fallbackProjectId}".`);
        await loadProjectList();
        await loadRevisionList(fallbackProjectId);
    } catch (error) {
        setPersistenceMessage(`Kunne ikke slette prosjektet. ${error.message}`, true);
    }
}

function addNewRow() {
    const tfm = window.prompt("TFM for ny rad:", "999");
    if (tfm === null) {
        return;
    }

    const description = window.prompt("Beskrivelse for ny rad:", "Ny aktivitet");
    if (description === null) {
        return;
    }

    const newRow = {
        uid: createRowId(),
        tfm: tfm.trim() || "999",
        description: description.trim() || "Ny aktivitet",
        comments: "",
        marks: {},
        section: false,
    };

    replaceRows([...getContentRows(), newRow]);
    rebuildMatrix();
    const newRowIndex = rows.findIndex((row) => row.uid === newRow.uid);
    if (newRowIndex >= 0) {
        focusRow(newRowIndex);
    }
    buildContractSummary();
    setPersistenceMessage(`Ny rad "${newRow.description}" er lagt til.`);
    scheduleAutosave();
}

function swapRowState(firstIndex, secondIndex) {
    [rows[firstIndex], rows[secondIndex]] = [rows[secondIndex], rows[firstIndex]];
    [baseMarksByRow[firstIndex], baseMarksByRow[secondIndex]] = [baseMarksByRow[secondIndex], baseMarksByRow[firstIndex]];

    const firstComment = commentState.get(firstIndex) ?? rows[firstIndex].comments ?? "";
    const secondComment = commentState.get(secondIndex) ?? rows[secondIndex].comments ?? "";
    commentState.set(firstIndex, secondComment);
    commentState.set(secondIndex, firstComment);

    const firstConfirmation = Boolean(confirmationState.get(firstIndex));
    const secondConfirmation = Boolean(confirmationState.get(secondIndex));
    confirmationState.set(firstIndex, secondConfirmation);
    confirmationState.set(secondIndex, firstConfirmation);
}

function moveActiveRow(direction) {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        return;
    }

    const currentRowId = rows[activeRowIndex].uid;
    const contentRows = getContentRows();
    const currentContentIndex = contentRows.findIndex((row) => row.uid === currentRowId);
    const targetContentIndex = currentContentIndex + direction;

    if (currentContentIndex < 0 || targetContentIndex < 0 || targetContentIndex >= contentRows.length) {
        return;
    }

    [contentRows[currentContentIndex], contentRows[targetContentIndex]] = [contentRows[targetContentIndex], contentRows[currentContentIndex]];
    replaceRows(contentRows);
    rebuildMatrix();
    const nextActiveIndex = rows.findIndex((row) => row.uid === currentRowId);
    if (nextActiveIndex >= 0) {
        focusRow(nextActiveIndex);
    }
    buildContractSummary();
    setPersistenceMessage(`Raden "${rows[nextActiveIndex].description}" er flyttet ${direction < 0 ? "opp" : "ned"}.`);
    scheduleAutosave();
}

function deleteActiveRow() {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        return;
    }

    const rowToDelete = rows[activeRowIndex];
    const shouldDelete = window.confirm(`Er du sikker på at du vil slette raden "${rowToDelete.tfm} - ${rowToDelete.description}"?`);

    if (!shouldDelete) {
        setPersistenceMessage("Sletting av rad ble avbrutt.");
        return;
    }

    replaceRows(getContentRows().filter((row) => row.uid !== rowToDelete.uid));
    rebuildMatrix();

    const nextActiveIndex = rows.findIndex((row, rowIndex) => !row.section && rowIndex >= activeRowIndex);
    const fallbackIndex = nextActiveIndex >= 0 ? nextActiveIndex : rows.findIndex((row) => !row.section);

    if (fallbackIndex >= 0) {
        focusRow(fallbackIndex);
    } else {
        activeRowIndex = -1;
        updateRowMetaPanel();
    }

    buildContractSummary();
    setPersistenceMessage(`Raden "${rowToDelete.description}" er slettet.`);
    scheduleAutosave();
}

function buildContractSummary() {
    const projectType = projectTypeSelect.value;
    const confirmedCount = getConfirmedRowCount();
    const openRiskCount = getOpenRiskCount();
    const projectText = getProjectTypeLabel(projectType);
    const contentRows = getContentRowCount();
    const completionRate = contentRows ? Math.round((confirmedCount / contentRows) * 100) : 0;
    const readinessLabel = openRiskCount === 0 && completionRate === 100
        ? "Klar for eksport"
        : completionRate >= 70
            ? "Nær klar"
            : "Under arbeid";
    const nextActionText = openRiskCount
        ? "Fokuser på åpne avklaringer, bekreft UE-ansvar og gå deretter gjennom kommentarene før eksport."
        : "Alle rader er avklart. Prosjektet er klart for eksport og en siste kvalitetssjekk.";
    const blockers = rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(({ row, rowIndex }) => !row.section && getRiskState(rowIndex).level !== "ok")
        .slice(0, 4)
        .map(({ row, rowIndex }) => `${row.tfm} ${row.description} - ${getExportRiskLabel(rowIndex)}`);
    const bhSuggestion = lastBhAnalysis?.findings?.length
        ? lastBhAnalysis.findings[0]
        : "Ingen nye BH-signaler registrert ennå.";
    const exportReady = contentRows > 0 && openRiskCount === 0 && completionRate === 100;

    contractSummary.innerHTML = `
        <article class="summary-panel">
            <h3>Vedlegg X - Grensesnittmatrise</h3>
            <p>Prosjekttype: <strong>${projectText}</strong></p>
            <p>Valgt TUE-struktur: <strong>${describeTueConfig()}</strong></p>
            <p>UE bekreftet: <strong>${confirmedCount} av ${contentRows}</strong></p>
            <p>Fremdrift: <strong>${completionRate} %</strong></p>
        </article>
        <article class="summary-panel">
            <h3>Beslutningsstatus</h3>
            <p>Status: <strong>${readinessLabel}</strong></p>
            <p>Åpne avklaringer: <strong>${openRiskCount}</strong></p>
            <p>Neste steg: <strong>${nextActionText}</strong></p>
        </article>
        <article class="summary-panel">
            <h3>BH-signaler</h3>
            <p>Foreslått prosjektspor: <strong>${getProjectTypeLabel(lastBhAnalysis?.projectType || projectType)}</strong></p>
            <p>Viktigste signal: <strong>${escapeHtml(bhSuggestion)}</strong></p>
            <p>Analysepoeng: <strong>${lastBhAnalysis?.keywordScore || 0}</strong></p>
        </article>
        <article class="summary-panel">
            <h3>Prioriterte avklaringer</h3>
            <ul>
                ${
                    blockers.length
                        ? blockers.map((item) => `<li>${escapeHtml(item)}</li>`).join("")
                        : "<li>Ingen åpne avklaringer gjenstår i matrisen.</li>"
                }
            </ul>
        </article>
    `;

    if (workspaceReadinessLabel) {
        workspaceReadinessLabel.textContent = readinessLabel;
    }

    if (workspaceNextAction) {
        workspaceNextAction.textContent = nextActionText;
    }

    if (workspaceBlockers) {
        workspaceBlockers.innerHTML = blockers.length
            ? blockers.map((item) => `<p>${escapeHtml(item)}</p>`).join("")
            : "<p>Ingen åpne avklaringer. Prosjektet er klart for siste eksportkontroll.</p>";
    }

    if (exportExcelButton) {
        exportExcelButton.disabled = !exportReady;
        exportExcelButton.title = exportReady ? "Eksporter prosjektet til CSV/Excel" : "Avklar alle åpne rader og bekreft UE før eksport.";
    }

    if (exportPdfButton) {
        exportPdfButton.disabled = !exportReady;
        exportPdfButton.title = exportReady ? "Eksporter prosjektet til utskrift/PDF" : "Avklar alle åpne rader og bekreft UE før eksport.";
    }

    updateMatrixOverview();
    updateWorkflowOverview();
}

function getProjectTypeLabel(projectType = projectTypeSelect?.value || "bolig") {
    return {
        bolig: "Bolig",
        leilighet: "Leilighetsbygg",
        rekkehus: "Rekkehus / småhus",
        studentbolig: "Studentbolig",
        kontor: "Kontor",
        skole: "Skole",
        barnehage: "Barnehage",
        universitet: "Universitet / campus",
        sykehus: "Sykehus",
        helsehus: "Helsehus / klinikk",
        sykehjem: "Sykehjem / omsorg",
        hotell: "Hotell",
        handel: "Handel / kjøpesenter",
        idrett: "Idrett / svømmehall",
        kultur: "Kultur / forsamlingsbygg",
        logistikk: "Logistikk / lager",
        industri: "Industri",
        verksted: "Verksted",
        datahall: "Datahall",
        laboratorium: "Laboratorium",
        parkering: "P-hus / mobilitet",
        samferdsel: "Samferdsel / terminal",
        forsvar: "Forsvar / beredskap",
        rehab: "Rehabilitering / ombygging",
        mixeduse: "Kombinasjonsbygg",
    }[projectType] || "Prosjekt";
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}

function escapeCsvValue(value) {
    const normalized = String(value ?? "").replaceAll('"', '""').replace(/\r?\n/g, " ");
    return `"${normalized}"`;
}

function getRowExportState(rowIndex, discipline, responsibility) {
    const button = matrixBody.querySelector(
        `button[data-row="${rowIndex}"][data-discipline="${discipline}"][data-responsibility="${responsibility}"]`
    );
    return button?.dataset.state || "";
}

function buildExportTableHtml() {
    const topHeaderCells = disciplines.map((discipline) => `<th colspan="${responsibilities.length}">${escapeHtml(discipline)}</th>`).join("");
    const subHeaderCells = disciplines
        .flatMap(() => responsibilities)
        .map((responsibility) => `<th>${escapeHtml(responsibility)}</th>`)
        .join("");

    const bodyRows = rows
        .map((row, rowIndex) => {
            if (row.section) {
                return `
                    <tr class="section-row">
                        <td>${escapeHtml(row.tfm)}</td>
                        <td>${escapeHtml(getSectionLabel(rowIndex))}</td>
                        <td colspan="${disciplines.length * responsibilities.length}"></td>
                    </tr>
                `;
            }

            const cellMarkup = disciplines
                .flatMap((discipline) =>
                    responsibilities.map((responsibility) => `<td>${escapeHtml(getRowExportState(rowIndex, discipline, responsibility))}</td>`)
                )
                .join("");
            const comment = escapeHtml(commentState.get(rowIndex) ?? row.comments ?? "");
            const risk = escapeHtml(getExportRiskLabel(rowIndex));
            const confirmed = confirmationState.get(rowIndex) ? "Ja" : "Nei";
            const riskClass = getRiskState(rowIndex).level === "ok" ? "risk-ok-export" : "risk-warning-export";

            return `
                <tr>
                    <td>${escapeHtml(row.tfm)}</td>
                    <td>
                        <div class="desc">${escapeHtml(row.description)}</div>
                        ${comment ? `<div class="comment">Kommentar: ${comment}</div>` : ""}
                        <div class="meta"><span class="${riskClass}">${risk}</span> · UE bekreftet: ${confirmed}.</div>
                    </td>
                    ${cellMarkup}
                </tr>
            `;
        })
        .join("");

    return `
        <table class="export-matrix">
            <thead>
                <tr>
                    <th rowspan="2">TFM</th>
                    <th rowspan="2">Beskrivelse</th>
                    ${topHeaderCells}
                </tr>
                <tr>
                    ${subHeaderCells}
                </tr>
            </thead>
            <tbody>${bodyRows}</tbody>
        </table>
    `;
}

function buildExcelExportRows() {
    const headerRow = [
        "TFM",
        "Beskrivelse",
        "Kommentar",
        "Avklaring",
        "UE bekreftet",
        ...disciplines.flatMap((discipline) => responsibilities.map((responsibility) => `${discipline} ${responsibility}`)),
    ];

    const exportRows = [
        ["Prosjekt-ID", getCurrentProjectId()],
        ["Prosjekttype", getProjectTypeLabel()],
        ["TUE-struktur", describeTueConfig()],
        ["Eksportert", new Date().toLocaleString("no-NO")],
        [],
        headerRow,
    ];

    rows.forEach((row, rowIndex) => {
        if (row.section) {
            exportRows.push([row.tfm, getSectionLabel(rowIndex), "", "Seksjon", "", ...new Array(disciplines.length * responsibilities.length).fill("")]);
            return;
        }

        exportRows.push([
            row.tfm,
            row.description,
            commentState.get(rowIndex) ?? row.comments ?? "",
            getExportRiskLabel(rowIndex),
            confirmationState.get(rowIndex) ? "Ja" : "Nei",
            ...disciplines.flatMap((discipline) =>
                responsibilities.map((responsibility) => getRowExportState(rowIndex, discipline, responsibility))
            ),
        ]);
    });

    return exportRows;
}

function exportProjectToExcel() {
    const csvContent = buildExcelExportRows()
        .map((row) => row.map((value) => escapeCsvValue(value)).join(";"))
        .join("\r\n");
    const blob = new Blob([`\uFEFF${csvContent}`], { type: "text/csv;charset=utf-8;" });
    const downloadUrl = window.URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    const safeProjectId = getCurrentProjectId().replace(/[^\w-]+/g, "_");

    downloadLink.href = downloadUrl;
    downloadLink.download = `${safeProjectId || "grensesnittmatrise"}.csv`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    downloadLink.remove();
    window.setTimeout(() => window.URL.revokeObjectURL(downloadUrl), 0);
    setPersistenceMessage(`Excel-eksport er lastet ned for prosjekt "${getCurrentProjectId()}".`);
}

function exportProjectToPrintView() {
    const exportWindow = window.open("", "_blank", "noopener,noreferrer");

    if (!exportWindow) {
        setPersistenceMessage("Nettleseren blokkerte eksportvinduet. Tillat popup-vindu og prøv igjen.", true);
        return;
    }

    const projectId = getCurrentProjectId();
    const projectType = getProjectTypeLabel();
    const tueDescription = describeTueConfig();
    const savedTimestamp = new Date().toLocaleString("no-NO");
    const exportHighlights = buildExportHighlights();
    const actionItems = buildExportActionItems();
    const summaryTable = buildExportTableHtml();
    const exportHtml = `
        <!DOCTYPE html>
        <html lang="no">
        <head>
            <meta charset="UTF-8">
            <title>Grensesnittmatrise - ${escapeHtml(projectId)}</title>
            <style>
                :root {
                    --ink: #1b2529;
                    --muted: #4f5a5f;
                    --line: #cbd7d4;
                    --panel: #f7fbfb;
                    --panel-soft: #fcf8f1;
                    --accent: #0a4d50;
                    --accent-soft: #dceeed;
                    --warn: #8a5600;
                    --warn-soft: #f7e7c7;
                    --ok: #206746;
                    --ok-soft: #dfeee7;
                }
                body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; margin: 24px; color: var(--ink); }
                h1, h2, h3, p { margin-top: 0; }
                .export-header { display: grid; grid-template-columns: 1.8fr 1fr; gap: 18px; margin-bottom: 22px; }
                .card { border: 1px solid var(--line); border-radius: 16px; padding: 16px 18px; background: var(--panel); }
                .hero-card { background: linear-gradient(180deg, #f8fbfb, #eef6f5); }
                .eyebrow { margin: 0 0 10px; text-transform: uppercase; letter-spacing: 0.12em; font-size: 11px; color: var(--accent); font-weight: 700; }
                .meta { color: var(--muted); font-size: 0.95rem; line-height: 1.5; }
                .stats { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-bottom: 18px; }
                .stat { background: white; }
                .stat strong { display: block; font-size: 1.35rem; margin-bottom: 4px; color: var(--accent); }
                .section-grid { display: grid; grid-template-columns: 1.2fr 1fr; gap: 18px; margin-bottom: 22px; }
                .action-list { margin: 0; padding-left: 18px; line-height: 1.6; }
                .legend-row { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 12px; }
                .pill { display: inline-block; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }
                .pill-ok { background: var(--ok-soft); color: var(--ok); }
                .pill-warn { background: var(--warn-soft); color: var(--warn); }
                .export-matrix { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10px; }
                .export-matrix th, .export-matrix td { border: 1px solid #bfc8c6; padding: 4px; text-align: center; vertical-align: top; }
                .export-matrix thead th { background: #d7ebe8; }
                .export-matrix td:nth-child(1) { width: 60px; text-align: left; }
                .export-matrix td:nth-child(2), .export-matrix th:nth-child(2) { width: 280px; text-align: left; }
                .section-row td { background: #e6f0ef; font-weight: 700; }
                .desc { font-weight: 700; margin-bottom: 4px; }
                .comment { margin-bottom: 4px; color: #374247; }
                .meta { color: #4f5a5f; }
                .risk-ok-export { color: var(--ok); font-weight: 700; }
                .risk-warning-export { color: var(--warn); font-weight: 700; }
                .table-title { margin: 0 0 10px; }
                @page { size: A3 landscape; margin: 14mm; }
                @media print {
                    .print-actions { display: none; }
                    body { margin: 0; }
                }
            </style>
        </head>
        <body>
            <div class="print-actions" style="margin-bottom:16px;">
                <button onclick="window.print()">Skriv ut / lagre som PDF</button>
            </div>
            <div class="export-header">
                <div class="card hero-card">
                    <p class="eyebrow">Kontraktsunderlag</p>
                    <h1>Grensesnittmatrise</h1>
                    <p><strong>Prosjekt-ID:</strong> ${escapeHtml(projectId)}</p>
                    <p><strong>Prosjekttype:</strong> ${escapeHtml(projectType)}</p>
                    <p><strong>TUE-struktur:</strong> ${escapeHtml(tueDescription)}</p>
                    <div class="legend-row">
                        <span class="pill pill-ok">OK: Avklart og bekreftet</span>
                        <span class="pill pill-warn">Åpen: Krever oppfølging</span>
                    </div>
                </div>
                <div class="card">
                    <h3>Eksport</h3>
                    <p class="meta">Generert ${escapeHtml(savedTimestamp)}</p>
                    <p class="meta">Domene: grensesnittmatrise.no</p>
                    <p class="meta">Dokumentet er egnet som arbeidsunderlag for koordinering, gjennomgang og vedlegg til kontrakt.</p>
                </div>
            </div>
            <div class="stats">
                <div class="card stat"><strong>${exportHighlights.totalRows}</strong><span>Rader i matrisen</span></div>
                <div class="card stat"><strong>${exportHighlights.confirmedCount}</strong><span>UE-bekreftede rader</span></div>
                <div class="card stat"><strong>${exportHighlights.openRiskCount}</strong><span>Åpne avklaringer</span></div>
                <div class="card stat"><strong>${exportHighlights.completionRate} %</strong><span>Fremdrift</span></div>
            </div>
            <div class="section-grid">
                <div class="card">
                    <h3>Oppsummering</h3>
                    <p class="meta">Kommenterte rader: ${exportHighlights.commentedCount}</p>
                    <p class="meta">Dette dokumentet samler TFM, ansvar, kommentarer og bekreftelser for videre koordinering.</p>
                </div>
                <div class="card" style="background: var(--panel-soft);">
                    <h3>Prioriterte oppfølgingspunkter</h3>
                    <ul class="action-list">
                        ${actionItems.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
                    </ul>
                </div>
            </div>
            <h2 class="table-title">Detaljert matrise</h2>
            ${summaryTable}
        </body>
        </html>
    `;

    exportWindow.document.open();
    exportWindow.document.write(exportHtml);
    exportWindow.document.close();
    setPersistenceMessage(`Eksportvisning er åpnet for prosjekt "${projectId}".`);
}

function clearFocusedRow() {
    matrixBody.querySelectorAll(".row-focus").forEach((row) => row.classList.remove("row-focus"));
}

function focusRow(rowIndex) {
    const row = getRowElement(rowIndex);
    if (!row || row.classList.contains("filtered-out")) {
        return;
    }

    clearFocusedRow();
    row.classList.add("row-focus");
    focusedRowIndex = rowIndex;
    activeRowIndex = rowIndex;
    updateRowMetaPanel();
    row.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
}

function focusRowByUid(rowUid, options = {}) {
    const rowIndex = rows.findIndex((row) => row.uid === rowUid);
    if (rowIndex < 0) {
        return false;
    }

    if (options.step) {
        setWorkflowStep(options.step, { scroll: false });
    }

    const row = rows[rowIndex];
    if (!row.section) {
        setActiveSectionFilter(getRowSectionCode(row), { updateHash: false });
    }

    window.setTimeout(() => {
        focusRow(rowIndex);
    }, 120);

    return true;
}

function filterMatrixRows() {
    const query = (matrixSearchInput?.value || "").trim().toLowerCase();
    const showOpenOnly = Boolean(showOpenOnlyInput?.checked);
    const conflictRowIds = getOfferConflictRowIds();
    let firstVisibleRow = null;
    let visibleContentRows = 0;
    const rowMatches = rows.map((row, rowIndex) => {
        const searchableText = `${row.tfm} ${row.description} ${commentState.get(rowIndex) ?? row.comments ?? ""}`.toLowerCase();
        return !query || searchableText.includes(query);
    });
    const sectionHasMatch = new Map();
    let currentSectionIndex = -1;

    rows.forEach((row, rowIndex) => {
        if (row.section) {
            currentSectionIndex = rowIndex;
            sectionHasMatch.set(rowIndex, false);
            return;
        }

        const sectionMatches = activeSectionFilter === "all" || getRowSectionCode(row) === activeSectionFilter;
        const reviewMatches = activeReviewFilter === "conflicts"
            ? conflictRowIds.has(row.uid)
            : rowMatchesReviewFilter(row, rowIndex);
        const rowMatchesFilter = rowMatches[rowIndex]
            && sectionMatches
            && (!showOpenOnly || getRiskState(rowIndex).level !== "ok")
            && reviewMatches;

        if (currentSectionIndex >= 0 && rowMatchesFilter) {
            sectionHasMatch.set(currentSectionIndex, true);
        }
    });

    currentSectionIndex = -1;
    rows.forEach((row, rowIndex) => {
        const rowElement = getRowElement(rowIndex);
        if (!rowElement) {
            return;
        }

        let isVisible = rowMatches[rowIndex];

        if (row.section) {
            currentSectionIndex = rowIndex;
            isVisible = Boolean(sectionHasMatch.get(rowIndex))
                && (activeSectionFilter === "all" || Number(row.tfm) === activeSectionFilter);
            rowElement.classList.toggle("collapsed-section", isSectionCollapsed(rowIndex));
        } else {
            if (activeSectionFilter !== "all" && getRowSectionCode(row) !== activeSectionFilter) {
                isVisible = false;
            }

            if (showOpenOnly && getRiskState(rowIndex).level === "ok") {
                isVisible = false;
            }

            if (isVisible) {
                isVisible = activeReviewFilter === "conflicts"
                    ? conflictRowIds.has(row.uid)
                    : rowMatchesReviewFilter(row, rowIndex);
            }

            if (currentSectionIndex >= 0 && isSectionCollapsed(currentSectionIndex) && !query && !showOpenOnly) {
                isVisible = false;
            }
        }

        rowElement.classList.toggle("filtered-out", !isVisible);

        if (isVisible && !row.section && firstVisibleRow === null) {
            firstVisibleRow = rowIndex;
        }

        if (isVisible && !row.section) {
            visibleContentRows += 1;
        }
    });

    if (focusedRowIndex >= 0) {
        const focusedRow = getRowElement(focusedRowIndex);
        if (focusedRow?.classList.contains("filtered-out")) {
            clearFocusedRow();
            focusedRowIndex = -1;
            activeRowIndex = -1;
            updateRowMetaPanel();
        }
    }

    if (query && firstVisibleRow !== null && focusedRowIndex === -1) {
        focusRow(firstVisibleRow);
    }

    updateMatrixOverview(visibleContentRows);
    updateMatrixFilterFeedback(visibleContentRows, query, showOpenOnly);
}

function jumpToNextUnresolvedRow() {
    const unresolvedRows = rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(({ row, rowIndex }) => !row.section && getRiskState(rowIndex).level !== "ok");

    if (!unresolvedRows.length) {
        return;
    }

    const visibleUnresolvedRows = unresolvedRows.filter(({ rowIndex }) => {
        const row = getRowElement(rowIndex);
        return row && !row.classList.contains("filtered-out");
    });
    const rowsToUse = visibleUnresolvedRows.length ? visibleUnresolvedRows : unresolvedRows;

    if (!visibleUnresolvedRows.length && matrixSearchInput) {
        matrixSearchInput.value = "";
        filterMatrixRows();
    }

    const nextUnresolved = rowsToUse.find(({ rowIndex }) => rowIndex > focusedRowIndex) || rowsToUse[0];
    focusRow(nextUnresolved.rowIndex);
}

function applyProjectLogic() {
    const projectType = projectTypeSelect.value;
    const tueDescription = describeTueConfig();

    const sdIntegrationActive = projectType !== "bolig";
    const packageMessage = `Valgt TUE-oppsett: ${tueDescription}.`;

    projectLogicStatus.textContent = sdIntegrationActive
        ? `SD-integrasjon er aktiv for ${projectType}. ${packageMessage}`
        : `SD-integrasjon er deaktivert for boligprosjekt. ${packageMessage}`;

    buildContractSummary();
}

function createChoiceCell(rowIndex, discipline, responsibility) {
    const cell = document.createElement("td");
    cell.className = "choice";

    if (responsibility === responsibilities[0]) {
        cell.classList.add("group-start");
    }

    if (responsibility === responsibilities[responsibilities.length - 1]) {
        cell.classList.add("group-end");
    }

    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("aria-label", `${rows[rowIndex].description} - ${discipline} - ${responsibility}`);
    button.dataset.row = String(rowIndex);
    button.dataset.discipline = discipline;
    button.dataset.responsibility = responsibility;
    button.textContent = "";
    button.title = `${discipline} ${responsibility}`;

    // Event handling delegated to matrixBody (see bottom of file)

    const initialState = rows[rowIndex].marks[`${discipline}:${responsibility}`] || "";
    applyState(button, initialState);
    cell.appendChild(button);
    return cell;
}

function createMatrixRow(rowData, rowIndex) {
    const row = document.createElement("tr");
    row.dataset.rowIndex = String(rowIndex);
    confirmationState.set(rowIndex, Boolean(confirmationState.get(rowIndex)));
    commentState.set(rowIndex, commentState.get(rowIndex) ?? rowData.comments ?? "");

    if (rowData.section) {
        row.classList.add("section-row");
    }

    row.addEventListener("click", (event) => {
        if (event.target instanceof HTMLButtonElement) {
            return;
        }

        if (rowData.section) {
            return;
        }

        focusRow(rowIndex);
    });

    const tfmCell = document.createElement("td");
    tfmCell.className = "tfm-cell";
    tfmCell.textContent = rowData.tfm;
    row.appendChild(tfmCell);

    const descriptionCell = document.createElement("td");
    descriptionCell.className = "description-cell";

    if (rowData.section) {
        const toggleButton = document.createElement("button");
        toggleButton.type = "button";
        toggleButton.className = "section-toggle";
        toggleButton.dataset.rowIndex = String(rowIndex);
        toggleButton.textContent = isSectionCollapsed(rowIndex) ? "+" : "-";
        toggleButton.setAttribute("aria-label", `${isSectionCollapsed(rowIndex) ? "Utvid" : "Skjul"} seksjon ${rowData.description}`);
        toggleButton.addEventListener("click", (event) => {
            event.stopPropagation();
            toggleSection(rowIndex);
        });

        const sectionLabel = document.createElement("span");
        sectionLabel.textContent = getSectionLabel(rowIndex);
        descriptionCell.appendChild(toggleButton);
        descriptionCell.appendChild(sectionLabel);
    } else {
        descriptionCell.appendChild(renderRowDescriptionContent(rowIndex));
    }

    row.appendChild(descriptionCell);

    disciplines.forEach((discipline) => {
        responsibilities.forEach((responsibility) => {
            row.appendChild(createChoiceCell(rowIndex, discipline, responsibility));
        });
    });

    return row;
}

function buildMatrix() {
    rows.forEach((rowData, rowIndex) => {
        matrixBody.appendChild(createMatrixRow(rowData, rowIndex));
    });
}

function getVisibleRowIndices() {
    const indices = [];
    for (let i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (activeSectionFilter !== "all") {
            var code = getRowSectionCode(row);
            // Include section headers for the active filter
            if (row.section && Number(row.tfm) !== activeSectionFilter) continue;
            if (!row.section && code !== activeSectionFilter) continue;
        }
        indices.push(i);
    }
    return indices;
}

function buildMatrixInBatches(batchSize = 40) {
    matrixBody.innerHTML = "";
    matrixBuildInProgress = true;

    var visibleIndices = getVisibleRowIndices();

    return new Promise((resolve) => {
        let cursor = 0;

        const renderBatch = () => {
            const fragment = document.createDocumentFragment();
            const end = Math.min(cursor + batchSize, visibleIndices.length);

            for (let i = cursor; i < end; i++) {
                fragment.appendChild(createMatrixRow(rows[visibleIndices[i]], visibleIndices[i]));
            }

            matrixBody.appendChild(fragment);
            cursor = end;

            if (workflowProgressText) {
                const percent = visibleIndices.length ? Math.round((cursor / visibleIndices.length) * 100) : 100;
                workflowProgressText.textContent = `Bygger matrise... ${percent} %`;
            }

            if (cursor < visibleIndices.length) {
                window.requestAnimationFrame(renderBatch);
                return;
            }

            matrixBuildInProgress = false;
            resolve();
        };

        window.requestAnimationFrame(renderBatch);
    });
}

function rebuildMatrix() {
    matrixBody.innerHTML = "";
    activeRowIndex = -1;
    focusedRowIndex = -1;
    buildMatrix();
    updateAllRiskCells();
    filterMatrixRows();
}

function markHeaderGroups() {
    const topHeaders = document.querySelectorAll(".matrix thead tr:first-child th[colspan='6']");
    const subHeaders = document.querySelectorAll(".matrix thead tr.subhead th");

    topHeaders.forEach((header) => {
        header.classList.add("group-start", "group-end");
    });

    subHeaders.forEach((header, index) => {
        if (index % responsibilities.length === 0) {
            header.classList.add("group-start");
        }

        if ((index + 1) % responsibilities.length === 0) {
            header.classList.add("group-end");
        }
    });
}

workflowStepButtons.forEach((button) => button.addEventListener("click", () => {
    const targetStep = Number(button.dataset.stepTarget || 1);
    setWorkflowStep(targetStep);
}));

projectTypeSelect.addEventListener("change", () => {
    applyProjectLogic();
    scheduleAutosave();
});
projectSearchInput?.addEventListener("input", () => {
    renderProjectList(cachedProjects);
});
tueCoreModelSelect?.addEventListener("change", () => {
    syncTueBuilderUI();
    applyProjectLogic();
    scheduleAutosave();
});
tueLocksModelSelect?.addEventListener("change", () => {
    syncTueBuilderUI();
    applyProjectLogic();
    scheduleAutosave();
});
tueAdkModelSelect?.addEventListener("change", () => {
    syncTueBuilderUI();
    applyProjectLogic();
    scheduleAutosave();
});
packageOptionInputs.forEach((input) => input.addEventListener("change", () => {
    syncTueBuilderUI();
    applyProjectLogic();
    scheduleAutosave();
}));
refreshSummaryButton.addEventListener("click", buildContractSummary);
exportExcelButton?.addEventListener("click", exportProjectToExcel);
exportPdfButton?.addEventListener("click", exportProjectToPrintView);
applyPackagePresetButton.addEventListener("click", () => {
    ensureMatrixInitialized();
    applyPackagePreset();
    setWorkflowStep(3, { scroll: false });
    scheduleAutosave();
});
matrixSearchInput?.addEventListener("input", filterMatrixRows);
showOpenOnlyInput?.addEventListener("change", filterMatrixRows);
addRowButton?.addEventListener("click", () => {
    ensureMatrixInitialized();
    addNewRow();
});
deleteRowButton?.addEventListener("click", () => {
    ensureMatrixInitialized();
    deleteActiveRow();
});
moveRowUpButton?.addEventListener("click", () => {
    ensureMatrixInitialized();
    moveActiveRow(-1);
});
moveRowDownButton?.addEventListener("click", () => {
    ensureMatrixInitialized();
    moveActiveRow(1);
});
jumpUnresolvedButton?.addEventListener("click", () => {
    ensureMatrixInitialized();
    jumpToNextUnresolvedRow();
});
quickConfirmRowButton?.addEventListener("click", () => {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        return;
    }

    setConfirmation(activeRowIndex, true);
    buildContractSummary();
    scheduleAutosave();
});
quickNextUnresolvedButton?.addEventListener("click", () => {
    ensureMatrixInitialized();
    jumpToNextUnresolvedRow();
});
quickClearCommentButton?.addEventListener("click", () => {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        return;
    }

    commentState.set(activeRowIndex, "");
    if (currentRowComment) {
        currentRowComment.value = "";
    }
    updateRowMetaPanel();
    scheduleAutosave();
});
currentRowConfirm?.addEventListener("change", () => {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        currentRowConfirm.checked = false;
        return;
    }

    setConfirmation(activeRowIndex, currentRowConfirm.checked);
    buildContractSummary();
    scheduleAutosave();
});
currentRowTfm?.addEventListener("input", () => {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        currentRowTfm.value = "";
        return;
    }

    rows[activeRowIndex].tfm = currentRowTfm.value;
    updateRowDisplay(activeRowIndex);
    scheduleAutosave();
});
currentRowDescription?.addEventListener("input", () => {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        currentRowDescription.value = "";
        return;
    }

    rows[activeRowIndex].description = currentRowDescription.value;
    updateRowDisplay(activeRowIndex);
    scheduleAutosave();
});
currentRowComment?.addEventListener("input", () => {
    if (activeRowIndex < 0 || rows[activeRowIndex]?.section) {
        currentRowComment.value = "";
        return;
    }

    commentState.set(activeRowIndex, currentRowComment.value);
    scheduleAutosave();
});
saveProjectButton?.addEventListener("click", saveProject);
loadProjectButton?.addEventListener("click", loadProject);
refreshProjectListButton?.addEventListener("click", async () => {
    await loadProjectList();
    await loadRevisionList(getCurrentProjectId());
});
newProjectButton?.addEventListener("click", async () => {
    const newProjectId = `prosjekt-${Date.now()}`;
    if (projectIdInput) {
        projectIdInput.value = newProjectId;
    }
    resetProjectState();
    rememberLastProject(newProjectId);
    setPersistenceMessage(`Nytt prosjekt "${newProjectId}" er klart.`);
    setAutosaveMessage("Nytt prosjekt opprettet. Autolagrer ved første endring.");
    setWorkflowStep(1, { scroll: false });
    loadProjectList();
    loadRevisionList(newProjectId);
});
deleteProjectButton?.addEventListener("click", deleteCurrentProject);
projectIdInput?.addEventListener("change", () => {
    rememberLastProject(getCurrentProjectId());
    loadProjectList();
    loadRevisionList(getCurrentProjectId());
    setWorkflowStep(1, { scroll: false });
    scheduleAutosave();
});
analyzeBhButton.addEventListener("click", () => {
    const sourceText = `${uploadedBhText}`.trim();

    if (!sourceText) {
        bhAnalysisStatus.textContent = "Last opp ett eller flere dokumenter fra byggherre først.";
        lastBhAnalysis = null;
        renderBhAnalysisInsights();
        return;
    }

    const analysis = applyBhSuggestionsFromText(sourceText);
    applyProjectLogic();
    bhAnalysisStatus.textContent = `Underlaget er analysert. ${analysis.keywordScore} signaler ble funnet, og forslag til prosjekttype/TUE er oppdatert. Trykk deretter på 'Bruk pakkeoppsett i matrisen' for å klargjore utsendelsesgrunnlaget.`;
    setWorkflowStep(2, { scroll: false });
    scheduleAutosave();
});
bhUploadInput.addEventListener("change", async () => {
    const [file] = bhUploadInput.files || [];

    if (!file) {
        uploadedBhText = "";
        return;
    }

    uploadedBhText = await file.text();
    bhAnalysisStatus.textContent = `Lastet inn underlag: ${file.name}. Klar for analyse.`;
    setWorkflowStep(2, { scroll: false });
    scheduleAutosave();
});
analyzeOffersButton?.addEventListener("click", analyzeOffersAgainstMatrix);
renderBhAnalysisInsights();
renderProjectLibraryStats();
renderOfferAnalysis();
updateWorkflowOverview();
offerFindingsList?.addEventListener("click", (event) => {
    const target = event.target instanceof HTMLElement ? event.target.closest("[data-row-uid]") : null;
    const rowUid = target?.getAttribute("data-row-uid");
    if (!rowUid) {
        return;
    }

    focusRowByUid(rowUid, { step: 3 });
});
jumpConflictRowButton?.addEventListener("click", () => {
    const conflictRowUid = lastOfferAnalysis?.findings?.find((finding) => finding.rowUid)?.rowUid;
    if (conflictRowUid) {
        focusRowByUid(conflictRowUid, { step: 3 });
    }
});
jumpUncommentedRowButton?.addEventListener("click", () => {
    const firstCommentGap = getRowsNeedingComment()[0];
    if (firstCommentGap) {
        focusRow(firstCommentGap.rowIndex);
    }
});
focusOfferStepButton?.addEventListener("click", () => {
    setWorkflowStep(4);
});

document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        matrixSearchInput?.focus();
        matrixSearchInput?.select();
        return;
    }

    const target = event.target;
    const isTypingTarget = target instanceof HTMLElement
        && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable);
    const isMatrixButton = target instanceof HTMLButtonElement && target.closest(".matrix");

    if (isTypingTarget) {
        if (event.key === "Escape" && target === matrixSearchInput && matrixSearchInput?.value) {
            matrixSearchInput.value = "";
            filterMatrixRows();
        }
        return;
    }

    if (isMatrixButton) {
        return;
    }

    if (event.key === "ArrowDown") {
        event.preventDefault();
        focusAdjacentContentRow(1);
        return;
    }

    if (event.key === "ArrowUp") {
        event.preventDefault();
        focusAdjacentContentRow(-1);
        return;
    }

    if (event.key === "Escape" && matrixSearchInput?.value) {
        matrixSearchInput.value = "";
        filterMatrixRows();
    }
});

async function initializeApp() {
    initializeRows(defaultRows);
    syncTueBuilderUI();
    applyProjectLogic();
    applyReviewFilter(getSavedReviewFilter(), { skipRefilter: true });
    applyReviewMode(getSavedReviewMode());
    activeSectionFilter = getSectionFilterFromHash();
    setWorkflowStep(1, { scroll: false });
    updateAllRiskCells();

    const rememberedProjectId = getRememberedProject();
    if (rememberedProjectId && projectIdInput) {
        projectIdInput.value = rememberedProjectId;
    }

    loadProjectList();
    loadRevisionList();
    if (rememberedProjectId) {
        loadProject();
    }

    updateWorkflowOverview();
    setWorkflowStep(getRecommendedWorkflowStep(), { scroll: false });
}

initializeApp();

toggleReviewModeButton?.addEventListener("click", () => {
    applyReviewMode(!reviewModeEnabled);
});

// ── Chapter tabs (delegated) ──
var chapterTabNav = document.querySelector(".chapter-tabs");
if (chapterTabNav) {
    chapterTabNav.addEventListener("click", function(e) {
        var tab = e.target.closest(".chapter-tab");
        if (!tab) return;
        chapterTabNav.querySelector(".chapter-tab.active")?.classList.remove("active");
        tab.classList.add("active");
        var chapter = tab.getAttribute("data-chapter");
        setActiveSectionFilter(chapter === "all" ? "all" : Number(chapter));
    });
}

// ── Delegated matrix button events (replaces per-cell listeners) ──
if (matrixBody) {
    matrixBody.addEventListener("click", function(e) {
        var btn = e.target.closest("button[data-row]");
        if (!btn) return;
        var ri = Number(btn.dataset.row);
        var disc = btn.dataset.discipline;
        var resp = btn.dataset.responsibility;
        setResponsibilityValue(ri, disc, resp, nextState(btn.dataset.state || ""));
    });

    matrixBody.addEventListener("focusin", function(e) {
        var btn = e.target.closest("button[data-row]");
        if (!btn) return;
        focusRow(Number(btn.dataset.row));
    });

    matrixBody.addEventListener("keydown", function(e) {
        var btn = e.target.closest("button[data-row]");
        if (!btn) return;
        var ri = Number(btn.dataset.row);
        var disc = btn.dataset.discipline;
        var resp = btn.dataset.responsibility;
        var key = e.key;

        if (key === "ArrowRight") { e.preventDefault(); moveMatrixButtonFocus(btn, 0, 1); }
        else if (key === "ArrowLeft") { e.preventDefault(); moveMatrixButtonFocus(btn, 0, -1); }
        else if (key === "ArrowDown") { e.preventDefault(); moveMatrixButtonFocus(btn, 1, 0); }
        else if (key === "ArrowUp") { e.preventDefault(); moveMatrixButtonFocus(btn, -1, 0); }
        else if (key === " " || key === "Spacebar") { e.preventDefault(); setResponsibilityValue(ri, disc, resp, nextState(btn.dataset.state || "")); }
        else if (key.toLowerCase() === "h") { e.preventDefault(); setResponsibilityValue(ri, disc, resp, "H"); }
        else if (key.toLowerCase() === "d") { e.preventDefault(); setResponsibilityValue(ri, disc, resp, "D"); }
        else if (key === "Delete" || key === "Backspace") { e.preventDefault(); setResponsibilityValue(ri, disc, resp, ""); }
    });
}

reviewFilterButtons.forEach((button) => {
    button.addEventListener("click", () => {
        applyReviewFilter(button.dataset.reviewFilter || "all");
    });
});

matrixSectionResetButton?.addEventListener("click", () => {
    setActiveSectionFilter("all");
});

matrixSectionFirstRowButton?.addEventListener("click", () => {
    focusFirstVisibleContentRow();
});

matrixSectionNextOpenButton?.addEventListener("click", () => {
    const openRows = getVisibleContentRowIndexes({ openOnly: true });

    if (!openRows.length) {
        showToast("Ingen apne avklaringer i dette utvalget akkurat na.", "success");
        return;
    }

    const nextOpen = openRows.find((rowIndex) => rowIndex > focusedRowIndex) ?? openRows[0];
    focusRow(nextOpen);
});

window.addEventListener("hashchange", () => {
    setActiveSectionFilter(getSectionFilterFromHash(), { updateHash: false });
});

// ── Toast notification system ──
const toastContainer = document.getElementById("toast-container");

function showToast(message, type = "info", duration = 3500) {
    if (!toastContainer) return;
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    const icons = { success: "OK", error: "!", info: "i" };
    const iconEl = document.createElement("span");
    iconEl.className = "toast-icon";
    iconEl.textContent = icons[type] || icons.info;
    const msgEl = document.createElement("span");
    msgEl.textContent = message;
    toast.appendChild(iconEl);
    toast.appendChild(msgEl);
    toastContainer.appendChild(toast);
    setTimeout(() => {
        toast.classList.add("toast-leaving");
        setTimeout(() => toast.remove(), 300);
    }, duration);
}

// ── Back to top button ──
const backToTopButton = document.getElementById("back-to-top");

if (backToTopButton) {
    window.addEventListener("scroll", () => {
        backToTopButton.classList.toggle("visible", window.scrollY > 400);
    }, { passive: true });

    backToTopButton.addEventListener("click", () => {
        window.scrollTo({ top: 0, behavior: "smooth" });
    });
}

// ── Sticky topbar scroll effect ──
const siteTopbar = document.getElementById("site-topbar");

if (siteTopbar) {
    window.addEventListener("scroll", () => {
        siteTopbar.classList.toggle("scrolled", window.scrollY > 20);
    }, { passive: true });
}

// ── Topbar step pills sync ──
const topbarStepPills = Array.from(document.querySelectorAll(".topbar-step-pill"));
const topbarProgressFill = document.getElementById("topbar-progress-fill");
const topbarProgressLabel = document.getElementById("topbar-progress-label");

topbarStepPills.forEach((pill) => {
    pill.addEventListener("click", () => {
        const step = Number(pill.dataset.stepTarget);
        if (step) setWorkflowStep(step);
    });
});

// Patch setWorkflowStep to sync topbar pills
const _originalSetWorkflowStep = setWorkflowStep;
setWorkflowStep = function patchedSetWorkflowStep(stepNumber, options) {
    _originalSetWorkflowStep(stepNumber, options);
    const nextStep = Math.max(1, Math.min(4, Number(stepNumber) || 1));
    topbarStepPills.forEach((pill) => {
        pill.classList.toggle("active", Number(pill.dataset.stepTarget) === nextStep);
    });
};

// Patch updateWorkflowOverview to sync topbar progress
const _originalUpdateWorkflowOverview = updateWorkflowOverview;
updateWorkflowOverview = function patchedUpdateWorkflowOverview() {
    _originalUpdateWorkflowOverview();
    const progressText = workflowProgressValue?.textContent || "0 %";
    if (topbarProgressLabel) topbarProgressLabel.textContent = progressText;
    const match = progressText.match(/\d+/);
    const percent = match ? Number(match[0]) : 0;
    if (topbarProgressFill) topbarProgressFill.style.width = `${percent}%`;
};

// ── Loading button helper ──
function withLoading(button, asyncFn) {
    return async function (...args) {
        if (!button || button.classList.contains("is-loading")) return;
        button.classList.add("is-loading");
        try {
            await asyncFn.apply(this, args);
        } finally {
            button.classList.remove("is-loading");
        }
    };
}

// ── Patch save/export to show toasts ──
const _origSaveClick = saveProjectButton?.onclick;

if (saveProjectButton) {
    saveProjectButton.addEventListener("click", () => {
        setTimeout(() => {
            const statusText = persistenceStatus?.textContent || "";
            if (statusText.includes("Lagret") || statusText.includes("ok")) {
                showToast("Prosjekt lagret!", "success");
            } else if (statusText.includes("Kunne ikke") || statusText.includes("feil")) {
                showToast("Lagring feilet.", "error");
            }
        }, 600);
    });
}

if (exportExcelButton) {
    exportExcelButton.addEventListener("click", () => {
        setTimeout(() => showToast("Excel-eksport startet.", "info"), 200);
    });
}

if (exportPdfButton) {
    exportPdfButton.addEventListener("click", () => {
        setTimeout(() => showToast("PDF-eksport startet.", "info"), 200);
    });
}

if (applyPackagePresetButton) {
    applyPackagePresetButton.addEventListener("click", () => {
        showToast("Pakkeoppsett er brukt i matrisen.", "success");
    });
}

if (analyzeBhButton) {
    analyzeBhButton.addEventListener("click", () => {
        setTimeout(() => {
            if (lastBhAnalysis && lastBhAnalysis.keywordScore > 0) {
                showToast(`Analyse ferdig: ${lastBhAnalysis.keywordScore} signaler funnet.`, "success");
            } else {
                showToast("Ingen tydelige signaler funnet.", "info");
            }
        }, 100);
    });
}

// ══════════════════════════════════════════════════════════════
// MULTI-DOCUMENT UPLOAD & COMPLEXITY ANALYSIS SYSTEM
// ══════════════════════════════════════════════════════════════

const docDropzone = document.getElementById("doc-dropzone");
const docListSection = document.getElementById("doc-list-section");
const docList = document.getElementById("doc-list");
const docCountLabel = document.getElementById("doc-count-label");
const clearAllDocsButton = document.getElementById("clear-all-docs");
const complexityResult = document.getElementById("complexity-result");
const complexityFill = document.getElementById("complexity-fill");
const complexityLevel = document.getElementById("complexity-level");
const complexityDescription = document.getElementById("complexity-description");
const tueRecommendationCard = document.getElementById("tue-recommendation-card");
const tueRecSummary = document.getElementById("tue-rec-summary");
const tueRecReason = document.getElementById("tue-rec-reason");
const applyTueRecommendationButton = document.getElementById("apply-tue-recommendation");
const matrixScopeCard = document.getElementById("matrix-scope-card");
const matrixScopeSummary = document.getElementById("matrix-scope-summary");
const matrixScopeDetail = document.getElementById("matrix-scope-detail");
const applyMatrixScopeButton = document.getElementById("apply-matrix-scope");
const breeamLevelSelect = document.getElementById("breeam-level");
const breeamHelp = document.getElementById("breeam-help");
const breeamCard = document.getElementById("breeam-card");
const breeamCardLevel = document.getElementById("breeam-card-level");
const breeamCardDetail = document.getElementById("breeam-card-detail");
const breeamRowCount = document.getElementById("breeam-row-count");
const applyBreeamRowsButton = document.getElementById("apply-breeam-rows");
const offerUploadInput = document.getElementById("offer-upload");
const offerDropzone = document.getElementById("offer-dropzone");
const offerListSection = document.getElementById("offer-list-section");
const offerList = document.getElementById("offer-list");
const offerCountLabel = document.getElementById("offer-count-label");
const clearAllOffersButton = document.getElementById("clear-all-offers");
const analyzeOffersButton = document.getElementById("analyze-offers");
const offerAnalysisStatus = document.getElementById("offer-analysis-status");
const offerAnalysisKpis = document.getElementById("offer-analysis-kpis");
const offerFindingsList = document.getElementById("offer-findings-list");

function formatFileSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getFileExtIcon(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    const map = { txt: "TXT", md: "MD", csv: "CSV", json: "JSON", pdf: "PDF", docx: "DOC", xlsx: "XLS" };
    return map[ext] || "FIL";
}

function addDocument(name, content, size) {
    uploadedDocuments.push({ name, content, size, id: Date.now() + Math.random() });
    renderDocumentList();
}

function addOfferDocument(name, content, size) {
    uploadedOfferDocuments.push({ name, content, size, id: Date.now() + Math.random() });
    lastOfferAnalysis = null;
    renderOfferDocumentList();
    renderOfferAnalysis();
}

function removeDocument(id) {
    const idx = uploadedDocuments.findIndex(function(d) { return d.id === id; });
    if (idx >= 0) uploadedDocuments.splice(idx, 1);
    renderDocumentList();
}

function renderDocumentList() {
    if (!docListSection || !docList) return;

    if (uploadedDocuments.length === 0) {
        docListSection.hidden = true;
        return;
    }

    docListSection.hidden = false;
    if (docCountLabel) {
        docCountLabel.textContent = uploadedDocuments.length === 1
            ? "1 dokument"
            : `${uploadedDocuments.length} dokumenter`;
    }

    docList.innerHTML = "";
    uploadedDocuments.forEach(function(doc) {
        const item = document.createElement("div");
        item.className = "doc-item";
        item.innerHTML = `
            <span class="doc-item-icon">${getFileExtIcon(doc.name)}</span>
            <div class="doc-item-meta">
                <div class="doc-item-name" title="${escapeHtml(doc.name)}">${escapeHtml(doc.name)}</div>
                <div class="doc-item-size">${formatFileSize(doc.size)} · ${doc.content.length} tegn</div>
            </div>
        `;
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "doc-item-remove";
        removeBtn.textContent = "×";
        removeBtn.title = "Fjern dokument";
        removeBtn.addEventListener("click", function() { removeDocument(doc.id); });
        item.appendChild(removeBtn);
        docList.appendChild(item);
    });
}

function removeOfferDocument(id) {
    const idx = uploadedOfferDocuments.findIndex(function(d) { return d.id === id; });
    if (idx >= 0) uploadedOfferDocuments.splice(idx, 1);
    lastOfferAnalysis = null;
    renderOfferDocumentList();
    renderOfferAnalysis();
}

function renderOfferDocumentList() {
    if (!offerListSection || !offerList) return;

    if (uploadedOfferDocuments.length === 0) {
        offerListSection.hidden = true;
        return;
    }

    offerListSection.hidden = false;
    if (offerCountLabel) {
        offerCountLabel.textContent = uploadedOfferDocuments.length === 1
            ? "1 dokument"
            : `${uploadedOfferDocuments.length} dokumenter`;
    }

    offerList.innerHTML = "";
    uploadedOfferDocuments.forEach(function(doc) {
        const item = document.createElement("div");
        item.className = "doc-item";
        item.innerHTML = `
            <span class="doc-item-icon">${getFileExtIcon(doc.name)}</span>
            <div class="doc-item-meta">
                <div class="doc-item-name" title="${escapeHtml(doc.name)}">${escapeHtml(doc.name)}</div>
                <div class="doc-item-size">${formatFileSize(doc.size)} · ${doc.content.length} tegn</div>
            </div>
        `;
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "doc-item-remove";
        removeBtn.textContent = "×";
        removeBtn.title = "Fjern tilbud";
        removeBtn.addEventListener("click", function() { removeOfferDocument(doc.id); });
        item.appendChild(removeBtn);
        offerList.appendChild(item);
    });
}

function processFiles(fileList, addCallback) {
    Array.from(fileList).forEach(function(file) {
        const name = file.name.toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
                    let text = "";
                    workbook.SheetNames.forEach(function(sheetName) {
                        const sheet = workbook.Sheets[sheetName];
                        text += "--- " + sheetName + " ---\n";
                        text += XLSX.utils.sheet_to_csv(sheet) + "\n\n";
                    });
                    addCallback(file.name, text.trim(), file.size);
                    showToast(`"${file.name}" (${workbook.SheetNames.length} ark) lagt til.`, "info");
                } catch (err) {
                    showToast(`Kunne ikke lese "${file.name}": ${err.message}`, "error");
                }
            };
            reader.readAsArrayBuffer(file);
        } else if (file.name.toLowerCase().endsWith(".pdf")) {
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
                    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(e.target.result) }).promise;
                    let text = "";
                    for (let i = 1; i <= pdf.numPages; i++) {
                        const page = await pdf.getPage(i);
                        const content = await page.getTextContent();
                        text += content.items.map(function(item) { return item.str; }).join(" ") + "\n";
                    }
                    addCallback(file.name, text.trim(), file.size);
                    showToast(`"${file.name}" (${pdf.numPages} sider) lagt til.`, "info");
                } catch (err) {
                    showToast(`Kunne ikke lese "${file.name}": ${err.message}`, "error");
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            const reader = new FileReader();
            reader.onload = function(e) {
                const content = e.target.result || "";
                addCallback(file.name, content, file.size);
                showToast(`"${file.name}" lagt til.`, "info");
            };
            reader.readAsText(file);
        }
    });
}

// Prevent browser from opening dropped files anywhere on the page
document.addEventListener("dragover", function(e) { e.preventDefault(); });
document.addEventListener("drop", function(e) { e.preventDefault(); });

// Drag & drop + click to open file picker
if (docDropzone) {
    docDropzone.addEventListener("click", function(e) {
        if (e.target.closest(".doc-item-remove")) return;
        if (bhUploadInput) bhUploadInput.click();
    });

    ["dragenter", "dragover"].forEach(function(evt) {
        docDropzone.addEventListener(evt, function(e) {
            e.preventDefault();
            e.stopPropagation();
            docDropzone.classList.add("drag-over");
        });
    });

    docDropzone.addEventListener("dragleave", function(e) {
        e.preventDefault();
        if (!docDropzone.contains(e.relatedTarget)) {
            docDropzone.classList.remove("drag-over");
        }
    });

    docDropzone.addEventListener("drop", function(e) {
        e.preventDefault();
        e.stopPropagation();
        docDropzone.classList.remove("drag-over");
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
            handleFiles(e.dataTransfer.files);
        }
    });
}

if (bhUploadInput) {
    bhUploadInput.addEventListener("change", function() {
        if (bhUploadInput.files && bhUploadInput.files.length) {
            handleFiles(bhUploadInput.files);
            bhUploadInput.value = "";
        }
    });
}

function handleFiles(fileList) {
    processFiles(fileList, addDocument);
}

if (clearAllDocsButton) {
    clearAllDocsButton.addEventListener("click", function() {
        uploadedDocuments.length = 0;
        renderDocumentList();
        showToast("Alle dokumenter fjernet.", "info");
    });
}

if (offerDropzone) {
    offerDropzone.addEventListener("click", function(e) {
        if (e.target.closest(".doc-item-remove")) return;
        if (offerUploadInput) offerUploadInput.click();
    });

    ["dragenter", "dragover"].forEach(function(evt) {
        offerDropzone.addEventListener(evt, function(e) {
            e.preventDefault();
            e.stopPropagation();
            offerDropzone.classList.add("drag-over");
        });
    });

    offerDropzone.addEventListener("dragleave", function(e) {
        e.preventDefault();
        if (!offerDropzone.contains(e.relatedTarget)) {
            offerDropzone.classList.remove("drag-over");
        }
    });

    offerDropzone.addEventListener("drop", function(e) {
        e.preventDefault();
        e.stopPropagation();
        offerDropzone.classList.remove("drag-over");
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
            processFiles(e.dataTransfer.files, addOfferDocument);
        }
    });
}

if (offerUploadInput) {
    offerUploadInput.addEventListener("change", function() {
        if (offerUploadInput.files && offerUploadInput.files.length) {
            processFiles(offerUploadInput.files, addOfferDocument);
            offerUploadInput.value = "";
        }
    });
}

if (clearAllOffersButton) {
    clearAllOffersButton.addEventListener("click", function() {
        uploadedOfferDocuments.length = 0;
        renderOfferDocumentList();
        lastOfferAnalysis = null;
        renderOfferAnalysis();
        showToast("Alle tilbud fjernet.", "info");
    });
}

function renderOfferAnalysis() {
    if (offerAnalysisKpis) {
        const summary = lastOfferAnalysis || { documentCount: 0, findingCount: 0, conflictCount: 0, warningCount: 0 };
        offerAnalysisKpis.innerHTML = `
            <div class="overview-card"><span class="overview-label">Tilbud</span><strong>${summary.documentCount || 0}</strong><span class="overview-detail">Opplastede dokumenter</span></div>
            <div class="overview-card"><span class="overview-label">Funn</span><strong>${summary.findingCount || 0}</strong><span class="overview-detail">Registrerte signaler</span></div>
            <div class="overview-card"><span class="overview-label">Konflikter</span><strong>${summary.conflictCount || 0}</strong><span class="overview-detail">Mulige avvik mot matrisen</span></div>
            <div class="overview-card"><span class="overview-label">Advarsler</span><strong>${summary.warningCount || 0}</strong><span class="overview-detail">Forbehold og uklart omfang</span></div>
        `;
    }

    if (offerFindingsList) {
        const findings = lastOfferAnalysis?.findings || [];
        if (!findings.length) {
            offerFindingsList.innerHTML = "<p>Ingen analyse kjørt ennå.</p>";
        } else {
            offerFindingsList.innerHTML = findings.map(function(finding) {
                if (finding.rowUid) {
                    return `<button type="button" class="offer-finding-item" data-row-uid="${escapeHtml(finding.rowUid)}"><strong>${escapeHtml(finding.level)}</strong>: ${escapeHtml(finding.message)}</button>`;
                }
                return `<p><strong>${escapeHtml(finding.level)}</strong>: ${escapeHtml(finding.message)}</p>`;
            }).join("");
        }
    }

    if (activeReviewFilter === "conflicts") {
        filterMatrixRows();
    }

    updateMatrixCommandCenter();
}

function getOfferKeywordsForRow(row) {
    return uniqueList(
        String(row.description || "")
            .toLowerCase()
            .split(/[^a-zA-Z0-9æøåÆØÅ]+/)
            .filter(function(part) { return part.length >= 4; })
    ).slice(0, 6);
}

async function analyzeOffersAgainstMatrix() {
    await ensureMatrixInitialized({ focusFirstRow: false });

    const offerParts = uploadedOfferDocuments.map(function(doc) { return doc.content; });
    const offerText = offerParts.join("\n\n").toLowerCase();
    if (!offerText.trim()) {
        lastOfferAnalysis = null;
        if (offerAnalysisStatus) {
            offerAnalysisStatus.textContent = "Last opp ett eller flere tilbud først.";
        }
        renderOfferAnalysis();
        return;
    }

    const findings = [];
    const disclaimerSignals = ["ikke inkludert", "ikke medtatt", "medtas ikke", "unntatt", "opsjon", "forbehold", "avklares", "annen entrepren", "byggherre leverer", "bh leverer"];

    if (getOpenRiskCount() > 0) {
        findings.push({
            level: "Advarsel",
            message: `Matrisen har fortsatt ${getOpenRiskCount()} åpne avklaringer. Tilbud bør vurderes mot et så lukket grunnlag som mulig.`,
        });
    }

    getContentRows().forEach(function(row) {
        const keywords = getOfferKeywordsForRow(row);
        if (!keywords.length) return;

        const keywordHit = keywords.some(function(keyword) { return offerText.includes(keyword); });
        if (!keywordHit) return;

        const disclaimerHit = disclaimerSignals.find(function(signal) { return offerText.includes(signal); });
        if (!disclaimerHit) return;

        findings.push({
            level: "Konflikt",
            rowUid: row.uid,
            rowTfm: row.tfm,
            rowDescription: row.description,
            message: `${row.tfm} ${row.description}: tilbudet nevner mulig forbehold eller avgrensning ("${disclaimerHit}") og bør kontrolleres mot matrisen.`,
        });
    });

    const generalSignals = [
        { needle: "opsjon", level: "Advarsel", message: "Tilbudet inneholder opsjoner. Sjekk at opsjoner ikke erstatter omfang som er satt som grunnkrav i matrisen." },
        { needle: "forbehold", level: "Advarsel", message: "Tilbudet inneholder forbehold. Gå gjennom om disse strider mot satt ansvar eller leveranseomfang." },
        { needle: "ikke inkludert", level: "Advarsel", message: "Tilbudet oppgir at noe ikke er inkludert. Sammenlign dette mot relevante matriserader." },
    ];

    generalSignals.forEach(function(signal) {
        if (offerText.includes(signal.needle)) {
            findings.push({ level: signal.level, message: signal.message });
        }
    });

    if (!findings.length) {
        findings.push({
            level: "Info",
            message: "Ingen tydelige motstridssignaler ble funnet i første kontroll. Gjennomgå likevel tilbudene manuelt før kontrahering.",
        });
    }

    lastOfferAnalysis = {
        documentCount: uploadedOfferDocuments.length,
        findingCount: findings.length,
        conflictCount: findings.filter(function(item) { return item.level === "Konflikt"; }).length,
        warningCount: findings.filter(function(item) { return item.level === "Advarsel"; }).length,
        findings,
    };

    if (offerAnalysisStatus) {
        offerAnalysisStatus.textContent = `Tilbudsanalyse ferdig. ${lastOfferAnalysis.findingCount} funn registrert mot gjeldende matrisegrunnlag.`;
    }

    renderOfferAnalysis();
    updateWorkflowOverview();
}

// ══════════════════════════════════════════════════════════════
// COMPLEXITY ANALYSIS ENGINE
// ══════════════════════════════════════════════════════════════

const complexityKeywords = {
    // SD/BAS signals (high complexity)
    sdBas: {
        keywords: ["sd-anlegg", "bas-anlegg", "sd/bas", "bygningsautomasjon", "sentraldrift", "ddc",
                    "bacnet", "modbus", "lon", "knx", "dali", "toppsystem"],
        weight: 8,
        label: "SD/BAS-styring"
    },
    // Automation signals
    automation: {
        keywords: ["frekvensomformer", "vfd", "automatikk", "automasjon", "reguleringsventil",
                    "styreventil", "motorstyrt", "pid-reguler"],
        weight: 6,
        label: "Automasjon"
    },
    // Access control
    accessControl: {
        keywords: ["adgangskontroll", "adk", "kortleser", "nøkkelkort", "passersystem",
                    "innbruddsalarm", "alarm-anlegg", "tyverialarm"],
        weight: 5,
        label: "Adgangskontroll"
    },
    // Fire/safety
    fireSafety: {
        keywords: ["brannalarm", "brannvarsl", "sprinkler", "nødlys", "ledesystem",
                    "brannspjeld", "røykdetektor", "brannventilasjon", "trykksetting"],
        weight: 5,
        label: "Brann og sikkerhet"
    },
    // Cooling
    cooling: {
        keywords: ["kjølemaskin", "kjøleanlegg", "varmepumpe", "chiller", "kjølebaffel",
                    "fancoil", "komfortkjøling", "prosesskjøling", "frikjøling"],
        weight: 6,
        label: "Kjøling"
    },
    // Ventilation complexity
    ventilation: {
        keywords: ["vav-system", "dvc", "behovsstyrt", "roterende varmegjenvinner", "kryssvarme",
                    "aggregat", "tilluftsaggregat", "avtrekksvifte", "kanalanlegg"],
        weight: 4,
        label: "Ventilasjon"
    },
    // Electrical complexity
    electrical: {
        keywords: ["hovedfordeling", "underfordeling", "trafo", "nødstrøm", "ups",
                    "reservekraft", "dieselaggregat", "likestrøm", "nødstrømsaggregat"],
        weight: 5,
        label: "Elkraft"
    },
    // Integration signals
    integration: {
        keywords: ["integrasjon", "grensesnitt", "tverrfaglig", "koordinering",
                    "samordning", "felles skap", "felles kabelgate", "ip-nettverk"],
        weight: 3,
        label: "Tverrfaglig integrasjon"
    },
    // Sanitær complexity
    sanitary: {
        keywords: ["legionella", "varmtvannsberedning", "fjernvarme", "energimåler",
                    "vannbehandling", "sirkulasjonspumpe", "tappevannssentral", "energibrønn"],
        weight: 4,
        label: "Sanitær/VVS"
    },
    // Locks & hardware
    locks: {
        keywords: ["lås og beslag", "beslag", "dørlukker", "elektrisk sluttstykke",
                    "motorlås", "dørmagneter", "dørautomatikk"],
        weight: 4,
        label: "Lås og beslag"
    },
    // Scale indicators
    scale: {
        keywords: ["storskala", "storkjøkken", "auditorium", "svømmehall",
                    "operasjonsstue", "laboratorium", "cleanroom", "renrom",
                    "datasenter", "datahall", "serverrom"],
        weight: 7,
        label: "Spesialtilpasninger"
    },
    // Simple indicators (negative complexity)
    simple: {
        keywords: ["enebolig", "hytte", "garasje", "carport", "bod"],
        weight: -5,
        label: "Enkelt prosjekt"
    },
    // BREEAM / environmental certification
    breeam: {
        keywords: ["breeam", "breeam-nor", "miljøsertifiser", "energimerke", "eos",
                    "energioppfølging", "undermåler", "undermåling", "sub-metering",
                    "commissioning", "igangkjøring", "sesongtest", "lekkasjedeteksjon",
                    "dagslysstyring", "voc-sensor", "co2-sensor", "vannbesparende",
                    "lysforurensning", "solcelle", "fornybar energi"],
        weight: 7,
        label: "BREEAM / miljøsertifisering"
    }
};

const projectTypeComplexityBase = {
    bolig: 10, leilighet: 25, rekkehus: 15, studentbolig: 30,
    kontor: 40, skole: 45, barnehage: 30, universitet: 55,
    sykehus: 85, helsehus: 55, sykehjem: 45, hotell: 50,
    handel: 35, idrett: 50, kultur: 45, logistikk: 25,
    industri: 40, verksted: 25, datahall: 75, laboratorium: 70,
    parkering: 20, samferdsel: 55, forsvar: 65, rehab: 35, mixeduse: 50
};

function analyzeComplexity(allText) {
    const text = allText.toLowerCase();
    const signals = [];
    let rawScore = 0;

    Object.entries(complexityKeywords).forEach(function(entry) {
        const category = entry[0];
        var config = entry[1];
        let hitCount = 0;

        config.keywords.forEach(function(kw) {
            const regex = new RegExp(kw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
            const matches = text.match(regex);
            if (matches) hitCount += matches.length;
        });

        if (hitCount > 0) {
            const categoryScore = config.weight * Math.min(hitCount, 5);
            rawScore += categoryScore;
            signals.push({
                category: category,
                label: config.label,
                hits: hitCount,
                score: categoryScore,
                weight: config.weight
            });
        }
    });

    // Add project type base score
    const projectType = projectTypeSelect ? projectTypeSelect.value : "bolig";
    const baseScore = projectTypeComplexityBase[projectType] || 20;
    rawScore += baseScore;

    // Document volume bonus
    const docVolumeBonus = Math.min(uploadedDocuments.length * 3, 25);
    rawScore += docVolumeBonus;

    // Text length bonus (more text = more complex project typically)
    const textLengthBonus = Math.min(Math.floor(text.length / 2000), 15);
    rawScore += textLengthBonus;

    // Normalize to 0-100
    const normalizedScore = Math.max(0, Math.min(100, rawScore));

    let level, levelLabel, description;
    if (normalizedScore < 30) {
        level = "low";
        levelLabel = "Enkel";
        description = "Prosjektet ser ut til å ha lav teknisk kompleksitet. En forenklet matrise med grunnleggende rader er anbefalt.";
    } else if (normalizedScore < 60) {
        level = "medium";
        levelLabel = "Middels";
        description = "Prosjektet har moderat kompleksitet. Standard matrise med typiske tekniske grensesnitt anbefales.";
    } else {
        level = "high";
        levelLabel = "Kompleks";
        description = "Prosjektet har høy teknisk kompleksitet. Full matrise med alle relevante rader bør brukes, og TUE-strukturen bør vurderes nøye.";
    }

    return {
        score: normalizedScore,
        level: level,
        levelLabel: levelLabel,
        description: description,
        signals: signals.sort(function(a, b) { return b.score - a.score; }),
        projectType: projectType,
        baseScore: baseScore,
        docCount: uploadedDocuments.length,
        textLength: text.length,
        tueRecommendation: deriveTueRecommendation(signals, normalizedScore, projectType),
        matrixScope: deriveMatrixScope(signals, normalizedScore, projectType)
    };
}

function deriveTueRecommendation(signals, score, _projectType) {
    const signalCategories = new Set(signals.map(function(s) { return s.category; }));
    const hasSD = signalCategories.has("sdBas");
    const hasAutomation = signalCategories.has("automation");
    const hasAccessControl = signalCategories.has("accessControl");
    const hasLocks = signalCategories.has("locks");
    const hasCooling = signalCategories.has("cooling");
    const hasElectrical = signalCategories.has("electrical");

    // Complex projects with many signals: totalteknisk
    if (score >= 70 && hasSD && hasAutomation && hasElectrical) {
        return {
            coreModel: "totaltechnical",
            locksModel: hasLocks || hasAccessControl ? "separate" : "integrated",
            adkModel: hasAccessControl ? "locks" : "el",
            summary: "Totalteknisk pakke anbefales",
            reason: "Prosjektet har mange tverrfaglige avhengigheter mellom EL, AUT og SD. " +
                    "Med en totalteknisk pakke reduseres grensesnittene betydelig, og én aktør " +
                    "får helhetsansvar for teknisk koordinering."
        };
    }

    // SD + Automation: EL + AUT + SD
    if (hasSD && hasAutomation) {
        return {
            coreModel: "el_aut_sd",
            locksModel: hasLocks || hasAccessControl ? "separate" : "integrated",
            adkModel: hasAccessControl ? "locks" : "el",
            summary: "EL + AUT + SD i felles pakke anbefales",
            reason: "Underlaget nevner både SD/BAS-signaler og automasjonskomponenter. " +
                    "Å samle disse i én leveranse gir enklere grensesnitt og bedre koordinering " +
                    "mellom styrings- og automatiseringsfagene."
        };
    }

    // Automation present: EL + AUT
    if (hasAutomation || (hasCooling && hasElectrical)) {
        return {
            coreModel: "el_aut",
            locksModel: hasLocks || hasAccessControl ? "separate" : "integrated",
            adkModel: hasAccessControl ? "locks" : "el",
            summary: "EL + AUT i felles pakke anbefales",
            reason: "Prosjektet har automasjonsavhengigheter som gjør det fornuftig å samle " +
                    "EL og AUT. SD kan fortsatt håndteres separat for tydeligere grensesnitt."
        };
    }

    // Simple project: separate
    return {
        coreModel: "separate",
        locksModel: hasLocks ? "separate" : "integrated",
        adkModel: "el",
        standalone: [],
        summary: "Separate tekniske UE-er anbefales",
        reason: "Prosjektet ser ut til å ha relativt tydelige faggrenser. " +
                "Separate UE-er gir mest fleksibilitet ved kontrahering og tydeligst ansvarsfordeling."
    };
}

function deriveMatrixScope(signals, score, projectType) {
    const signalCategories = new Set(signals.map(function(s) { return s.category; }));

    // Define which TFM sections/keywords are relevant per complexity level
    let relevantKeywords = [];
    let excludeKeywords = [];
    let rowEstimate;

    if (score < 30) {
        // Simple: only basic rows
        relevantKeywords = ["pumpe", "sanitær", "ventil", "kabel", "brannalarm", "dør"];
        excludeKeywords = ["frekvensomformer", "sd-anlegg", "adgangskontroll", "kjølemaskin",
                           "legionella", "energimåler", "nødstrøm", "ups", "reservekraft",
                           "bacnet", "modbus", "vav", "dvc"];
        rowEstimate = "20-40 rader";
        return {
            level: "minimal",
            label: "Forenklet matrise",
            description: `For et enkelt ${getProjectTypeLabel(projectType).toLowerCase()}-prosjekt trenger du bare grunnleggende rader. ` +
                         `Systemet fjerner avanserte rader som ikke er relevante. Anslagsvis ${rowEstimate}.`,
            relevantKeywords: relevantKeywords,
            excludeKeywords: excludeKeywords,
            rowEstimate: rowEstimate
        };
    }

    if (score < 60) {
        relevantKeywords = [];
        excludeKeywords = [];
        if (!signalCategories.has("cooling")) excludeKeywords.push("kjølemaskin", "chiller", "komfortkjøling");
        if (!signalCategories.has("accessControl")) excludeKeywords.push("adgangskontroll", "kortleser", "passersystem");
        if (!signalCategories.has("scale")) excludeKeywords.push("storkjøkken", "auditorium", "laboratorium");
        rowEstimate = "40-80 rader";
        return {
            level: "standard",
            label: "Standard matrise",
            description: `Middels kompleksitet for ${getProjectTypeLabel(projectType).toLowerCase()}. ` +
                         `Matrisen tilpasses basert på identifiserte signaler. Anslagsvis ${rowEstimate}.`,
            relevantKeywords: relevantKeywords,
            excludeKeywords: excludeKeywords,
            rowEstimate: rowEstimate
        };
    }

    // High complexity: full matrix
    rowEstimate = "80-265 rader";
    return {
        level: "full",
        label: "Komplett matrise",
        description: `Høy kompleksitet for ${getProjectTypeLabel(projectType).toLowerCase()}. ` +
                     `Alle tilgjengelige rader fra databasen bør brukes for å sikre at ingen grensesnitt går tapt. ` +
                     `Anslagsvis ${rowEstimate}.`,
        relevantKeywords: [],
        excludeKeywords: [],
        rowEstimate: rowEstimate
    };
}

// ══════════════════════════════════════════════════════════════
// BREEAM-NOR v6 — GRENSESNITT FOR MILJØSERTIFISERING
// ══════════════════════════════════════════════════════════════

const breeamLevelLabels = {
    none: "Ingen",
    pass: "Pass",
    good: "Good",
    very_good: "Very Good",
    excellent: "Excellent",
    outstanding: "Outstanding"
};

const breeamLevelOrder = ["pass", "good", "very_good", "excellent", "outstanding"];

function breeamLevelIndex(level) {
    var idx = breeamLevelOrder.indexOf(level);
    return idx >= 0 ? idx : -1;
}

const breeamRows = [
    {
        tfm: "800", description: "BREEAM-NOR v6 — Miljøsertifisering",
        comments: "", marks: {}, section: true, minLevel: "pass"
    },
    {
        tfm: "800", description: "Energimåling — varme (undermåler per system)",
        comments: "BREEAM Ene 01: Separat måling av varmeforbruk. Rør leverer følerlommer og vannmålere. Aut integrerer mot SD. EL leverer strømmålere for el-varme.",
        marks: { "Rør:P": "H", "Rør:L": "H", "Rør:M": "H", "EL:K": "H", "Aut:P": "D", "Aut:I": "H", "SD:I": "H" },
        minLevel: "pass"
    },
    {
        tfm: "800", description: "Energimåling — kjøling (undermåler per system)",
        comments: "BREEAM Ene 01: Separat måling av kjøleforbruk. Rør leverer målere på kjølekretser. Aut integrerer mot EOS/SD.",
        marks: { "Rør:P": "H", "Rør:L": "H", "Rør:M": "H", "EL:K": "H", "Aut:P": "D", "Aut:I": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "800", description: "Energimåling — ventilasjon (undermåler per aggregat)",
        comments: "BREEAM Ene 01: Måling av energiforbruk per ventilasjonsaggregat. EL leverer KWh-måler. Aut leser av og sender til SD/EOS.",
        marks: { "Vent:P": "D", "EL:P": "H", "EL:L": "H", "EL:M": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "pass"
    },
    {
        tfm: "800", description: "Energimåling — belysning (undermåler per sone/etasje)",
        comments: "BREEAM Ene 01: Separat måling av lysforbruk per etasje eller sone. EL leverer undermålere i fordeling.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:M": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "800", description: "Energimåling — stikkontakter / utstyr (undermåler)",
        comments: "BREEAM Ene 01: Separat måling av utstyrsforbruk. EL leverer undermålere.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:M": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "very_good"
    },
    {
        tfm: "800", description: "Energioppfølgingssystem (EOS) — toppsystem",
        comments: "BREEAM Ene 01: Samling av alle undermålere i et EOS med trendlogging og alarmgrenser. SD/Aut prosjekterer integrasjon. Alle fag leverer målere med kommunikasjon (BACnet/Modbus).",
        marks: { "EL:I": "H", "Aut:P": "H", "Aut:L": "H", "Aut:F": "H", "Aut:I": "H", "SD:P": "H", "SD:F": "H", "SD:I": "H" },
        minLevel: "pass"
    },
    {
        tfm: "810", description: "CO₂-sensorer i oppholdsrom",
        comments: "BREEAM Hea 02: CO₂-måling i alle rom med varig opphold. Aut/SD prosjekterer, leverer og integrerer. Vent tilpasser kanaler for behovsstyring. EL kabling.",
        marks: { "Vent:P": "D", "Vent:I": "H", "EL:K": "H", "Aut:P": "H", "Aut:L": "H", "Aut:M": "H", "Aut:F": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "810", description: "VOC-sensorer i oppholdsrom",
        comments: "BREEAM Hea 02 (Excellent+): VOC-måling for å dokumentere inneluftkvalitet. Aut leverer og integrerer mot SD. EL kabling.",
        marks: { "EL:K": "H", "Aut:P": "H", "Aut:L": "H", "Aut:M": "H", "Aut:F": "H", "SD:I": "H" },
        minLevel: "excellent"
    },
    {
        tfm: "810", description: "Fuktsensorer i våtrom / tekniske rom",
        comments: "BREEAM Hea 02: Fuktsensoring for å sikre akseptabelt inneklima og forebygge fuktskader. Aut integrerer mot SD.",
        marks: { "EL:K": "H", "Aut:P": "H", "Aut:L": "H", "Aut:M": "H", "Aut:F": "H", "SD:I": "H" },
        minLevel: "very_good"
    },
    {
        tfm: "810", description: "Sonestyring temperatur (individuell per sone)",
        comments: "BREEAM Hea 04: Individuelle temperatursoner med separat regulering. Aut prosjekterer soneløsning. Rør/Vent dimensjonerer for soneinndeling.",
        marks: { "Rør:P": "D", "Vent:P": "D", "Aut:P": "H", "Aut:L": "H", "Aut:F": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "820", description: "Dagslysstyring — automatisk dimming iht. dagslys",
        comments: "BREEAM Ene 04: Automatisk dimming av belysning basert på tilgjengelig dagslys. EL prosjekterer og leverer lysstyringsanlegg med dagslyssensorer.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:M": "H", "EL:K": "H", "EL:F": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "820", description: "Tilstedeværelsessensorer for belysning",
        comments: "BREEAM Ene 04: Automatisk av/på basert på bevegelse/tilstedeværelse. EL prosjekterer sensorplassering og styring.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:M": "H", "EL:K": "H", "EL:F": "H" },
        minLevel: "pass"
    },
    {
        tfm: "820", description: "Utvendig belysning med astronomisk ur / tidstyring",
        comments: "BREEAM Ene 04/Pol 04: Utvendig belysning med tidsur eller astronomisk klokke for å unngå lysforurensning.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:M": "H", "EL:F": "H", "Aut:I": "D", "SD:I": "D" },
        minLevel: "good"
    },
    {
        tfm: "830", description: "Vannmåling — forbruksmåler per system (KV, VV, hagevanning)",
        comments: "BREEAM Wat 01: Separat vannmåling per forbrukskategori. Rør leverer målere med pulsutgang eller bus. Aut integrerer mot SD/EOS.",
        marks: { "Rør:P": "H", "Rør:L": "H", "Rør:M": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "pass"
    },
    {
        tfm: "830", description: "Lekkasjedeteksjon — automatisk varsling ved vannlekkasje",
        comments: "BREEAM Wat 02: Lekkasjedetektorer ved kritiske punkter (teknisk rom, sjakter). Aut/SD varsler ved utløst alarm.",
        marks: { "Rør:P": "H", "Rør:L": "H", "Rør:M": "H", "EL:K": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "830", description: "Vannbesparende armaturer — dokumentasjon og beregning",
        comments: "BREEAM Wat 01: Alle sanitærarmaturer dokumenteres med maks vannforbruk (l/min). Rør prosjekterer og spesifiserer.",
        marks: { "Rør:P": "H", "Rør:L": "H", "Rør:F": "H" },
        minLevel: "pass"
    },
    {
        tfm: "840", description: "Utvidet igangkjøring (commissioning) — sesongtest",
        comments: "BREEAM Man 04: Alle tekniske systemer testes i både varme- og kjølesesong. Aut/SD koordinerer sesongtesting. Krever testplan og dokumentasjon.",
        marks: { "Rør:F": "D", "Vent:F": "D", "EL:F": "D", "Aut:P": "H", "Aut:F": "H", "SD:P": "H", "SD:F": "H", "SD:I": "H" },
        minLevel: "very_good"
    },
    {
        tfm: "840", description: "Funksjonstest og integrert systemtest (IST)",
        comments: "BREEAM Man 04: Dokumentert integrert systemtest der alle tekniske systemer verifiseres i samspill.",
        marks: { "Rør:F": "D", "Vent:F": "D", "EL:F": "D", "Aut:P": "H", "Aut:F": "H", "SD:F": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "810", description: "Behovsstyrt ventilasjon (DCV) med CO₂/temp/tilstedeværelse",
        comments: "BREEAM Ene 02 + Hea 02: Ventilasjonsmengde styres av sensorer i rom. Vent dimensjonerer for DCV. Aut leverer VAV-spjeld/aktuatorer.",
        marks: { "Vent:P": "H", "Vent:L": "H", "Vent:M": "H", "Vent:F": "D", "EL:K": "H", "Aut:P": "H", "Aut:L": "H", "Aut:F": "H", "SD:I": "H" },
        minLevel: "good"
    },
    {
        tfm: "820", description: "Lysforurensningsanalyse — utendørsbelysning",
        comments: "BREEAM Pol 04 (Outstanding): Dokumentert analyse av lysforurensning. EL gjennomfører beregning og velger armaturer med riktig avskjerming.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:F": "H" },
        minLevel: "outstanding"
    },
    {
        tfm: "840", description: "Fleksible tekniske føringer for fremtidig ombygging",
        comments: "BREEAM Wst 06: Tekniske sjakter og føringsveier dimensjoneres med reservekapasitet for fremtidige endringer.",
        marks: { "Rør:P": "D", "Vent:P": "D", "EL:P": "H", "EL:L": "H", "Aut:P": "D" },
        minLevel: "very_good"
    },
    {
        tfm: "800", description: "Solcelleanlegg (PV) — produksjonsmåling og integrasjon",
        comments: "BREEAM Ene 01/04: Solcelleproduksjon måles separat og integreres i EOS. EL leverer vekselretter og måler. SD logger produksjon.",
        marks: { "EL:P": "H", "EL:L": "H", "EL:M": "H", "EL:K": "H", "EL:F": "H", "Aut:I": "H", "SD:I": "H" },
        minLevel: "very_good"
    }
];

function getBreeamLevel() {
    return breeamLevelSelect ? breeamLevelSelect.value : "none";
}

function getFilteredBreeamRows(level) {
    if (level === "none") return [];
    var lvlIdx = breeamLevelIndex(level);
    if (lvlIdx < 0) return [];
    return breeamRows.filter(function(row) {
        return breeamLevelIndex(row.minLevel) <= lvlIdx;
    });
}

function getBreeamDescription(level) {
    var descriptions = {
        pass: "Grunnleggende BREEAM-krav. Energimåling, vannbesparende tiltak og tilstedeværelsesstyring er påkrevd.",
        good: "Krever undermåling per energipost, CO₂-styrt ventilasjon, dagslysstyring, lekkasjedeteksjon og integrert systemtest.",
        very_good: "Utvidet undermåling, fuktsensorer, sesongcommissioning, fleksible føringsveier og solcelleintegrasjon.",
        excellent: "Inkluderer VOC-sensorer for inneluftkvalitet og strenge krav til måling, integrasjon og dokumentasjon.",
        outstanding: "Høyeste nivå. Lysforurensningsanalyse, full EOS-integrasjon og alle BREEAM-grensesnitt aktiveres."
    };
    return descriptions[level] || "";
}

function renderBreeamCard(level) {
    if (!breeamCard) return;
    if (level === "none") {
        breeamCard.hidden = true;
        return;
    }
    breeamCard.hidden = false;
    if (breeamCardLevel) breeamCardLevel.textContent = "BREEAM-NOR v6 — " + breeamLevelLabels[level];
    if (breeamCardDetail) breeamCardDetail.textContent = getBreeamDescription(level);
    var filteredRows = getFilteredBreeamRows(level);
    var contentRows = filteredRows.filter(function(r) { return !r.section; });
    if (breeamRowCount) breeamRowCount.textContent = contentRows.length + " BREEAM-grensesnitt legges til i matrisen.";
}

if (breeamLevelSelect) {
    breeamLevelSelect.addEventListener("change", function() {
        var level = breeamLevelSelect.value;
        if (breeamHelp) {
            breeamHelp.textContent = level === "none" ? "" : getBreeamDescription(level);
        }
        renderBreeamCard(level);
        scheduleAutosave();
    });
}

if (applyBreeamRowsButton) {
    applyBreeamRowsButton.addEventListener("click", async function() {
        var level = getBreeamLevel();
        if (level === "none") {
            showToast("Velg et BREEAM-nivå i prosjektinnstillinger først.", "error");
            return;
        }

        var breeamFiltered = getFilteredBreeamRows(level);
        if (!breeamFiltered.length) return;

        // Remove any existing BREEAM rows (TFM 800-849)
        var cleaned = rows.filter(function(row) {
            var tfmNum = parseInt(row.tfm, 10);
            return isNaN(tfmNum) || tfmNum < 800 || tfmNum >= 850;
        });

        var combined = cleaned.concat(breeamFiltered);
        replaceRows(combined);

        if (matrixInitialized) {
            matrixInitialized = false;
            matrixBuildInProgress = false;
            await ensureMatrixInitialized({ focusFirstRow: false });
        }

        var contentCount = breeamFiltered.filter(function(r) { return !r.section; }).length;
        showToast(
            contentCount + " BREEAM-NOR v6 (" + breeamLevelLabels[level] + ") grensesnitt lagt til i matrisen.",
            "success",
            5000
        );
        scheduleAutosave();
    });
}

function getAllDocumentText() {
    const parts = [];
    uploadedDocuments.forEach(function(doc) { parts.push(doc.content); });
    if (uploadedBhText) parts.push(uploadedBhText);
    return parts.join("\n\n");
}

function renderComplexityResult(result) {
    if (!complexityResult) return;
    complexityResult.hidden = false;

    if (complexityFill) {
        complexityFill.style.width = result.score + "%";
        complexityFill.className = "complexity-fill " + result.level;
    }
    if (complexityLevel) {
        complexityLevel.textContent = result.levelLabel + " (" + result.score + ")";
        complexityLevel.className = "complexity-level " + result.level;
    }
    if (complexityDescription) {
        complexityDescription.textContent = result.description;
    }
}

function renderTueRecommendation(rec) {
    if (!tueRecommendationCard) return;
    tueRecommendationCard.hidden = false;
    if (tueRecSummary) tueRecSummary.textContent = rec.summary;
    if (tueRecReason) tueRecReason.textContent = rec.reason;
}

function renderMatrixScope(scope) {
    if (!matrixScopeCard) return;
    matrixScopeCard.hidden = false;
    if (matrixScopeSummary) matrixScopeSummary.textContent = scope.label;
    if (matrixScopeDetail) matrixScopeDetail.textContent = scope.description;
}

// Enhanced analyze button - runs complexity analysis on all documents
const _origAnalyzeBhClick = analyzeBhButton ? analyzeBhButton.onclick : null;

if (analyzeBhButton) {
    analyzeBhButton.addEventListener("click", function() {
        const allText = getAllDocumentText();
        if (!allText.trim()) {
            showToast("Ingen dokumenter eller tekst å analysere.", "error");
            return;
        }

        // Run original BH analysis
        applyBhSuggestionsFromText(allText);

        // Run complexity analysis
        const result = analyzeComplexity(allText);
        lastComplexityResult = result;

        renderComplexityResult(result);
        renderTueRecommendation(result.tueRecommendation);
        renderMatrixScope(result.matrixScope);

        // Auto-detect BREEAM from document text and show card
        var breeamLevel = getBreeamLevel();
        var hasBreeamSignal = result.signals.some(function(s) { return s.category === "breeam"; });
        if (hasBreeamSignal && breeamLevel === "none") {
            // Suggest BREEAM based on detected signals
            if (breeamLevelSelect) breeamLevelSelect.value = "very_good";
            breeamLevel = "very_good";
            showToast("BREEAM-signaler funnet i dokumentene. BREEAM-NOR v6 Very Good er foreslått.", "info", 5000);
        }
        renderBreeamCard(breeamLevel);

        // Update analysis status
        if (bhAnalysisStatus) {
            var breeamNote = breeamLevel !== "none" ? (" BREEAM: " + breeamLevelLabels[breeamLevel] + ".") : "";
            bhAnalysisStatus.textContent =
                `Analysert ${uploadedDocuments.length} dokument(er). ` +
                `Kompleksitet: ${result.levelLabel} (${result.score}/100). ` +
                `${result.signals.length} signalkategorier identifisert.` + breeamNote;
        }

        showToast(
            `Analyse ferdig: ${result.levelLabel} kompleksitet (${result.score}/100). ${result.signals.length} signaler funnet.`,
            "success",
            5000
        );
    });
}

// Apply TUE recommendation
if (applyTueRecommendationButton) {
    applyTueRecommendationButton.addEventListener("click", function() {
        if (!lastComplexityResult) return;
        const rec = lastComplexityResult.tueRecommendation;

        if (tueCoreModelSelect) tueCoreModelSelect.value = rec.coreModel;
        if (tueLocksModelSelect) tueLocksModelSelect.value = rec.locksModel;
        if (tueAdkModelSelect) tueAdkModelSelect.value = rec.adkModel;

        // Clear standalone checkboxes first
        packageOptionInputs.forEach(function(input) { input.checked = false; });

        if (rec.standalone && rec.standalone.length) {
            rec.standalone.forEach(function(val) {
                const input = packageOptionInputs.find(function(i) { return i.value === val; });
                if (input) input.checked = true;
            });
        }

        syncTueBuilderUI();
        showToast("TUE-anbefaling er brukt i prosjektoppsettet.", "success");
        scheduleAutosave();
    });
}

// Apply matrix scope - filter rows based on complexity
if (applyMatrixScopeButton) {
    applyMatrixScopeButton.addEventListener("click", async function() {
        if (!lastComplexityResult) return;
        const scope = lastComplexityResult.matrixScope;

        showToast("Tilpasser matrisen...", "info");

        // Load full row set from database
        const allRows = await loadExcelRowsData();

        if (scope.level === "full") {
            // Use all rows
            replaceRows(allRows);
        } else {
            // Filter rows based on scope
            const filteredRows = allRows.filter(function(row) {
                if (row.section) return true;
                const desc = (row.description || "").toLowerCase();
                const comment = (row.comments || "").toLowerCase();
                const combined = desc + " " + comment;

                // Check if row matches any exclude keyword
                const isExcluded = scope.excludeKeywords.some(function(kw) {
                    return combined.indexOf(kw.toLowerCase()) >= 0;
                });

                if (isExcluded) return false;

                // For minimal scope, only include rows matching relevant keywords
                if (scope.level === "minimal" && scope.relevantKeywords.length > 0) {
                    return scope.relevantKeywords.some(function(kw) {
                        return combined.indexOf(kw.toLowerCase()) >= 0;
                    });
                }

                return true;
            });

            replaceRows(filteredRows);
        }

        // Rebuild matrix if already initialized
        if (matrixInitialized) {
            matrixInitialized = false;
            matrixBuildInProgress = false;
            usingImportedBaseRows = true;
            hasProjectSpecificRows = true;
            await ensureMatrixInitialized({ focusFirstRow: true });
        } else {
            usingImportedBaseRows = true;
            hasProjectSpecificRows = true;
        }

        const rowCount = rows.filter(function(r) { return !r.section; }).length;
        showToast(`Matrise tilpasset: ${rowCount} rader basert på ${scope.label.toLowerCase()}.`, "success", 5000);
        scheduleAutosave();
    });
}

// ══════════════════════════════════════════════════════════════
// PHASE SIDEBAR (left vertical stepper)
// ══════════════════════════════════════════════════════════════

const phaseBtns = [
    document.getElementById("phase-btn-1"),
    document.getElementById("phase-btn-2"),
    document.getElementById("phase-btn-3"),
    document.getElementById("phase-btn-4"),
];
const phaseStatuses = [
    document.getElementById("phase-status-1"),
    document.getElementById("phase-status-2"),
    document.getElementById("phase-status-3"),
    document.getElementById("phase-status-4"),
];
const phaseLines = [
    document.getElementById("phase-line-1"),
    document.getElementById("phase-line-2"),
    document.getElementById("phase-line-3"),
];

// Phase buttons navigate steps
phaseBtns.forEach(function(btn) {
    if (!btn) return;
    btn.addEventListener("click", function() {
        var target = Number(btn.dataset.stepTarget);
        if (target >= 1 && target <= 4) {
            setWorkflowStep(target);
        }
    });
});

function syncPhaseSidebar() {
    var health = getWorkflowHealth();
    var stepChecks = [health.step1Checks, health.step2Checks, health.step3Checks, health.step4Checks];

    // Determine state per phase
    stepChecks.forEach(function(checks, i) {
        var allDone = checks.every(function(c) { return c.done; });
        var anyDone = checks.some(function(c) { return c.done; });
        var isActive = (i + 1) === currentWorkflowStep;

        var btn = phaseBtns[i];
        var status = phaseStatuses[i];
        if (!btn) return;

        btn.classList.remove("active", "done");
        if (isActive) {
            btn.classList.add("active");
        }
        if (allDone) {
            btn.classList.add("done");
        }

        if (status) {
            if (allDone) {
                status.textContent = "Ferdig";
            } else if (isActive) {
                status.textContent = "Pågår";
            } else if (anyDone) {
                status.textContent = "Startet";
            } else {
                status.textContent = "Venter";
            }
        }

        // Connector line
        if (phaseLines[i]) {
            phaseLines[i].classList.toggle("filled", allDone);
        }
    });

    if (phaseBtns[3]) phaseBtns[3].disabled = false;
}

// Patch updateWorkflowOverview to also sync sidebar
var _prevUpdateWorkflowOverview = updateWorkflowOverview;
updateWorkflowOverview = function patchedUpdateWorkflowOverview2() {
    _prevUpdateWorkflowOverview();
    syncPhaseSidebar();
};

// Patch setWorkflowStep to also sync sidebar active state
var _prevSetWorkflowStep = setWorkflowStep;
setWorkflowStep = function patchedSetWorkflowStep2(stepNumber, options) {
    _prevSetWorkflowStep(stepNumber, options);
    syncPhaseSidebar();
};

// Initial sync
syncPhaseSidebar();
