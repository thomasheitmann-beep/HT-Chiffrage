import React, { useState, useMemo } from "react";
import { Plus, Trash2, Settings2, FileText, ClipboardList, Zap, Download } from "lucide-react";

// ---------------------------------------------------------------------------
// HT MAINTENANCE — OUTIL DE CHIFFRAGE
// Palette : ardoise/graphite + ambre (couleurs d'atelier électrique HT/BT)
// ---------------------------------------------------------------------------

const INK = "#1B2733";
const INK_2 = "#2E3D4E";
const AMBER = "#E8A33D";
const PAPER = "#F5F6F4";
const LINE = "#DCE0E3";
const MUTED = "#8A93A0";

const uid = () => Math.random().toString(36).slice(2, 10);

// ---------------------------------------------------------------------------
// Catalogue repris du tableau réel "Niv 1-4 (standard)" du fichier
// QUO-CHIFFRAGE HT-BT d'origine. Pour chaque équipement : heures totales
// (niv 1-2 + niv 3-4 + préparation, réunies en une seule valeur puisque
// le fichier applique le même taux horaire aux 4 niveaux) et amortissement
// matériel (€, coût fixe indépendant du tarif jour). Le prix de revient
// calculé est : (heures ÷ heures/jour) × prix jour technicien × majoration
// + amortissement. Avec les valeurs par défaut (1 400 €/j technicien,
// 7h/jour = 200 €/h), on retombe exactement sur les prix du fichier source.
// ---------------------------------------------------------------------------

const DEFAULT_TARIFS = {
  technicien: { label: "Technicien (niv. 1-4 standard)", jour: 1400, demiJour: 800 },
  expert: { label: "Expert HT/BT", jour: 1600, demiJour: 950 },
  ingenieur: { label: "Ingénieur", jour: 1800, demiJour: 1100 },
};

const DEFAULT_MAJORATIONS = {
  semaine: { label: "Semaine (jour)", coef: 1 },
  samedi: { label: "Samedi", coef: 1.55 },
  nuit: { label: "Nuit", coef: 2 },
  dimancheFerie: { label: "Dimanche / Férié", coef: 2 },
};

const DEFAULT_DEGRESSIVITE = [
  { min: 0, max: 5, coef: 1, label: "1 à 5 équipements" },
  { min: 5, max: 10, coef: 0.98, label: "5 à 10 équipements" },
  { min: 10, max: 20, coef: 0.95, label: "10 à 20 équipements" },
  { min: 20, max: Infinity, coef: 0.9, label: "Plus de 20 équipements" },
];

const DEFAULT_COEF_CONTRAT = {
  aucun: { label: "Affaire standard", coef: 1 },
  facilitiesManager: { label: "Facilities Manager", coef: 0.95 },
  appelOffre: { label: "Concurrence / Appel d'offre", coef: 0.9 },
  derogation1: { label: "Dérogation commerciale 1", coef: 0.85 },
  derogation2: { label: "Dérogation commerciale 2", coef: 0.8 },
};

// heures = heures nv1-2 + heures nv3-4 + heures prépa (cumulées) — amort = amortissement matériel (€, fixe)
const DEFAULT_CATALOGUE_TEMPS = {
  hta: {
    label: "Cellules HTA",
    marquesRef: "ABB, Schneider Electric (SM6, FLUOKIT, PIX...), Siemens, CEM Gardy, Ormazabal, Pommier",
    items: [
      { id: "hta-interrupteur", label: "Interrupteur", heures: 1.25, amort: 0 },
      { id: "hta-interrupteur-rh", label: "Interrupteur avec relais homopolaire", heures: 1.93, amort: 21.18 },
      { id: "hta-comptage", label: "Comptage", heures: 1.42, amort: 0 },
      { id: "hta-disj-630", label: "Disjoncteur ≤630A (standard)", heures: 5.17, amort: 57.18 },
      { id: "hta-disj-800", label: "Disjoncteur ≥800A (gros calibres)", heures: 6.1, amort: 68.57 },
      { id: "hta-inter-fusible", label: "Inter fusible", heures: 1.42, amort: 0 },
      { id: "hta-inter-fusible-rh", label: "Inter fusible avec relais homopolaire", heures: 2.15, amort: 23.88 },
      { id: "hta-contacteur", label: "Contacteur", heures: 2.9, amort: 0 },
      { id: "hta-contacteur-rp", label: "Contacteur avec relais de protection", heures: 4.0, amort: 42.86 },
    ],
  },
  transfo: {
    label: "Transformateurs",
    marquesRef: "",
    items: [
      { id: "tr-huile-2000", label: "Transfo huile ≤ 2000 kVA", heures: 4.4, amort: 0 },
      { id: "tr-huile-plus2000", label: "Transfo huile > 2000 kVA", heures: 6.3, amort: 0 },
      { id: "tr-prelevement", label: "Prélèvement seul", heures: 1.0, amort: 0 },
      { id: "tr-sec-2000-nc", label: "Transfo sec ≤ 2000 kVA non capoté", heures: 3.7, amort: 58.5 },
      { id: "tr-sec-plus2000-nc", label: "Transfo sec > 2000 kVA non capoté", heures: 4.64, amort: 73.78 },
      { id: "tr-sec-2000-c", label: "Transfo sec ≤ 2000 kVA capoté", heures: 5.1, amort: 81.25 },
      { id: "tr-sec-plus2000-c", label: "Transfo sec > 2000 kVA capoté", heures: 6.3, amort: 100.75 },
    ],
  },
  btSecondaire: {
    label: "Disjoncteurs BT — injection secondaire",
    marquesRef: "Schneider (Masterpact, Compact NS/NSX), Eaton/Moeller, ABB SACE, Siemens, Legrand, GE",
    items: [
      { id: "bts-debro-630", label: "Disj débro ≤630A injection secondaire", heures: 1.2, amort: 25 },
      { id: "bts-debro-3200", label: "Disj débro ≤3200A injection secondaire", heures: 2.2, amort: 50 },
      { id: "bts-debro-4000", label: "Disj débro ≥4000A injection secondaire", heures: 3.0, amort: 70 },
      { id: "bts-fixe-630", label: "Disj fixe ≤630A injection secondaire", heures: 0.85, amort: 16.25 },
      { id: "bts-fixe-3200", label: "Disj fixe ≤3200A injection secondaire", heures: 1.6, amort: 35 },
      { id: "bts-fixe-4000", label: "Disj fixe ≥4000A injection secondaire", heures: 1.95, amort: 43.75 },
      { id: "bts-inter-debro-630", label: "Interrupteur débro ≤630A", heures: 0.71, amort: 0 },
      { id: "bts-inter-debro-plus630", label: "Interrupteur débro >630A", heures: 1.32, amort: 0 },
      { id: "bts-inter-fixe-630", label: "Interrupteur fixe ≤630A", heures: 0.36, amort: 0 },
      { id: "bts-inter-fixe-plus630", label: "Interrupteur fixe >630A", heures: 0.72, amort: 0 },
      { id: "bts-tiroir-inj", label: "Tiroir injection secondaire", heures: 1.6, amort: 35 },
      { id: "bts-tiroir-sans", label: "Tiroir sans injection", heures: 1.11, amort: 0 },
    ],
  },
  btPrimaire: {
    label: "Disjoncteurs BT — injection primaire",
    marquesRef: "",
    items: [
      { id: "btp-debro-630", label: "Disj débro ≤630A injection primaire", heures: 2.86, amort: 112.36 },
      { id: "btp-debro-3200", label: "Disj débro ≤3200A injection primaire", heures: 3.77, amort: 151.54 },
      { id: "btp-debro-4000", label: "Disj débro ≥4000A injection primaire", heures: 6.57, amort: 272.09 },
      { id: "btp-magneto-1250", label: "Disj magnéto-thermique ≤1250A primaire", heures: 6.5, amort: 258.31 },
      { id: "btp-magneto-plus1250", label: "Disj magnéto-thermique >1250A primaire", heures: 7.5, amort: 301.36 },
    ],
  },
  pfcRec: {
    label: "Divers — compensateurs et redresseurs",
    marquesRef: "",
    items: [
      { id: "pfc-bt-5", label: "PFC BT : 400V ≤ 5 gradins", heures: 2.8, amort: 0 },
      { id: "pfc-bt-plus5", label: "PFC BT : 400V > 5 gradins", heures: 4.0, amort: 0 },
      { id: "pfc-hta", label: "PFC HTA : 20kV ≤1500kvar / 5.5kV ≤150kvar", heures: 8.0, amort: 0 },
      { id: "rec-c13", label: "REC type C13-100", heures: 1.56, amort: 0 },
      { id: "rec-sces", label: "REC type SCES", heures: 2.8, amort: 0 },
    ],
  },
};

const DEFAULT_CATALOGUE_COEF = {
  batteries: {
    label: "Batteries",
    unite: "par batterie",
    categories: [
      { id: "bat-a", label: "7 à 12 Ah", min: 7, max: 12, coef: 2.1 },
      { id: "bat-b", label: "13 à 60 Ah", min: 13, max: 60, coef: 2.0 },
      { id: "bat-c", label: "61 à 125 Ah", min: 61, max: 125, coef: 1.7 },
      { id: "bat-d", label: "> 130 Ah", min: 130, max: Infinity, coef: 1.5 },
    ],
    critereLabel: "Capacité (Ah)",
  },
  composants: {
    label: "Composants",
    unite: "par composant",
    categories: [
      { id: "cmp-a", label: "1 € à 30 €", min: 1, max: 30, coef: 3.4 },
      { id: "cmp-b", label: "30 € à 60 €", min: 30, max: 60, coef: 2.9 },
      { id: "cmp-c", label: "60 € à 150 €", min: 60, max: 150, coef: 2.4 },
      { id: "cmp-d", label: "> 150 €", min: 150, max: Infinity, coef: 2.2 },
    ],
    critereLabel: "Prix d'achat unitaire (€)",
  },
};

// Catalogues à prix direct (pas de calcul heures × tarif) : analyses d'huile
// et nettoyage TGBT, dont les prix du fichier source sont fixes, au mètre
// ou à l'unité, indépendants du tarif jour technicien.
const DEFAULT_CATALOGUE_DIRECT = {
  analyseHuile: {
    label: "Analyses d'huile",
    unite: "u",
    items: [
      { id: "ah-pa", label: "Analyse type PA", prix: 257.4 },
      { id: "ah-b", label: "Analyse type B", prix: 378.4 },
      { id: "ah-bdf", label: "Analyse type B + DF", prix: 552.2 },
      { id: "ah-cdf", label: "Analyse type C + DF", prix: 744 },
      { id: "ah-df", label: "Analyse type DF", prix: 286 },
      { id: "ah-sf", label: "Analyse type SF (soufre corrosif)", prix: 154 },
      { id: "ah-td", label: "Analyse type TD (tangente delta)", prix: 165 },
      { id: "ah-pcb", label: "Analyse PCB", prix: 156 },
    ],
  },
  tgbt: {
    label: "TGBT — nettoyage",
    unite: "m",
    items: [
      { id: "tgbt-resserage", label: "Nettoyage + resserrage (mètres)", prix: 600 },
      { id: "tgbt-seul", label: "Nettoyage seul (mètres)", prix: 600 },
    ],
  },
};

// Familles pour les "lignes libres" (batteries, composants, poste manuel) —
// les équipements du catalogue sont renseignés directement par quantité
// dans la checklist, pas via ces lignes.
const FAMILLES_LIBRES = [
  { id: "batteries", label: "Batteries", type: "coef" },
  { id: "composants", label: "Composants", type: "coef" },
  { id: "manuel", label: "Poste libre (manuel)", type: "manuel" },
];

const LABEL_FAMILLE = {
  hta: "Cellules HTA",
  transfo: "Transformateurs",
  btSecondaire: "Disjoncteurs BT (inj. secondaire)",
  btPrimaire: "Disjoncteurs BT (inj. primaire)",
  pfcRec: "Divers — PFC / redresseurs",
  analyseHuile: "Analyses d'huile",
  tgbt: "TGBT — nettoyage",
  batteries: "Batteries",
  composants: "Composants",
  manuel: "Poste libre",
};

function euros(n) {
  if (!isFinite(n)) return "0 €";
  return n.toLocaleString("fr-FR", { maximumFractionDigits: 0 }) + " €";
}

function findCoefCategory(catalogue, valeur) {
  return catalogue.categories.find((c) => valeur >= c.min && valeur <= c.max) || catalogue.categories[0];
}

// ---------------------------------------------------------------------------

function SectionCard({ title, icon: Icon, children, right, subtitle }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="overflow-hidden">
      <div className="flex items-center justify-between px-5 py-4" style={{ borderBottom: `1px solid ${LINE}` }}>
        <div className="flex items-center gap-2">
          {Icon && <Icon size={17} color={INK_2} />}
          <div>
            <h2 style={{ color: INK, fontWeight: 600, fontSize: 15 }}>{title}</h2>
            {subtitle && <div style={{ color: MUTED, fontSize: 11.5 }}>{subtitle}</div>}
          </div>
        </div>
        {right}
      </div>
      <div className="p-5">{children}</div>
    </div>
  );
}

function NumberField({ value, onChange, suffix, width = 90 }) {
  return (
    <span className="inline-flex items-center gap-1">
      <input
        type="number"
        min="0"
        value={value}
        onChange={(e) => onChange(parseFloat(e.target.value) || 0)}
        style={{
          width,
          border: `1px solid ${LINE}`,
          borderRadius: 6,
          padding: "5px 8px",
          fontSize: 13,
          fontVariantNumeric: "tabular-nums",
          color: INK,
        }}
      />
      {suffix && <span style={{ fontSize: 12, color: MUTED }}>{suffix}</span>}
    </span>
  );
}

function TextField({ value, onChange, placeholder, style }) {
  return (
    <input
      type="text"
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: "6px 9px", fontSize: 13, color: INK, width: "100%", ...style }}
    />
  );
}

function Select({ value, onChange, options, style }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: "6px 9px", fontSize: 13, color: INK, background: "#fff", width: "100%", ...style }}
    >
      {options.map((o) => (
        <option key={o.value} value={o.value}>
          {o.label}
        </option>
      ))}
    </select>
  );
}

// ---------------------------------------------------------------------------

export default function ChiffrageHTMaintenance() {
  const [tab, setTab] = useState("chiffrage");

  const [tarifs, setTarifs] = useState(DEFAULT_TARIFS);
  const [majorations, setMajorations] = useState(DEFAULT_MAJORATIONS);
  const [degressivite, setDegressivite] = useState(DEFAULT_DEGRESSIVITE);
  const [coefContrat, setCoefContrat] = useState(DEFAULT_COEF_CONTRAT);
  const [heuresJour, setHeuresJour] = useState(7);

  const [catalogueTemps, setCatalogueTemps] = useState(DEFAULT_CATALOGUE_TEMPS);
  const [catalogueCoef, setCatalogueCoef] = useState(DEFAULT_CATALOGUE_COEF);
  const [catalogueDirect, setCatalogueDirect] = useState(DEFAULT_CATALOGUE_DIRECT);

  const [affaire, setAffaire] = useState({
    client: "",
    site: "",
    reference: "QUO-" + new Date().getFullYear() + "-001",
    contrat: "aucun",
    niveauTechnicien: "technicien",
    typeJournee: "semaine",
  });

  // Quantités par équipement (catalogue temps + catalogue direct) : { itemId: quantite }
  const [quantites, setQuantites] = useState({});
  const setQte = (id, v) => setQuantites((q) => ({ ...q, [id]: Math.max(0, v) }));

  // Lignes libres : batteries, composants, postes manuels (nombre illimité)
  const [lignesLibres, setLignesLibres] = useState([]);
  const [nbAAjouter, setNbAAjouter] = useState(1);

  const addLigneLibre = (n = 1) =>
    setLignesLibres((ls) => [
      ...ls,
      ...Array.from({ length: Math.max(1, Math.round(n)) }, () => ({
        id: uid(),
        famille: "batteries",
        quantite: 1,
        niveauTechnicien: "technicien",
        typeJournee: "semaine",
        prixAchatUnitaire: 20,
        libelleManuel: "",
        joursManuel: 1,
        prixJour: tarifs.technicien.jour,
      })),
    ]);
  const updateLigneLibre = (id, patch) => setLignesLibres((ls) => ls.map((l) => (l.id === id ? { ...l, ...patch } : l)));
  const removeLigneLibre = (id) => setLignesLibres((ls) => ls.filter((l) => l.id !== id));

  function computeLigneLibre(l) {
    const famille = FAMILLES_LIBRES.find((f) => f.id === l.famille);
    const majoration = majorations[l.typeJournee];
    let montant = 0;
    let joursHomme = 0;
    let detail = "";
    if (famille.type === "coef") {
      const cat = catalogueCoef[l.famille];
      const categorie = findCoefCategory(cat, l.prixAchatUnitaire);
      montant = l.prixAchatUnitaire * categorie.coef * l.quantite;
      detail = `${euros(l.prixAchatUnitaire)} × coef ${categorie.coef} (${categorie.label}) × ${l.quantite}`;
    } else {
      joursHomme = l.joursManuel * l.quantite;
      montant = joursHomme * l.prixJour * majoration.coef;
      detail = `${l.joursManuel} j/u × ${l.quantite} × ${euros(l.prixJour)}/j × ${majoration.coef}`;
    }
    return { montant, joursHomme, detail };
  }

  // ---- Lignes issues de la checklist équipements (quantité > 0 uniquement) ----
  const lignesCatalogue = useMemo(() => {
    const tarif = tarifs[affaire.niveauTechnicien];
    const majoration = majorations[affaire.typeJournee];
    const out = [];
    Object.entries(catalogueTemps).forEach(([famId, cat]) => {
      cat.items.forEach((item) => {
        const qte = quantites[item.id] || 0;
        if (qte > 0) {
          const joursHomme = (item.heures / heuresJour) * qte;
          const montantUnitaire = (item.heures / heuresJour) * tarif.jour * majoration.coef + item.amort;
          out.push({ id: item.id, famille: famId, label: item.label, qte, joursHomme, montant: montantUnitaire * qte });
        }
      });
    });
    Object.entries(catalogueDirect).forEach(([famId, cat]) => {
      cat.items.forEach((item) => {
        const qte = quantites[item.id] || 0;
        if (qte > 0) {
          out.push({ id: item.id, famille: famId, label: item.label, qte, joursHomme: 0, montant: item.prix * qte });
        }
      });
    });
    return out;
  }, [catalogueTemps, catalogueDirect, quantites, tarifs, majorations, affaire.niveauTechnicien, affaire.typeJournee, heuresJour]);

  const lignesLibresCalc = useMemo(() => lignesLibres.map((l) => ({ ligne: l, ...computeLigneLibre(l) })), [
    lignesLibres,
    tarifs,
    majorations,
    catalogueCoef,
  ]);

  // Compte les équipements physiques (hors analyses d'huile et mètres TGBT) pour la dégressivité
  const totalEquipements =
    lignesCatalogue.reduce((s, l) => s + (l.famille === "analyseHuile" || l.famille === "tgbt" ? 0 : l.qte), 0) +
    lignesLibres.reduce((s, l) => s + (Number(l.quantite) || 0), 0);

  const totalJoursHomme =
    lignesCatalogue.reduce((s, l) => s + l.joursHomme, 0) + lignesLibresCalc.reduce((s, l) => s + l.joursHomme, 0);

  const totalAvantCoef =
    lignesCatalogue.reduce((s, l) => s + l.montant, 0) + lignesLibresCalc.reduce((s, l) => s + l.montant, 0);

  const palierDegressif =
    degressivite.find((d) => totalEquipements > d.min && totalEquipements <= d.max) || degressivite[degressivite.length - 1];
  const coefContratActif = coefContrat[affaire.contrat] || { coef: 1, label: "—" };

  const totalApresDegressivite = totalAvantCoef * palierDegressif.coef;
  const totalFinal = totalApresDegressivite * coefContratActif.coef;

  const parFamille = useMemo(() => {
    const map = {};
    lignesCatalogue.forEach((l) => (map[l.famille] = (map[l.famille] || 0) + l.montant));
    lignesLibresCalc.forEach((l) => (map[l.ligne.famille] = (map[l.ligne.famille] || 0) + l.montant));
    return map;
  }, [lignesCatalogue, lignesLibresCalc]);

  const nbEquipementsSaisis = Object.values(quantites).filter((v) => v > 0).length + lignesLibres.length;

  // ---- Export Word (.doc) ----
  // Génère un fichier HTML avec en-têtes Word, reconnu et pleinement éditable
  // par Microsoft Word — aucune dépendance externe nécessaire.
  function exportWord() {
    const dateStr = new Date().toLocaleDateString("fr-FR");
    const ligneRows = [
      ...lignesCatalogue.map(
        (l) => `<tr><td>${l.label}</td><td>${LABEL_FAMILLE[l.famille] || l.famille}</td><td style="text-align:right;">${l.qte}</td><td style="text-align:right;">${euros(l.montant)}</td></tr>`
      ),
      ...lignesLibresCalc.map(
        ({ ligne: l, montant }) =>
          `<tr><td>${l.famille === "manuel" ? l.libelleManuel || "Poste libre" : LABEL_FAMILLE[l.famille]}</td><td>${LABEL_FAMILLE[l.famille] || l.famille}</td><td style="text-align:right;">${l.quantite}</td><td style="text-align:right;">${euros(montant)}</td></tr>`
      ),
    ].join("");

    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>Offre ${affaire.reference}</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; color:#1B2733; font-size: 12pt; }
  h1 { color:#1B2733; font-size: 20pt; border-bottom: 3px solid #E8A33D; padding-bottom: 8px; }
  h2 { color:#1B2733; font-size: 14pt; margin-top: 28px; }
  table { border-collapse: collapse; width:100%; margin-bottom:16px; }
  th, td { border:1px solid #DCE0E3; padding:6px 10px; font-size:10.5pt; text-align:left; }
  th { background:#1B2733; color:#fff; }
  .infos td { border:none; padding:2px 0; }
  .total-row td { font-weight:bold; font-size:13pt; border-top:2px solid #1B2733; }
</style></head>
<body>
  <h1>Offre de chiffrage — HT Maintenance</h1>
  <table class="infos">
    <tr><td><b>Référence</b></td><td>${affaire.reference || "—"}</td></tr>
    <tr><td><b>Client</b></td><td>${affaire.client || "—"}</td></tr>
    <tr><td><b>Site</b></td><td>${affaire.site || "—"}</td></tr>
    <tr><td><b>Date</b></td><td>${dateStr}</td></tr>
    <tr><td><b>Cadre contractuel</b></td><td>${coefContratActif.label}</td></tr>
  </table>

  <h2>Détail des équipements et prestations</h2>
  <table>
    <tr><th>Désignation</th><th>Famille</th><th>Qté</th><th>Montant HT</th></tr>
    ${ligneRows || '<tr><td colspan="4">Aucun équipement renseigné</td></tr>'}
  </table>

  <h2>Récapitulatif financier</h2>
  <table>
    <tr><td>Nombre total d'équipements</td><td style="text-align:right;">${totalEquipements}</td></tr>
    <tr><td>Total jours-hommes</td><td style="text-align:right;">${totalJoursHomme.toFixed(2)} j</td></tr>
    <tr><td>Montant HT avant coefficients</td><td style="text-align:right;">${euros(totalAvantCoef)}</td></tr>
    <tr><td>Dégressivité volume (${palierDegressif.label}, ×${palierDegressif.coef})</td><td style="text-align:right;">${euros(totalApresDegressivite)}</td></tr>
    <tr><td>Coefficient contractuel — ${coefContratActif.label} (×${coefContratActif.coef})</td><td style="text-align:right;">${euros(totalFinal)}</td></tr>
    <tr class="total-row"><td>TOTAL HT OFFRE</td><td style="text-align:right;">${euros(totalFinal)}</td></tr>
  </table>
</body>
</html>`;

    const blob = new Blob(["\ufeff", html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Offre_${(affaire.reference || "HT-Maintenance").replace(/[^a-zA-Z0-9_-]/g, "_")}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  const tabBtn = (id, label, Icon) => (
    <button
      onClick={() => setTab(id)}
      className="flex items-center gap-2 px-4 py-2.5 text-sm transition-colors"
      style={{
        color: tab === id ? INK : MUTED,
        fontWeight: tab === id ? 600 : 500,
        borderBottom: tab === id ? `2px solid ${AMBER}` : "2px solid transparent",
      }}
    >
      <Icon size={15} />
      {label}
    </button>
  );

  return (
    <div style={{ background: PAPER, minHeight: "100%", fontFamily: "Inter, system-ui, sans-serif" }} className="w-full">
      {/* Header */}
      <div style={{ background: INK }} className="px-6 py-5">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div style={{ background: AMBER, width: 34, height: 34, borderRadius: 8 }} className="flex items-center justify-center">
              <Zap size={18} color={INK} strokeWidth={2.5} />
            </div>
            <div>
              <div style={{ color: "#fff", fontWeight: 700, fontSize: 16, letterSpacing: 0.2 }}>HT Maintenance</div>
              <div style={{ color: "#9AA6B2", fontSize: 12 }}>Outil de chiffrage — HTA / BT</div>
            </div>
          </div>
          <div style={{ color: "#9AA6B2", fontSize: 12 }}>{affaire.reference}</div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
        <div className="max-w-6xl mx-auto flex px-6">
          {tabBtn("chiffrage", "Chiffrage", ClipboardList)}
          {tabBtn("recap", "Récapitulatif", FileText)}
          {tabBtn("parametres", "Paramètres & catalogue", Settings2)}
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-6 py-6 flex flex-col gap-5">
        {/* ---------------- ONGLET CHIFFRAGE ---------------- */}
        {tab === "chiffrage" && (
          <>
            <SectionCard title="Informations de l'affaire" icon={FileText}>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Référence</label>
                  <TextField value={affaire.reference} onChange={(v) => setAffaire((a) => ({ ...a, reference: v }))} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Client</label>
                  <TextField value={affaire.client} onChange={(v) => setAffaire((a) => ({ ...a, client: v }))} placeholder="Nom du client" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Site</label>
                  <TextField value={affaire.site} onChange={(v) => setAffaire((a) => ({ ...a, site: v }))} placeholder="Site / adresse" />
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Cadre contractuel</label>
                  <Select
                    value={affaire.contrat}
                    onChange={(v) => setAffaire((a) => ({ ...a, contrat: v }))}
                    options={Object.entries(coefContrat).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                  />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Technicien (appliqué aux équipements)</label>
                  <Select
                    value={affaire.niveauTechnicien}
                    onChange={(v) => setAffaire((a) => ({ ...a, niveauTechnicien: v }))}
                    options={Object.entries(tarifs).map(([k, v]) => ({ value: k, label: `${v.label} — ${euros(v.jour)}/j` }))}
                  />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Journée</label>
                  <Select
                    value={affaire.typeJournee}
                    onChange={(v) => setAffaire((a) => ({ ...a, typeJournee: v }))}
                    options={Object.entries(majorations).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                  />
                </div>
              </div>
            </SectionCard>

            <SectionCard
              title="Équipements du site"
              subtitle="Renseignez une quantité pour autant de types d'équipements que nécessaire — pas de limite"
              icon={ClipboardList}
              right={
                <span style={{ fontSize: 12, color: MUTED }}>
                  {nbEquipementsSaisis} type{nbEquipementsSaisis > 1 ? "s" : ""} renseigné{nbEquipementsSaisis > 1 ? "s" : ""}
                </span>
              }
            >
              <div className="flex flex-col gap-6">
                {Object.entries(catalogueTemps).map(([famId, cat]) => (
                  <div key={famId}>
                    <div style={{ fontWeight: 600, color: INK_2, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>
                      {cat.label}
                    </div>
                    <table className="w-full" style={{ fontSize: 13 }}>
                      <tbody>
                        {cat.items.map((item) => {
                          const qte = quantites[item.id] || 0;
                          return (
                            <tr key={item.id} style={{ borderBottom: `1px solid ${LINE}`, background: qte > 0 ? "#FBF3E4" : "transparent" }}>
                              <td className="py-2 pr-3" style={{ color: INK }}>
                                {item.label}
                              </td>
                              <td className="py-2 text-right" style={{ width: 90 }}>
                                <NumberField value={qte} onChange={(v) => setQte(item.id, v)} suffix="u" width={64} />
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                ))}

                {Object.entries(catalogueDirect).map(([famId, cat]) => (
                  <div key={famId}>
                    <div style={{ fontWeight: 600, color: INK_2, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>
                      {cat.label}
                    </div>
                    <table className="w-full" style={{ fontSize: 13 }}>
                      <thead>
                        <tr style={{ color: MUTED, textAlign: "left" }}>
                          <th className="pb-1.5 font-medium">Désignation</th>
                          <th className="pb-1.5 font-medium text-right">Prix / {cat.unite}</th>
                          <th className="pb-1.5 font-medium text-right">Qté</th>
                        </tr>
                      </thead>
                      <tbody>
                        {cat.items.map((item) => {
                          const qte = quantites[item.id] || 0;
                          return (
                            <tr key={item.id} style={{ borderTop: `1px solid ${LINE}`, background: qte > 0 ? "#FBF3E4" : "transparent" }}>
                              <td className="py-2 pr-3" style={{ color: INK }}>
                                {item.label}
                              </td>
                              <td className="py-2 text-right" style={{ color: MUTED, fontVariantNumeric: "tabular-nums" }}>
                                {euros(item.prix)}
                              </td>
                              <td className="py-2 text-right" style={{ width: 90 }}>
                                <NumberField value={qte} onChange={(v) => setQte(item.id, v)} suffix={cat.unite} width={64} />
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard
              title={`Lignes libres (${lignesLibres.length})`}
              subtitle="Batteries, composants, ou tout poste hors catalogue"
              icon={ClipboardList}
              right={
                <div className="flex items-center gap-2">
                  <NumberField value={nbAAjouter} onChange={(v) => setNbAAjouter(Math.max(1, Math.round(v)))} width={50} />
                  <button
                    onClick={() => addLigneLibre(nbAAjouter)}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium"
                    style={{ background: AMBER, color: INK }}
                  >
                    <Plus size={15} /> Ajouter
                  </button>
                </div>
              }
            >
              <div className="flex flex-col gap-3">
                {lignesLibresCalc.map(({ ligne: l, montant, detail }) => {
                  const famille = FAMILLES_LIBRES.find((f) => f.id === l.famille);
                  return (
                    <div key={l.id} style={{ border: `1px solid ${LINE}`, borderRadius: 8 }} className="p-4">
                      <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
                        <div className="md:col-span-2">
                          <label style={{ fontSize: 11, color: MUTED }}>Type</label>
                          <Select value={l.famille} onChange={(v) => updateLigneLibre(l.id, { famille: v })} options={FAMILLES_LIBRES.map((f) => ({ value: f.id, label: f.label }))} />
                        </div>

                        {famille.type === "coef" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: MUTED }}>{catalogueCoef[l.famille].critereLabel}</label>
                            <NumberField value={l.prixAchatUnitaire} onChange={(v) => updateLigneLibre(l.id, { prixAchatUnitaire: v })} width="100%" />
                          </div>
                        )}

                        {famille.type === "manuel" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: MUTED }}>Libellé</label>
                            <TextField value={l.libelleManuel} onChange={(v) => updateLigneLibre(l.id, { libelleManuel: v })} placeholder="Description du poste" />
                          </div>
                        )}

                        <div className="md:col-span-1">
                          <label style={{ fontSize: 11, color: MUTED }}>Qté</label>
                          <NumberField value={l.quantite} onChange={(v) => updateLigneLibre(l.id, { quantite: v })} width="100%" />
                        </div>

                        {famille.type === "manuel" && (
                          <div className="md:col-span-1">
                            <label style={{ fontSize: 11, color: MUTED }}>j/u</label>
                            <NumberField value={l.joursManuel} onChange={(v) => updateLigneLibre(l.id, { joursManuel: v })} width="100%" />
                          </div>
                        )}

                        {famille.type === "manuel" && (
                          <>
                            <div className="md:col-span-2">
                              <label style={{ fontSize: 11, color: MUTED }}>Technicien (préremplit le prix)</label>
                              <Select
                                value={l.niveauTechnicien}
                                onChange={(v) => updateLigneLibre(l.id, { niveauTechnicien: v, prixJour: tarifs[v].jour })}
                                options={Object.entries(tarifs).map(([k, v]) => ({ value: k, label: v.label }))}
                              />
                            </div>
                            <div className="md:col-span-1">
                              <label style={{ fontSize: 11, color: MUTED }}>Prix / jour</label>
                              <NumberField value={l.prixJour} onChange={(v) => updateLigneLibre(l.id, { prixJour: v })} suffix="€" width="100%" />
                            </div>
                            <div className="md:col-span-1">
                              <label style={{ fontSize: 11, color: MUTED }}>Journée</label>
                              <Select value={l.typeJournee} onChange={(v) => updateLigneLibre(l.id, { typeJournee: v })} options={Object.entries(majorations).map(([k, v]) => ({ value: k, label: v.label }))} />
                            </div>
                          </>
                        )}

                        <div className="md:col-span-1 flex justify-end">
                          <button onClick={() => removeLigneLibre(l.id)} style={{ color: "#B0473E" }}>
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </div>
                      <div className="flex items-center justify-between mt-3 pt-3" style={{ borderTop: `1px dashed ${LINE}` }}>
                        <span style={{ fontSize: 12, color: MUTED }}>{detail}</span>
                        <span style={{ fontSize: 15, fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{euros(montant)}</span>
                      </div>
                    </div>
                  );
                })}
                {lignesLibres.length === 0 && (
                  <div style={{ color: MUTED, fontSize: 13 }} className="text-center py-4">
                    Aucune ligne libre. Utile pour les batteries, composants ou prestations hors catalogue.
                  </div>
                )}
              </div>
            </SectionCard>
          </>
        )}

        {/* ---------------- ONGLET RECAP ---------------- */}
        {tab === "recap" && (
          <>
            <SectionCard
              title="Résumé de l'affaire"
              icon={FileText}
              right={
                <button
                  onClick={exportWord}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium"
                  style={{ background: AMBER, color: INK }}
                >
                  <Download size={15} /> Exporter en Word
                </button>
              }
            >
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-2">
                <div>
                  <div style={{ fontSize: 11, color: MUTED }}>Référence</div>
                  <div style={{ fontWeight: 600, color: INK }}>{affaire.reference || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: MUTED }}>Client</div>
                  <div style={{ fontWeight: 600, color: INK }}>{affaire.client || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: MUTED }}>Site</div>
                  <div style={{ fontWeight: 600, color: INK }}>{affaire.site || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: MUTED }}>Types d'équipements</div>
                  <div style={{ fontWeight: 600, color: INK }}>{nbEquipementsSaisis}</div>
                </div>
              </div>
            </SectionCard>

            <SectionCard title="Détail par équipement" icon={ClipboardList}>
              <table className="w-full" style={{ fontSize: 13 }}>
                <thead>
                  <tr style={{ color: MUTED, textAlign: "left" }}>
                    <th className="pb-2 font-medium">Équipement</th>
                    <th className="pb-2 font-medium text-right">Qté</th>
                    <th className="pb-2 font-medium text-right">Montant HT</th>
                  </tr>
                </thead>
                <tbody>
                  {lignesCatalogue.map((l) => (
                    <tr key={l.id} style={{ borderTop: `1px solid ${LINE}` }}>
                      <td className="py-1.5" style={{ color: INK }}>
                        {l.label} <span style={{ color: MUTED, fontSize: 11.5 }}>({LABEL_FAMILLE[l.famille]})</span>
                      </td>
                      <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                        {l.qte}
                      </td>
                      <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                        {euros(l.montant)}
                      </td>
                    </tr>
                  ))}
                  {lignesLibresCalc.map(({ ligne: l, montant }) => (
                    <tr key={l.id} style={{ borderTop: `1px solid ${LINE}` }}>
                      <td className="py-1.5" style={{ color: INK }}>
                        {l.famille === "manuel" ? l.libelleManuel || "Poste libre" : LABEL_FAMILLE[l.famille]}
                      </td>
                      <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                        {l.quantite}
                      </td>
                      <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                        {euros(montant)}
                      </td>
                    </tr>
                  ))}
                  {lignesCatalogue.length === 0 && lignesLibresCalc.length === 0 && (
                    <tr>
                      <td colSpan={3} className="py-4 text-center" style={{ color: MUTED }}>
                        Aucun équipement renseigné pour l'instant.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </SectionCard>

            <SectionCard title="Répartition par famille" icon={ClipboardList}>
              <table className="w-full" style={{ fontSize: 13 }}>
                <thead>
                  <tr style={{ color: MUTED, textAlign: "left" }}>
                    <th className="pb-2 font-medium">Famille</th>
                    <th className="pb-2 font-medium text-right">Montant HT</th>
                    <th className="pb-2 font-medium text-right">Part</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(parFamille).map(([key, montant]) => (
                    <tr key={key} style={{ borderTop: `1px solid ${LINE}` }}>
                      <td className="py-2" style={{ color: INK }}>
                        {LABEL_FAMILLE[key] || key}
                      </td>
                      <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                        {euros(montant)}
                      </td>
                      <td className="py-2 text-right" style={{ color: MUTED }}>
                        {totalAvantCoef > 0 ? Math.round((montant / totalAvantCoef) * 100) : 0}%
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </SectionCard>

            <SectionCard title="Récapitulatif financier" icon={FileText}>
              <table className="w-full" style={{ fontSize: 13.5 }}>
                <tbody>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2" style={{ color: MUTED }}>Nombre total d'équipements</td>
                    <td className="py-2 text-right" style={{ fontWeight: 500, color: INK, fontVariantNumeric: "tabular-nums" }}>{totalEquipements}</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2" style={{ color: MUTED }}>Total jours-hommes</td>
                    <td className="py-2 text-right" style={{ fontWeight: 500, color: INK, fontVariantNumeric: "tabular-nums" }}>{totalJoursHomme.toFixed(2)} j</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2" style={{ color: INK }}>Montant HT avant coefficients</td>
                    <td className="py-2 text-right" style={{ fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{euros(totalAvantCoef)}</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2" style={{ color: INK }}>Dégressivité volume ({palierDegressif.label}, ×{palierDegressif.coef})</td>
                    <td className="py-2 text-right" style={{ fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{euros(totalApresDegressivite)}</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2" style={{ color: INK }}>Coefficient contractuel — {coefContratActif.label} (×{coefContratActif.coef})</td>
                    <td className="py-2 text-right" style={{ fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{euros(totalFinal)}</td>
                  </tr>
                  <tr style={{ borderTop: `2px solid ${INK}` }}>
                    <td className="pt-3" style={{ fontWeight: 700, color: INK, fontSize: 16 }}>TOTAL HT OFFRE</td>
                    <td className="pt-3 text-right" style={{ fontWeight: 800, color: INK, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>{euros(totalFinal)}</td>
                  </tr>
                </tbody>
              </table>
            </SectionCard>
          </>
        )}

        {/* ---------------- ONGLET PARAMETRES ---------------- */}
        {tab === "parametres" && (
          <>
            <SectionCard title="Prix jour technicien" icon={Settings2}>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {Object.entries(tarifs).map(([key, t]) => (
                  <div key={key} style={{ border: `1px solid ${LINE}`, borderRadius: 8 }} className="p-4">
                    <div style={{ fontWeight: 600, color: INK, marginBottom: 10 }}>{t.label}</div>
                    <div className="flex flex-col gap-2">
                      <label style={{ fontSize: 11, color: MUTED }}>Prix / jour</label>
                      <NumberField value={t.jour} onChange={(v) => setTarifs((s) => ({ ...s, [key]: { ...s[key], jour: v } }))} suffix="€" width="100%" />
                      <label style={{ fontSize: 11, color: MUTED }}>Prix / demi-journée</label>
                      <NumberField value={t.demiJour} onChange={(v) => setTarifs((s) => ({ ...s, [key]: { ...s[key], demiJour: v } }))} suffix="€" width="100%" />
                    </div>
                  </div>
                ))}
              </div>
              <div className="mt-4 flex items-center gap-2">
                <label style={{ fontSize: 12, color: MUTED }}>Heures travaillées par jour</label>
                <NumberField value={heuresJour} onChange={setHeuresJour} suffix="h" />
                <span style={{ fontSize: 11.5, color: MUTED }}>
                  (avec les valeurs par défaut : {euros(tarifs.technicien.jour / heuresJour)}/h — cohérent avec le fichier source)
                </span>
              </div>
            </SectionCard>

            <SectionCard title="Coefficients de majoration (jour d'intervention)" icon={Settings2}>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                {Object.entries(majorations).map(([key, m]) => (
                  <div key={key}>
                    <label style={{ fontSize: 11, color: MUTED }}>{m.label}</label>
                    <NumberField value={m.coef} onChange={(v) => setMajorations((s) => ({ ...s, [key]: { ...s[key], coef: v } }))} suffix="×" width="100%" />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Dégressivité selon le volume d'équipements" icon={Settings2}>
              <div className="flex flex-col gap-2">
                {degressivite.map((d, i) => (
                  <div key={i} className="grid grid-cols-3 gap-3 items-center">
                    <span style={{ fontSize: 13, color: INK }}>{d.label}</span>
                    <NumberField value={d.coef} onChange={(v) => setDegressivite((arr) => arr.map((x, j) => (j === i ? { ...x, coef: v } : x)))} suffix="×" />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Coefficients contractuels" icon={Settings2}>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {Object.entries(coefContrat).map(([key, c]) => (
                  <div key={key} className="flex items-center justify-between gap-3">
                    <span style={{ fontSize: 13, color: INK }}>{c.label}</span>
                    <NumberField value={c.coef} onChange={(v) => setCoefContrat((s) => ({ ...s, [key]: { ...s[key], coef: v } }))} suffix="×" />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Catalogue équipements — heures et amortissement" subtitle="Prix de revient = (heures ÷ heures/jour) × tarif jour × majoration + amortissement" icon={ClipboardList}>
              <div className="flex flex-col gap-5">
                {Object.entries(catalogueTemps).map(([key, cat]) => (
                  <div key={key}>
                    <div style={{ fontWeight: 600, color: INK, fontSize: 13 }}>{cat.label}</div>
                    {cat.marquesRef && <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 8 }}>{cat.marquesRef}</div>}
                    <div className="flex flex-col gap-1.5 mt-2">
                      {cat.items.map((item, i) => (
                        <div key={item.id} className="grid grid-cols-4 gap-3 items-center">
                          <span style={{ fontSize: 13, color: INK, gridColumn: "span 2" }}>{item.label}</span>
                          <NumberField
                            value={item.heures}
                            onChange={(v) =>
                              setCatalogueTemps((s) => ({ ...s, [key]: { ...s[key], items: s[key].items.map((it, j) => (j === i ? { ...it, heures: v } : it)) } }))
                            }
                            suffix="h/u"
                          />
                          <NumberField
                            value={item.amort}
                            onChange={(v) =>
                              setCatalogueTemps((s) => ({ ...s, [key]: { ...s[key], items: s[key].items.map((it, j) => (j === i ? { ...it, amort: v } : it)) } }))
                            }
                            suffix="€ amort."
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Catalogue — batteries & composants (coefficient sur prix d'achat)" icon={ClipboardList}>
              <div className="flex flex-col gap-5">
                {Object.entries(catalogueCoef).map(([key, cat]) => (
                  <div key={key}>
                    <div style={{ fontWeight: 600, color: INK, fontSize: 13, marginBottom: 8 }}>{cat.label}</div>
                    <div className="flex flex-col gap-1.5">
                      {cat.categories.map((c, i) => (
                        <div key={c.id} className="grid grid-cols-3 gap-3 items-center">
                          <span style={{ fontSize: 13, color: INK, gridColumn: "span 2" }}>{c.label}</span>
                          <NumberField
                            value={c.coef}
                            onChange={(v) => setCatalogueCoef((s) => ({ ...s, [key]: { ...s[key], categories: s[key].categories.map((x, j) => (j === i ? { ...x, coef: v } : x)) } }))}
                            suffix="×"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Analyses d'huile & TGBT (prix directs)" icon={ClipboardList}>
              <div className="flex flex-col gap-5">
                {Object.entries(catalogueDirect).map(([key, cat]) => (
                  <div key={key}>
                    <div style={{ fontWeight: 600, color: INK, fontSize: 13, marginBottom: 8 }}>{cat.label}</div>
                    <div className="flex flex-col gap-1.5">
                      {cat.items.map((item, i) => (
                        <div key={item.id} className="grid grid-cols-3 gap-3 items-center">
                          <span style={{ fontSize: 13, color: INK, gridColumn: "span 2" }}>{item.label}</span>
                          <NumberField
                            value={item.prix}
                            onChange={(v) => setCatalogueDirect((s) => ({ ...s, [key]: { ...s[key], items: s[key].items.map((it, j) => (j === i ? { ...it, prix: v } : it)) } }))}
                            suffix="€"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </SectionCard>
          </>
        )}
      </div>
    </div>
  );
}

