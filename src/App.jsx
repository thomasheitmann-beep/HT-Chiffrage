import React, { useState, useEffect, useMemo } from "react";
import { Plus, Trash2, Settings2, FileText, ClipboardList, Zap, Download, Copy } from "lucide-react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, WidthType, ShadingType } from "docx";
import { doc, getDoc, setDoc, collection, getDocs, addDoc, deleteDoc } from "firebase/firestore";
import { signInAnonymously, onAuthStateChanged } from "firebase/auth";
import { db, auth } from "./firebase";

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
const STORAGE_KEY = "ht-chiffrage-maintenance-v1";
const FIRESTORE_DOC = "ht-chiffrage/etat";
const DEVIS_COLLECTION = "ht-chiffrage-devis";

// ---------------------------------------------------------------------------
// Catalogue repris du tableau réel "Niv 1-4 (standard)" du fichier
// QUO-CHIFFRAGE HT-BT d'origine. Pour chaque équipement : heures totales
// (niv 1-2 + niv 3-4 + préparation) et amortissement matériel (€, coût fixe
// indépendant du tarif jour). Prix de revient = (heures ÷ heures/jour) ×
// prix jour technicien × majoration + amortissement. Avec les valeurs par
// défaut (1 400 €/j technicien, 7h/jour = 200 €/h), on retombe exactement
// sur les prix du fichier source.
// ---------------------------------------------------------------------------

const DEFAULT_TARIFS = {
  expert: { label: "Expert HT/BT", jour: 1600 },
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
// Pour chaque équipement : heuresSimple (niveau 1-2), heuresComplexe (niveau
// 3-4) et heuresPrepa (préparation, comptée dans tous les niveaux), repris
// du fichier source. niveauDefaut = niveau de prestation appliqué par défaut
// dans un chiffrage (modifiable équipement par équipement dans le Chiffrage).
const DEFAULT_CATALOGUE_TEMPS = {
  hta: {
    label: "Cellules HTA",
    marquesRef: "ABB, Schneider Electric (SM6, FLUOKIT, PIX...), Siemens, CEM Gardy, Ormazabal, Pommier",
    items: [
      { id: "hta-interrupteur", label: "Interrupteur", heuresSimple: 0.7, heuresComplexe: 0.45, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "hta-interrupteur-rh", label: "Interrupteur avec relais homopolaire", heuresSimple: 0.7, heuresComplexe: 1.03, heuresPrepa: 0.2, amort: 21.18, niveauDefaut: "complet" },
      { id: "hta-comptage", label: "Comptage", heuresSimple: 0.6, heuresComplexe: 0.72, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "hta-disj-630", label: "Disjoncteur ≤630A (standard)", heuresSimple: 0.93, heuresComplexe: 3.74, heuresPrepa: 0.5, amort: 57.18, niveauDefaut: "complet" },
      { id: "hta-disj-800", label: "Disjoncteur ≥800A (gros calibres)", heuresSimple: 1.68, heuresComplexe: 3.92, heuresPrepa: 0.5, amort: 68.57, niveauDefaut: "complet" },
      { id: "hta-inter-fusible", label: "Inter fusible", heuresSimple: 0.6, heuresComplexe: 0.72, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "hta-inter-fusible-rh", label: "Inter fusible avec relais homopolaire", heuresSimple: 0.6, heuresComplexe: 1.35, heuresPrepa: 0.2, amort: 23.88, niveauDefaut: "complet" },
      { id: "hta-contacteur", label: "Contacteur", heuresSimple: 0.9, heuresComplexe: 1.9, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "hta-contacteur-rp", label: "Contacteur avec relais de protection", heuresSimple: 0.9, heuresComplexe: 2.6, heuresPrepa: 0.5, amort: 42.86, niveauDefaut: "complet" },
    ],
  },
  transfo: {
    label: "Transformateurs",
    marquesRef: "",
    items: [
      { id: "tr-huile-2000", label: "Transfo huile ≤ 2000 kVA", heuresSimple: 2.58, heuresComplexe: 1.72, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "tr-huile-plus2000", label: "Transfo huile > 2000 kVA", heuresSimple: 3.64, heuresComplexe: 2.56, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "tr-prelevement", label: "Prélèvement seul", heuresSimple: 0, heuresComplexe: 1.0, heuresPrepa: 0, amort: 0, niveauDefaut: "complet" },
      { id: "tr-sec-2000-nc", label: "Transfo sec ≤ 2000 kVA non capoté", heuresSimple: 2.73, heuresComplexe: 0.87, heuresPrepa: 0.1, amort: 58.5, niveauDefaut: "complet" },
      { id: "tr-sec-plus2000-nc", label: "Transfo sec > 2000 kVA non capoté", heuresSimple: 3.64, heuresComplexe: 0.9, heuresPrepa: 0.1, amort: 73.78, niveauDefaut: "complet" },
      { id: "tr-sec-2000-c", label: "Transfo sec ≤ 2000 kVA capoté", heuresSimple: 3.69, heuresComplexe: 1.31, heuresPrepa: 0.1, amort: 81.25, niveauDefaut: "complet" },
      { id: "tr-sec-plus2000-c", label: "Transfo sec > 2000 kVA capoté", heuresSimple: 4.37, heuresComplexe: 1.83, heuresPrepa: 0.1, amort: 100.75, niveauDefaut: "complet" },
    ],
  },
  btSecondaire: {
    label: "Disjoncteurs BT — injection secondaire",
    marquesRef: "Schneider (Masterpact, Compact NS/NSX), Eaton/Moeller, ABB SACE, Siemens, Legrand, GE",
    items: [
      { id: "bts-debro-630", label: "Disj débro ≤630A injection secondaire", heuresSimple: 0.61, heuresComplexe: 0.39, heuresPrepa: 0.2, amort: 25, niveauDefaut: "complet" },
      { id: "bts-debro-3200", label: "Disj débro ≤3200A injection secondaire", heuresSimple: 1.22, heuresComplexe: 0.78, heuresPrepa: 0.2, amort: 50, niveauDefaut: "complet" },
      { id: "bts-debro-4000", label: "Disj débro ≥4000A injection secondaire", heuresSimple: 2.02, heuresComplexe: 0.78, heuresPrepa: 0.2, amort: 70, niveauDefaut: "complet" },
      { id: "bts-fixe-630", label: "Disj fixe ≤630A injection secondaire", heuresSimple: 0.26, heuresComplexe: 0.39, heuresPrepa: 0.2, amort: 16.25, niveauDefaut: "complet" },
      { id: "bts-fixe-3200", label: "Disj fixe ≤3200A injection secondaire", heuresSimple: 0.62, heuresComplexe: 0.78, heuresPrepa: 0.2, amort: 35, niveauDefaut: "complet" },
      { id: "bts-fixe-4000", label: "Disj fixe ≥4000A injection secondaire", heuresSimple: 0.97, heuresComplexe: 0.78, heuresPrepa: 0.2, amort: 43.75, niveauDefaut: "complet" },
      { id: "bts-inter-debro-630", label: "Interrupteur débro ≤630A", heuresSimple: 0.61, heuresComplexe: 0, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "bts-inter-debro-plus630", label: "Interrupteur débro >630A", heuresSimple: 1.22, heuresComplexe: 0, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "bts-inter-fixe-630", label: "Interrupteur fixe ≤630A", heuresSimple: 0.26, heuresComplexe: 0, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "bts-inter-fixe-plus630", label: "Interrupteur fixe >630A", heuresSimple: 0.62, heuresComplexe: 0, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
      { id: "bts-tiroir-inj", label: "Tiroir injection secondaire", heuresSimple: 1.01, heuresComplexe: 0.39, heuresPrepa: 0.2, amort: 35, niveauDefaut: "complet" },
      { id: "bts-tiroir-sans", label: "Tiroir sans injection", heuresSimple: 1.01, heuresComplexe: 0, heuresPrepa: 0.1, amort: 0, niveauDefaut: "complet" },
    ],
  },
  btPrimaire: {
    label: "Disjoncteurs BT — injection primaire",
    marquesRef: "",
    items: [
      { id: "btp-debro-630", label: "Disj débro ≤630A injection primaire", heuresSimple: 0.61, heuresComplexe: 2.0, heuresPrepa: 0.25, amort: 112.36, niveauDefaut: "complet" },
      { id: "btp-debro-3200", label: "Disj débro ≤3200A injection primaire", heuresSimple: 1.22, heuresComplexe: 2.3, heuresPrepa: 0.25, amort: 151.54, niveauDefaut: "complet" },
      { id: "btp-debro-4000", label: "Disj débro ≥4000A injection primaire", heuresSimple: 2.02, heuresComplexe: 4.3, heuresPrepa: 0.25, amort: 272.09, niveauDefaut: "complet" },
      { id: "btp-magneto-1250", label: "Disj magnéto-thermique ≤1250A primaire", heuresSimple: 2.5, heuresComplexe: 3.5, heuresPrepa: 0.5, amort: 258.31, niveauDefaut: "complet" },
      { id: "btp-magneto-plus1250", label: "Disj magnéto-thermique >1250A primaire", heuresSimple: 3.0, heuresComplexe: 4.0, heuresPrepa: 0.5, amort: 301.36, niveauDefaut: "complet" },
    ],
  },
  pfcRec: {
    label: "Divers — compensateurs et redresseurs",
    marquesRef: "",
    items: [
      { id: "pfc-bt-5", label: "PFC BT : 400V ≤ 5 gradins", heuresSimple: 0, heuresComplexe: 0, heuresPrepa: 2.8, amort: 0, niveauDefaut: "complet" },
      { id: "pfc-bt-plus5", label: "PFC BT : 400V > 5 gradins", heuresSimple: 0, heuresComplexe: 0, heuresPrepa: 4.0, amort: 0, niveauDefaut: "complet" },
      { id: "pfc-hta", label: "PFC HTA : 20kV ≤1500kvar / 5.5kV ≤150kvar", heuresSimple: 4.0, heuresComplexe: 4.0, heuresPrepa: 0, amort: 0, niveauDefaut: "complet" },
      { id: "rec-c13", label: "REC type C13-100", heuresSimple: 0, heuresComplexe: 1.56, heuresPrepa: 0, amort: 0, niveauDefaut: "complet" },
      { id: "rec-sces", label: "REC type SCES", heuresSimple: 0, heuresComplexe: 2.8, heuresPrepa: 0, amort: 0, niveauDefaut: "complet" },
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
// et nettoyage TGBT, dont les prix du fichier source sont fixes.
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

// Familles pour les "lignes libres" (batteries, composants, poste manuel)
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

// Niveau de prestation réalisée sur l'équipement (pas le technicien) :
// simple (1-2), complexe (3-4), ou complet (1-4 = simple + complexe).
// La préparation est comptée dans tous les niveaux.
const NIVEAUX_PRESTATION = {
  simple: "Niveau 1-2 (simple)",
  complexe: "Niveau 3-4 (complexe)",
  complet: "Niveau 1-4 (complet)",
};
function heuresPourNiveau(item, niveau) {
  if (niveau === "simple") return item.heuresSimple + item.heuresPrepa;
  if (niveau === "complexe") return item.heuresComplexe + item.heuresPrepa;
  return item.heuresSimple + item.heuresComplexe + item.heuresPrepa; // complet
}

function euros(n) {
  if (!isFinite(n)) return "0 €";
  return n.toLocaleString("fr-FR", { maximumFractionDigits: 0 }) + " €";
}

function findCoefCategory(catalogue, valeur) {
  return catalogue.categories.find((c) => valeur >= c.min && valeur <= c.max) || catalogue.categories[0];
}

function nouveauPosteEquipement(n) {
  return { id: uid(), nom: `Poste ${n}`, niveauTechnicien: "expert", typeJournee: "semaine", quantites: {}, niveaux: {} };
}

// ---------------------------------------------------------------------------

function SectionCard({ title, icon: Icon, children, right, subtitle }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="overflow-hidden">
      <div className="flex flex-wrap items-center justify-between gap-2 px-4 md:px-5 py-4" style={{ borderBottom: `1px solid ${LINE}` }}>
        <div className="flex items-center gap-2">
          {Icon && <Icon size={17} color={INK_2} />}
          <div>
            <h2 style={{ color: INK, fontWeight: 600, fontSize: 15 }}>{title}</h2>
            {subtitle && <div style={{ color: MUTED, fontSize: 11.5 }}>{subtitle}</div>}
          </div>
        </div>
        {right}
      </div>
      <div className="p-4 md:p-5">{children}</div>
    </div>
  );
}

// Champ numérique robuste : garde une saisie texte locale tant que le champ
// a le focus (permet d'effacer, taper "1" puis "0" pour "10", virgule ou
// point) et ne se resynchronise sur la valeur numérique qu'à la perte de
// focus ou si la valeur change depuis l'extérieur.
function NumberField({ value, onChange, suffix, width = 90 }) {
  const [text, setText] = useState(String(value ?? 0));
  const [focused, setFocused] = useState(false);

  useEffect(() => {
    if (!focused) setText(String(value ?? 0));
  }, [value, focused]);

  return (
    <span className="inline-flex items-center gap-1">
      <input
        type="text"
        inputMode="decimal"
        value={text}
        onFocus={() => setFocused(true)}
        onChange={(e) => {
          const raw = e.target.value;
          if (raw === "" || /^-?\d*([.,]\d*)?$/.test(raw)) {
            setText(raw);
            const n = parseFloat(raw.replace(",", "."));
            onChange(isNaN(n) ? 0 : n);
          }
        }}
        onBlur={() => {
          setFocused(false);
          setText(String(value ?? 0));
        }}
        style={{
          width,
          border: `1px solid ${LINE}`,
          borderRadius: 6,
          padding: "5px 8px",
          fontSize: 13,
          fontVariantNumeric: "tabular-nums",
          color: INK,
          background: "#fff",
          colorScheme: "light",
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
      style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: "6px 9px", fontSize: 13, color: INK, background: "#fff", colorScheme: "light", width: "100%", ...style }}
    />
  );
}

function Select({ value, onChange, options, style }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: "6px 9px", fontSize: 13, color: INK, background: "#fff", colorScheme: "light", width: "100%", ...style }}
    >
      {options.map((o) => (
        <option key={o.value} value={o.value}>
          {o.label}
        </option>
      ))}
    </select>
  );
}

// Table réutilisable pour un poste d'équipement (mêmes catalogues pour tous
// les postes) — défini en dehors du composant principal pour ne pas être
// recréé à chaque rendu (sinon React démonte/remonte tout le tableau à
// chaque frappe, ce qui fait perdre le focus et remonter la page).
function TableauCatalogue({ poste, catalogueTemps, catalogueDirect, setQtePoste, setNiveauPoste }) {
  return (
    <div className="flex flex-col gap-6">
      {Object.entries(catalogueTemps).map(([famId, cat]) => (
        <div key={famId}>
          <div style={{ fontWeight: 600, color: INK_2, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>{cat.label}</div>
          <div className="overflow-x-auto">
            <table className="w-full" style={{ fontSize: 13 }}>
              <tbody>
                {cat.items.map((item) => {
                  const qte = poste.quantites[item.id] || 0;
                  const niveau = (poste.niveaux && poste.niveaux[item.id]) || item.niveauDefaut || "complet";
                  return (
                    <tr key={item.id} style={{ borderBottom: `1px solid ${LINE}`, background: qte > 0 ? "#FBF3E4" : "transparent" }}>
                      <td className="py-2 pr-3" style={{ color: INK }}>
                        {item.label}
                      </td>
                      <td className="py-2 pr-3 text-right" style={{ width: 170 }}>
                        {qte > 0 && (
                          <Select
                            value={niveau}
                            onChange={(v) => setNiveauPoste(poste.id, item.id, v)}
                            options={Object.entries(NIVEAUX_PRESTATION).map(([k, label]) => ({ value: k, label }))}
                            style={{ fontSize: 12, padding: "4px 6px" }}
                          />
                        )}
                      </td>
                      <td className="py-2 text-right" style={{ width: 90 }}>
                        <NumberField value={qte} onChange={(v) => setQtePoste(poste.id, item.id, v)} suffix="u" width={64} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      ))}

      {Object.entries(catalogueDirect).map(([famId, cat]) => (
        <div key={famId}>
          <div style={{ fontWeight: 600, color: INK_2, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>{cat.label}</div>
          <div className="overflow-x-auto">
            <table className="w-full" style={{ fontSize: 13 }}>
              <tbody>
                {cat.items.map((item) => {
                  const qte = poste.quantites[item.id] || 0;
                  return (
                    <tr key={item.id} style={{ borderBottom: `1px solid ${LINE}`, background: qte > 0 ? "#FBF3E4" : "transparent" }}>
                      <td className="py-2 pr-3" style={{ color: INK }}>
                        {item.label}
                      </td>
                      <td className="py-2 text-right" style={{ width: 90 }}>
                        <NumberField value={qte} onChange={(v) => setQtePoste(poste.id, item.id, v)} suffix={cat.unite} width={64} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------------

export default function ChiffrageHTMaintenance() {
  const [tab, setTab] = useState("chiffrage");

  // Sauvegarde automatique dans le navigateur : relit ce qui a été enregistré
  // au dernier passage (une seule fois, au tout premier rendu).
  const [saved] = useState(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });

  // Vérifie que les données enregistrées (cache local ou cloud) correspondent
  // au schéma actuel — sinon on repart des valeurs par défaut à jour plutôt
  // que de charger d'anciennes structures incompatibles (ex: anciens tarifs
  // multi-niveaux, ancien champ "heures" unique par équipement).
  const tarifsValides = (t) => !!(t && t.expert && Object.keys(t).length === 1);
  const catalogueTempsValide = (c) => !!(c && c.hta && c.hta.items && c.hta.items[0] && c.hta.items[0].heuresSimple != null);

  const [tarifs, setTarifs] = useState(tarifsValides(saved.tarifs) ? saved.tarifs : DEFAULT_TARIFS);
  const [majorations, setMajorations] = useState(saved.majorations || DEFAULT_MAJORATIONS);
  const [degressivite, setDegressivite] = useState(saved.degressivite || DEFAULT_DEGRESSIVITE);
  const [coefContrat, setCoefContrat] = useState(saved.coefContrat || DEFAULT_COEF_CONTRAT);
  const [heuresJour, setHeuresJour] = useState(saved.heuresJour ?? 7);

  const [catalogueTemps, setCatalogueTemps] = useState(catalogueTempsValide(saved.catalogueTemps) ? saved.catalogueTemps : DEFAULT_CATALOGUE_TEMPS);
  const [catalogueCoef, setCatalogueCoef] = useState(saved.catalogueCoef || DEFAULT_CATALOGUE_COEF);
  const [catalogueDirect, setCatalogueDirect] = useState(saved.catalogueDirect || DEFAULT_CATALOGUE_DIRECT);

  // Contenu vierge d'un chiffrage — les tarifs/catalogues (ci-dessus) restent
  // partagés entre tous les chiffrages, seuls affaire/postes/lignes sont propres
  // à chaque chiffrage.
  function devisVierge() {
    return {
      affaire: {
        client: "",
        site: "",
        reference: "DEV-" + new Date().getFullYear() + "-001",
        contrat: "aucun",
        degressiviteActive: true,
      },
      postesEquipement: [nouveauPosteEquipement(1)],
      lignesLibres: [],
    };
  }

  const [affaire, setAffaire] = useState(saved.affaire || devisVierge().affaire);
  const [postesEquipement, setPostesEquipement] = useState(saved.postesEquipement || devisVierge().postesEquipement);
  const [lignesLibres, setLignesLibres] = useState(saved.lignesLibres || []);
  const [nbAAjouter, setNbAAjouter] = useState(1);

  const addPosteEquipement = () => setPostesEquipement((ps) => [...ps, nouveauPosteEquipement(ps.length + 1)]);
  const dupliquerPosteEquipement = (id) =>
    setPostesEquipement((ps) => {
      const src = ps.find((p) => p.id === id);
      if (!src) return ps;
      return [...ps, { ...src, id: uid(), nom: `${src.nom} (copie)`, quantites: { ...src.quantites }, niveaux: { ...(src.niveaux || {}) } }];
    });
  const updatePosteEquipement = (id, patch) => setPostesEquipement((ps) => ps.map((p) => (p.id === id ? { ...p, ...patch } : p)));
  const removePosteEquipement = (id) => setPostesEquipement((ps) => (ps.length > 1 ? ps.filter((p) => p.id !== id) : ps));
  const setQtePoste = (posteId, itemId, v) =>
    setPostesEquipement((ps) => ps.map((p) => (p.id === posteId ? { ...p, quantites: { ...p.quantites, [itemId]: Math.max(0, v) } } : p)));
  const setNiveauPoste = (posteId, itemId, niveau) =>
    setPostesEquipement((ps) => ps.map((p) => (p.id === posteId ? { ...p, niveaux: { ...(p.niveaux || {}), [itemId]: niveau } } : p)));

  const [syncState, setSyncState] = useState("idle"); // idle | loading | syncing | synced | error
  const [authReady, setAuthReady] = useState(false);

  // Liste des chiffrages (devis) disponibles, et celui actuellement affiché.
  const [devisList, setDevisList] = useState(saved.devisList || []);
  const [currentDevisId, setCurrentDevisId] = useState(saved.currentDevisId || null);
  const [devisReady, setDevisReady] = useState(false); // évite d'écraser un chiffrage avant la fin du chargement initial

  // Connexion anonyme automatique — nécessaire car les règles Firestore de ce
  // projet exigent request.auth != null. Aucune interface de connexion pour
  // l'utilisateur : ça se passe en arrière-plan, comme un accès invité.
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (user) => {
      if (user) {
        setAuthReady(true);
      } else {
        signInAnonymously(auth).catch(() => setSyncState("error"));
      }
    });
    return unsubscribe;
  }, []);

  // Charge les paramètres/catalogues globaux (partagés entre tous les chiffrages)
  useEffect(() => {
    if (!authReady) return;
    setSyncState("loading");
    getDoc(doc(db, FIRESTORE_DOC))
      .then((snap) => {
        if (snap.exists()) {
          const data = snap.data();
          setTarifs(tarifsValides(data.tarifs) ? data.tarifs : DEFAULT_TARIFS);
          if (data.majorations) setMajorations(data.majorations);
          if (data.degressivite) setDegressivite(data.degressivite);
          if (data.coefContrat) setCoefContrat(data.coefContrat);
          if (data.heuresJour != null) setHeuresJour(data.heuresJour);
          setCatalogueTemps(catalogueTempsValide(data.catalogueTemps) ? data.catalogueTemps : DEFAULT_CATALOGUE_TEMPS);
          if (data.catalogueCoef) setCatalogueCoef(data.catalogueCoef);
          if (data.catalogueDirect) setCatalogueDirect(data.catalogueDirect);
        }
        setSyncState("synced");
      })
      .catch(() => setSyncState("error"));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [authReady]);

  // Charge la liste des chiffrages existants, puis affiche le dernier ouvert
  // (ou en crée un premier si la liste est vide).
  useEffect(() => {
    if (!authReady) return;
    getDocs(collection(db, DEVIS_COLLECTION))
      .then(async (snap) => {
        let list = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        list.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
        if (list.length === 0) {
          const blank = devisVierge();
          const payload = { ...blank, updatedAt: Date.now() };
          const ref = await addDoc(collection(db, DEVIS_COLLECTION), payload);
          list = [{ id: ref.id, ...payload }];
        }
        setDevisList(list.map((d) => ({ id: d.id, reference: d.affaire?.reference, client: d.affaire?.client, updatedAt: d.updatedAt })));
        const target = list.find((d) => d.id === currentDevisId) || list[0];
        setCurrentDevisId(target.id);
        setAffaire(target.affaire || devisVierge().affaire);
        setPostesEquipement(target.postesEquipement || devisVierge().postesEquipement);
        setLignesLibres(target.lignesLibres || []);
        setDevisReady(true);
      })
      .catch(() => setSyncState("error"));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [authReady]);

  // Passe à un autre chiffrage existant (sélecteur)
  const chargerDevis = (id) => {
    setDevisReady(false);
    getDoc(doc(db, DEVIS_COLLECTION, id))
      .then((snap) => {
        if (snap.exists()) {
          const data = snap.data();
          setCurrentDevisId(id);
          setAffaire(data.affaire || devisVierge().affaire);
          setPostesEquipement(data.postesEquipement || devisVierge().postesEquipement);
          setLignesLibres(data.lignesLibres || []);
        }
        setDevisReady(true);
      })
      .catch(() => setDevisReady(true));
  };

  const nouveauChiffrage = async () => {
    const blank = devisVierge();
    setDevisReady(false);
    try {
      const payload = { ...blank, updatedAt: Date.now() };
      const ref = await addDoc(collection(db, DEVIS_COLLECTION), payload);
      setDevisList((l) => [{ id: ref.id, reference: blank.affaire.reference, client: "", updatedAt: payload.updatedAt }, ...l]);
      setCurrentDevisId(ref.id);
      setAffaire(blank.affaire);
      setPostesEquipement(blank.postesEquipement);
      setLignesLibres(blank.lignesLibres);
    } finally {
      setDevisReady(true);
    }
  };

  const dupliquerChiffrage = async () => {
    setDevisReady(false);
    try {
      const payload = {
        affaire: { ...affaire, reference: affaire.reference + " (copie)" },
        postesEquipement: postesEquipement.map((p) => ({ ...p, quantites: { ...p.quantites } })),
        lignesLibres: lignesLibres.map((l) => ({ ...l })),
        updatedAt: Date.now(),
      };
      const ref = await addDoc(collection(db, DEVIS_COLLECTION), payload);
      setDevisList((l) => [{ id: ref.id, reference: payload.affaire.reference, client: payload.affaire.client, updatedAt: payload.updatedAt }, ...l]);
      setCurrentDevisId(ref.id);
      setAffaire(payload.affaire);
      setPostesEquipement(payload.postesEquipement);
      setLignesLibres(payload.lignesLibres);
    } finally {
      setDevisReady(true);
    }
  };

  const supprimerChiffrage = async () => {
    if (devisList.length <= 1) return;
    if (!window.confirm(`Supprimer définitivement le chiffrage « ${affaire.reference} » ? Cette action est irréversible.`)) return;
    const idASupprimer = currentDevisId;
    try {
      await deleteDoc(doc(db, DEVIS_COLLECTION, idASupprimer));
    } catch {
      // on continue quand même côté interface
    }
    const reste = devisList.filter((d) => d.id !== idASupprimer);
    setDevisList(reste);
    chargerDevis(reste[0].id);
  };

  // Enregistre automatiquement les paramètres/catalogues (partagés entre
  // chiffrages) dans le cloud, après une courte pause.
  useEffect(() => {
    if (!authReady) return;
    const payload = { tarifs, majorations, degressivite, coefContrat, heuresJour, catalogueTemps, catalogueCoef, catalogueDirect };
    const t = setTimeout(() => {
      setDoc(doc(db, FIRESTORE_DOC), payload)
        .then(() => setSyncState("synced"))
        .catch(() => setSyncState("error"));
    }, 1000);
    return () => clearTimeout(t);
  }, [authReady, tarifs, majorations, degressivite, coefContrat, heuresJour, catalogueTemps, catalogueCoef, catalogueDirect]);

  // Enregistre automatiquement le chiffrage courant (affaire + postes +
  // lignes libres) dans son propre document, sans toucher aux autres.
  useEffect(() => {
    if (!authReady || !devisReady || !currentDevisId) return;
    setSyncState("syncing");
    const t = setTimeout(() => {
      const payload = { affaire, postesEquipement, lignesLibres, updatedAt: Date.now() };
      setDoc(doc(db, DEVIS_COLLECTION, currentDevisId), payload)
        .then(() => {
          setSyncState("synced");
          setDevisList((l) => l.map((d) => (d.id === currentDevisId ? { ...d, reference: affaire.reference, client: affaire.client, updatedAt: payload.updatedAt } : d)));
        })
        .catch(() => setSyncState("error"));
    }, 1000);
    return () => clearTimeout(t);
  }, [authReady, devisReady, currentDevisId, affaire, postesEquipement, lignesLibres]);

  // Cache local (rechargement rapide avant que le cloud ne réponde)
  useEffect(() => {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          tarifs,
          majorations,
          degressivite,
          coefContrat,
          heuresJour,
          catalogueTemps,
          catalogueCoef,
          catalogueDirect,
          affaire,
          postesEquipement,
          lignesLibres,
          devisList,
          currentDevisId,
        })
      );
    } catch {
      // stockage local indisponible (navigation privée, quota dépassé...) — on continue sans bloquer
    }
  }, [tarifs, majorations, degressivite, coefContrat, heuresJour, catalogueTemps, catalogueCoef, catalogueDirect, affaire, postesEquipement, lignesLibres, devisList, currentDevisId]);

  // Enregistre immédiatement les paramètres/catalogues, sans attendre le
  // délai automatique (utile pour une confirmation explicite à l'utilisateur).
  const enregistrerParametresMaintenant = () => {
    setSyncState("syncing");
    setDoc(doc(db, FIRESTORE_DOC), { tarifs, majorations, degressivite, coefContrat, heuresJour, catalogueTemps, catalogueCoef, catalogueDirect })
      .then(() => setSyncState("synced"))
      .catch(() => setSyncState("error"));
  };

  const reinitialiserParametres = () => {
    if (!window.confirm("Réinitialiser tous les paramètres et catalogues aux valeurs par défaut ?")) return;
    setTarifs(DEFAULT_TARIFS);
    setMajorations(DEFAULT_MAJORATIONS);
    setDegressivite(DEFAULT_DEGRESSIVITE);
    setCoefContrat(DEFAULT_COEF_CONTRAT);
    setHeuresJour(7);
    setCatalogueTemps(DEFAULT_CATALOGUE_TEMPS);
    setCatalogueCoef(DEFAULT_CATALOGUE_COEF);
    setCatalogueDirect(DEFAULT_CATALOGUE_DIRECT);
  };

  // RAZ du chiffrage en cours : vide l'affaire, les postes d'équipements et
  // les lignes libres pour repartir sur un devis neuf. Ne touche pas aux
  // tarifs/catalogues de Paramètres.
  const razChiffrage = () => {
    if (!window.confirm("Remettre à zéro le chiffrage en cours (infos affaire, équipements saisis, lignes libres) ? Cette action est irréversible.")) return;
    const blank = devisVierge();
    setAffaire(blank.affaire);
    setPostesEquipement(blank.postesEquipement);
    setLignesLibres(blank.lignesLibres);
  };


  const addLigneLibre = (n = 1) =>
    setLignesLibres((ls) => [
      ...ls,
      ...Array.from({ length: Math.max(1, Math.round(n)) }, () => ({
        id: uid(),
        famille: "batteries",
        quantite: 1,
        niveauTechnicien: "expert",
        typeJournee: "semaine",
        prixAchatUnitaire: 20,
        libelleManuel: "",
        joursManuel: 1,
        prixJour: tarifs.expert.jour,
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

  // ---- Lignes issues des postes d'équipements (quantité > 0 uniquement) ----
  const lignesCatalogue = useMemo(() => {
    const out = [];
    postesEquipement.forEach((poste) => {
      const majoration = majorations[poste.typeJournee];
      Object.entries(catalogueTemps).forEach(([famId, cat]) => {
        cat.items.forEach((item) => {
          const qte = poste.quantites[item.id] || 0;
          if (qte > 0) {
            const tarif = tarifs.expert;
            const niveau = (poste.niveaux && poste.niveaux[item.id]) || item.niveauDefaut || "complet";
            const heures = heuresPourNiveau(item, niveau);
            const joursHomme = (heures / heuresJour) * qte;
            const montantUnitaire = (heures / heuresJour) * tarif.jour * majoration.coef + item.amort;
            out.push({ id: `${poste.id}-${item.id}`, posteId: poste.id, posteNom: poste.nom, famille: famId, label: item.label, qte, joursHomme, montant: montantUnitaire * qte });
          }
        });
      });
      Object.entries(catalogueDirect).forEach(([famId, cat]) => {
        cat.items.forEach((item) => {
          const qte = poste.quantites[item.id] || 0;
          if (qte > 0) {
            out.push({ id: `${poste.id}-${item.id}`, posteId: poste.id, posteNom: poste.nom, famille: famId, label: item.label, qte, joursHomme: 0, montant: item.prix * qte });
          }
        });
      });
    });
    return out;
  }, [postesEquipement, catalogueTemps, catalogueDirect, tarifs, majorations, heuresJour]);

  const lignesLibresCalc = useMemo(() => lignesLibres.map((l) => ({ ligne: l, ...computeLigneLibre(l) })), [
    lignesLibres,
    tarifs,
    majorations,
    catalogueCoef,
  ]);

  // Les lignes libres (batteries, composants, poste manuel) ne comptent pas
  // comme des équipements pour le calcul de la dégressivité — seuls les
  // équipements physiques du catalogue (hors analyses d'huile et TGBT) comptent.
  const totalEquipements = lignesCatalogue.reduce((s, l) => s + (l.famille === "analyseHuile" || l.famille === "tgbt" ? 0 : l.qte), 0);

  const totalJoursHomme =
    lignesCatalogue.reduce((s, l) => s + l.joursHomme, 0) + lignesLibresCalc.reduce((s, l) => s + l.joursHomme, 0);

  const totalAvantCoef =
    lignesCatalogue.reduce((s, l) => s + l.montant, 0) + lignesLibresCalc.reduce((s, l) => s + l.montant, 0);

  const palierDegressif =
    degressivite.find((d) => totalEquipements > d.min && totalEquipements <= d.max) || degressivite[degressivite.length - 1];
  const coefContratActif = coefContrat[affaire.contrat] || { coef: 1, label: "—" };

  const totalApresDegressivite = affaire.degressiviteActive ? totalAvantCoef * palierDegressif.coef : totalAvantCoef;
  const totalFinal = totalApresDegressivite * coefContratActif.coef;

  const parFamille = useMemo(() => {
    const map = {};
    lignesCatalogue.forEach((l) => (map[l.famille] = (map[l.famille] || 0) + l.montant));
    lignesLibresCalc.forEach((l) => (map[l.ligne.famille] = (map[l.ligne.famille] || 0) + l.montant));
    return map;
  }, [lignesCatalogue, lignesLibresCalc]);

  const nbEquipementsSaisis =
    postesEquipement.reduce((s, p) => s + Object.values(p.quantites).filter((v) => v > 0).length, 0) + lignesLibres.length;

  // ---- Export Word (.docx) ----
  // Génère un vrai fichier .docx (Office Open XML), ouvrable par Word, Pages
  // et Google Docs — nécessite `npm install docx` dans le projet.
  async function exportWord() {
    const dateStr = new Date().toLocaleDateString("fr-FR");

    const noBorder = { style: "none", size: 0, color: "FFFFFF" };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

    // Largeur totale utile de la page (twips) : A4, marges par défaut ~1440
    // twips de chaque côté sur une page de 12240 twips → ~9360 twips utiles.
    const PAGE_WIDTH = 9360;
    const colWidths = (percents) => percents.map((p) => Math.round((PAGE_WIDTH * p) / 100));

    const txt = (text, opts = {}) => new TextRun({ text: String(text), bold: !!opts.bold, color: opts.color });
    const para = (text, opts = {}) => new Paragraph({ alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [txt(text, opts)] });

    // --- Tableau infos affaire (2 colonnes) ---
    const infoW = colWidths([30, 70]);
    const infoRows = [
      ["Référence", affaire.reference || "—"],
      ["Client", affaire.client || "—"],
      ["Site", affaire.site || "—"],
      ["Date", dateStr],
      ["Cadre contractuel", coefContratActif.label],
    ].map(
      ([label, value]) =>
        new TableRow({
          children: [
            new TableCell({ width: { size: infoW[0], type: WidthType.DXA }, borders: noBorders, children: [para(label, { bold: true })] }),
            new TableCell({ width: { size: infoW[1], type: WidthType.DXA }, borders: noBorders, children: [para(value)] }),
          ],
        })
    );
    const infoTable = new Table({ width: { size: PAGE_WIDTH, type: WidthType.DXA }, columnWidths: infoW, rows: infoRows });

    // --- Tableau détail équipements (5 colonnes) ---
    const equipW = colWidths([28, 15, 22, 10, 25]);
    const headerCell = (text, i) =>
      new TableCell({ width: { size: equipW[i], type: WidthType.DXA }, shading: { type: ShadingType.SOLID, fill: "1B2733" }, children: [para(text, { bold: true, color: "FFFFFF" })] });
    const dataCell = (text, i, opts = {}) => new TableCell({ width: { size: equipW[i], type: WidthType.DXA }, children: [para(text, opts)] });

    const equipHeader = new TableRow({ children: ["Désignation", "Poste", "Famille", "Qté", "Montant HT"].map(headerCell) });
    const equipDataRows = [
      ...lignesCatalogue.map(
        (l) =>
          new TableRow({
            children: [
              dataCell(l.label, 0),
              dataCell(l.posteNom, 1),
              dataCell(LABEL_FAMILLE[l.famille] || l.famille, 2),
              dataCell(l.qte, 3, { right: true }),
              dataCell(euros(l.montant), 4, { right: true }),
            ],
          })
      ),
      ...lignesLibresCalc.map(
        ({ ligne: l, montant }) =>
          new TableRow({
            children: [
              dataCell(l.famille === "manuel" ? l.libelleManuel || "Poste libre" : LABEL_FAMILLE[l.famille], 0),
              dataCell("—", 1),
              dataCell(LABEL_FAMILLE[l.famille] || l.famille, 2),
              dataCell(l.quantite, 3, { right: true }),
              dataCell(euros(montant), 4, { right: true }),
            ],
          })
      ),
    ];
    if (equipDataRows.length === 0) {
      equipDataRows.push(new TableRow({ children: [new TableCell({ columnSpan: 5, children: [para("Aucun équipement renseigné")] })] }));
    }
    const equipTable = new Table({ width: { size: PAGE_WIDTH, type: WidthType.DXA }, columnWidths: equipW, rows: [equipHeader, ...equipDataRows] });

    // --- Tableau récapitulatif financier (2 colonnes) ---
    const recapW = colWidths([70, 30]);
    const degressiviteLabel = affaire.degressiviteActive ? `Dégressivité volume (${palierDegressif.label}, ×${palierDegressif.coef})` : "Dégressivité volume (désactivée)";
    const recapRows = [
      ["Nombre total d'équipements", totalEquipements, false],
      ["Total jours-hommes", `${totalJoursHomme.toFixed(2)} j`, false],
      ["Montant HT avant coefficients", euros(totalAvantCoef), false],
      [degressiviteLabel, euros(totalApresDegressivite), false],
      [`Coefficient contractuel — ${coefContratActif.label} (×${coefContratActif.coef})`, euros(totalFinal), false],
      ["TOTAL HT OFFRE", euros(totalFinal), true],
    ].map(
      ([label, value, bold]) =>
        new TableRow({
          children: [
            new TableCell({ width: { size: recapW[0], type: WidthType.DXA }, borders: noBorders, children: [para(label, { bold })] }),
            new TableCell({ width: { size: recapW[1], type: WidthType.DXA }, borders: noBorders, children: [para(value, { bold, right: true })] }),
          ],
        })
    );
    const recapTable = new Table({ width: { size: PAGE_WIDTH, type: WidthType.DXA }, columnWidths: recapW, rows: recapRows });

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: "Offre de chiffrage — HT Maintenance", heading: HeadingLevel.HEADING_1 }),
            infoTable,
            new Paragraph({ text: "", spacing: { after: 200 } }),
            new Paragraph({ text: "Détail des équipements et prestations", heading: HeadingLevel.HEADING_2, spacing: { before: 200 } }),
            equipTable,
            new Paragraph({ text: "", spacing: { after: 200 } }),
            new Paragraph({ text: "Récapitulatif financier", heading: HeadingLevel.HEADING_2, spacing: { before: 200 } }),
            recapTable,
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Offre_${(affaire.reference || "HT-Maintenance").replace(/[^a-zA-Z0-9_-]/g, "_")}.docx`;
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
    <div style={{ background: PAPER, minHeight: "100%", fontFamily: "Inter, system-ui, sans-serif", colorScheme: "light" }} className="w-full">
      {/* Header */}
      <div style={{ background: INK }} className="px-4 md:px-6 py-4 md:py-5">
        <div className="max-w-6xl mx-auto flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-3">
            <div style={{ background: AMBER, width: 34, height: 34, borderRadius: 8 }} className="flex items-center justify-center shrink-0">
              <Zap size={18} color={INK} strokeWidth={2.5} />
            </div>
            <div>
              <div style={{ color: "#fff", fontWeight: 700, fontSize: 16, letterSpacing: 0.2 }}>HT Maintenance</div>
              <div style={{ color: "#9AA6B2", fontSize: 12 }} className="hidden sm:block">Outil de chiffrage — HTA / BT</div>
            </div>
          </div>
          <div className="flex items-center gap-3 flex-wrap">
            <span style={{ color: "#9AA6B2", fontSize: 11 }}>
              {syncState === "loading" && "Chargement…"}
              {syncState === "syncing" && "Enregistrement…"}
              {syncState === "synced" && "✓ Synchronisé"}
              {syncState === "error" && "⚠ Hors ligne"}
            </span>
            <div style={{ color: "#9AA6B2", fontSize: 12 }}>{affaire.reference}</div>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
        <div className="max-w-6xl mx-auto flex px-4 md:px-6 overflow-x-auto">
          {tabBtn("chiffrage", "Chiffrage", ClipboardList)}
          {tabBtn("recap", "Récapitulatif", FileText)}
          {tabBtn("parametres", "Paramètres & catalogue", Settings2)}
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-4 md:px-6 py-5 md:py-6 flex flex-col gap-5">
        {/* ---------------- ONGLET CHIFFRAGE ---------------- */}
        {tab === "chiffrage" && (
          <>
            <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="p-4 flex flex-wrap items-center gap-3 justify-between">
              <div className="flex items-center gap-2 flex-wrap">
                <span style={{ fontSize: 11, color: MUTED }}>Chiffrage</span>
                <Select
                  value={currentDevisId || ""}
                  onChange={(id) => chargerDevis(id)}
                  options={devisList.map((d) => ({ value: d.id, label: `${d.reference || "Sans référence"}${d.client ? " — " + d.client : ""}` }))}
                  style={{ minWidth: 240, width: "auto" }}
                />
              </div>
              <div className="flex items-center gap-2">
                <button onClick={nouveauChiffrage} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                  <Plus size={15} /> Nouveau
                </button>
                <button onClick={dupliquerChiffrage} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: INK_2 }}>
                  <Copy size={14} /> Dupliquer
                </button>
                {devisList.length > 1 && (
                  <button onClick={supprimerChiffrage} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: "#B0473E" }}>
                    <Trash2 size={14} /> Supprimer
                  </button>
                )}
              </div>
            </div>

            <SectionCard
              title="Informations de l'affaire"
              icon={FileText}
              right={
                <button onClick={razChiffrage} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: "#B0473E" }}>
                  <Trash2 size={14} /> RAZ le chiffrage
                </button>
              }
            >
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
              <div className="flex items-end gap-6">
                <div>
                  <label style={{ fontSize: 11, color: MUTED }}>Cadre contractuel</label>
                  <Select
                    value={affaire.contrat}
                    onChange={(v) => setAffaire((a) => ({ ...a, contrat: v }))}
                    options={Object.entries(coefContrat).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                    style={{ maxWidth: 320 }}
                  />
                </div>
                <label className="flex items-center gap-2 pb-2" style={{ fontSize: 13, color: INK, cursor: "pointer" }}>
                  <input
                    type="checkbox"
                    checked={affaire.degressiviteActive}
                    onChange={(e) => setAffaire((a) => ({ ...a, degressiviteActive: e.target.checked }))}
                  />
                  Appliquer la dégressivité volume
                </label>
              </div>
            </SectionCard>

            {postesEquipement.map((poste) => (
              <SectionCard
                key={poste.id}
                title={poste.nom}
                subtitle="Renseignez une quantité pour autant de types d'équipements que nécessaire — pas de limite"
                icon={ClipboardList}
                right={
                  <div className="flex items-center gap-2">
                    <button onClick={() => dupliquerPosteEquipement(poste.id)} title="Dupliquer ce poste" style={{ color: INK_2 }} className="p-1.5">
                      <Copy size={16} />
                    </button>
                    {postesEquipement.length > 1 && (
                      <button onClick={() => removePosteEquipement(poste.id)} title="Supprimer ce poste" style={{ color: "#B0473E" }} className="p-1.5">
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                }
              >
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-5">
                  <div>
                    <label style={{ fontSize: 11, color: MUTED }}>Nom du poste</label>
                    <TextField value={poste.nom} onChange={(v) => updatePosteEquipement(poste.id, { nom: v })} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: MUTED }}>Journée</label>
                    <Select
                      value={poste.typeJournee}
                      onChange={(v) => updatePosteEquipement(poste.id, { typeJournee: v })}
                      options={Object.entries(majorations).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                    />
                  </div>
                </div>
                <TableauCatalogue poste={poste} catalogueTemps={catalogueTemps} catalogueDirect={catalogueDirect} setQtePoste={setQtePoste} setNiveauPoste={setNiveauPoste} />
              </SectionCard>
            ))}

            <button
              onClick={addPosteEquipement}
              className="flex items-center justify-center gap-1.5 px-4 py-3 rounded-lg text-sm font-medium border-2 border-dashed"
              style={{ borderColor: LINE, color: INK_2 }}
            >
              <Plus size={16} /> Ajouter un poste (même choix d'équipements)
            </button>

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
            <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="p-4 flex flex-wrap items-center gap-3 justify-between">
              <div className="flex items-center gap-2 flex-wrap">
                <span style={{ fontSize: 11, color: MUTED }}>Affaire</span>
                <Select
                  value={currentDevisId || ""}
                  onChange={(id) => chargerDevis(id)}
                  options={devisList.map((d) => ({ value: d.id, label: `${d.reference || "Sans référence"}${d.client ? " — " + d.client : ""}` }))}
                  style={{ minWidth: 240, width: "auto" }}
                />
              </div>
              <button onClick={exportWord} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                <Download size={15} /> Exporter en Word
              </button>
            </div>

            <SectionCard title="Résumé de l'affaire" icon={FileText}>
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
              <div className="overflow-x-auto">
                <table className="w-full" style={{ fontSize: 13 }}>
                  <thead>
                    <tr style={{ color: MUTED, textAlign: "left" }}>
                      <th className="pb-2 font-medium">Équipement</th>
                      <th className="pb-2 font-medium">Poste</th>
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
                        <td className="py-1.5" style={{ color: MUTED }}>
                          {l.posteNom}
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
                        <td className="py-1.5" style={{ color: MUTED }}>
                          —
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
                        <td colSpan={4} className="py-4 text-center" style={{ color: MUTED }}>
                          Aucun équipement renseigné pour l'instant.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </SectionCard>

            <SectionCard title="Répartition par famille" icon={ClipboardList}>
              <div className="overflow-x-auto">
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
              </div>
            </SectionCard>

            <SectionCard title="Récapitulatif financier" icon={FileText}>
              <div className="overflow-x-auto">
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
                    <td className="py-2" style={{ color: INK }}>
                      {affaire.degressiviteActive ? `Dégressivité volume (${palierDegressif.label}, ×${palierDegressif.coef})` : "Dégressivité volume (désactivée)"}
                    </td>
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
              </div>
            </SectionCard>
          </>
        )}

        {/* ---------------- ONGLET PARAMETRES ---------------- */}
        {tab === "parametres" && (
          <>
            <div className="flex items-center justify-between flex-wrap gap-2">
              <span style={{ fontSize: 12, color: MUTED }}>Vos modifications sont enregistrées automatiquement dans ce navigateur.</span>
              <div className="flex items-center gap-3">
                <button onClick={enregistrerParametresMaintenant} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                  Enregistrer les valeurs
                </button>
                <button onClick={reinitialiserParametres} className="text-sm underline" style={{ color: "#B0473E" }}>
                  Réinitialiser les valeurs par défaut
                </button>
              </div>
            </div>

            <SectionCard title="Prix jour technicien" icon={Settings2}>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {Object.entries(tarifs).map(([key, t]) => (
                  <div key={key} style={{ border: `1px solid ${LINE}`, borderRadius: 8 }} className="p-4">
                    <div style={{ fontWeight: 600, color: INK, marginBottom: 10 }}>{t.label}</div>
                    <div className="flex flex-col gap-2">
                      <label style={{ fontSize: 11, color: MUTED }}>Prix / jour</label>
                      <NumberField value={t.jour} onChange={(v) => setTarifs((s) => ({ ...s, [key]: { ...s[key], jour: v } }))} suffix="€" width="100%" />
                      <label style={{ fontSize: 11, color: MUTED }}>Prix / demi-journée (jour ÷ 2 × 1,2)</label>
                      <div style={{ fontSize: 13, color: INK, fontWeight: 600, padding: "5px 0" }}>{euros((t.jour / 2) * 1.2)}</div>
                    </div>
                  </div>
                ))}
              </div>
              <div className="mt-4 flex items-center gap-2">
                <label style={{ fontSize: 12, color: MUTED }}>Heures travaillées par jour</label>
                <NumberField value={heuresJour} onChange={setHeuresJour} suffix="h" />
                <span style={{ fontSize: 11.5, color: MUTED }}>
                  (avec les valeurs par défaut : {euros(tarifs.expert.jour / heuresJour)}/h — cohérent avec le fichier source)
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

            <SectionCard
              title="Catalogue équipements — temps par niveau de prestation"
              subtitle="Niveau 1-2 = temps simple + préparation, Niveau 3-4 = temps complexe + préparation, Niveau 1-4 = tout. Prix de revient = (heures du niveau ÷ heures/jour) × tarif jour + amortissement"
              icon={ClipboardList}
            >
              <div className="flex flex-col gap-6">
                {Object.entries(catalogueTemps).map(([key, cat]) => (
                  <div key={key}>
                    <div style={{ fontWeight: 600, color: INK, fontSize: 13 }}>{cat.label}</div>
                    {cat.marquesRef && <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 8 }}>{cat.marquesRef}</div>}
                    <div className="flex flex-col gap-3 mt-2">
                      {cat.items.map((item, i) => {
                        const heuresDefaut = heuresPourNiveau(item, item.niveauDefaut || "complet");
                        const prixRevient = (heuresDefaut / heuresJour) * tarifs.expert.jour + item.amort;
                        const majItem = (patch) =>
                          setCatalogueTemps((s) => ({ ...s, [key]: { ...s[key], items: s[key].items.map((it, j) => (j === i ? { ...it, ...patch } : it)) } }));
                        return (
                          <div key={item.id} style={{ border: `1px solid ${LINE}`, borderRadius: 8 }} className="p-3">
                            <div style={{ fontSize: 13.5, fontWeight: 600, color: INK, marginBottom: 10 }}>{item.label}</div>
                            <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
                              <div>
                                <label style={{ fontSize: 11, color: MUTED }}>Simple (1-2)</label>
                                <NumberField value={item.heuresSimple} onChange={(v) => majItem({ heuresSimple: v })} suffix="h" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED }}>Complexe (3-4)</label>
                                <NumberField value={item.heuresComplexe} onChange={(v) => majItem({ heuresComplexe: v })} suffix="h" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED }}>Préparation</label>
                                <NumberField value={item.heuresPrepa} onChange={(v) => majItem({ heuresPrepa: v })} suffix="h" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED }}>Amortissement</label>
                                <NumberField value={item.amort} onChange={(v) => majItem({ amort: v })} suffix="€" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED }}>Niveau par défaut</label>
                                <Select
                                  value={item.niveauDefaut || "complet"}
                                  onChange={(v) => majItem({ niveauDefaut: v })}
                                  options={Object.entries(NIVEAUX_PRESTATION).map(([k, label]) => ({ value: k, label }))}
                                />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED }}>Prix de revient (défaut)</label>
                                <div style={{ fontSize: 15, fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums", padding: "5px 0" }}>{euros(prixRevient)}</div>
                              </div>
                            </div>
                          </div>
                        );
                      })}
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
