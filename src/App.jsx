import React, { useState, useEffect, useMemo } from "react";
import { Plus, Trash2, Settings2, FileText, ClipboardList, Zap, Download, Copy, Flame } from "lucide-react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, WidthType, ShadingType } from "docx";
import { doc, getDoc, setDoc, collection, getDocs, addDoc, deleteDoc } from "firebase/firestore";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
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
  { min: 20, max: 999999, coef: 0.9, label: "Plus de 20 équipements" },
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
  composants: {
    label: "Composants",
    unite: "par composant",
    categories: [
      { id: "cmp-a", label: "1 € à 30 €", min: 1, max: 30, coef: 3.4 },
      { id: "cmp-b", label: "30 € à 60 €", min: 30, max: 60, coef: 2.9 },
      { id: "cmp-c", label: "60 € à 150 €", min: 60, max: 150, coef: 2.4 },
      { id: "cmp-d", label: "> 150 €", min: 150, max: 999999, coef: 2.2 },
    ],
    critereLabel: "Prix d'achat unitaire (€)",
  },
  soustraitance: {
    label: "Sous-traitance",
    unite: "par prestation",
    categories: [{ id: "st-unique", label: "Coefficient sous-traitance", min: 0, max: 999999, coef: 1.4 }],
    critereLabel: "Prix sous-traitant (€)",
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
  batteries: {
    label: "Batteries",
    unite: "u",
    // Prix d'achat = tarif fournisseur indicatif (RS Components, Yuasa...) pour
    // une batterie plomb étanche 12V standard. Coefficient repris des paliers
    // d'origine (7-12 Ah ×2,1 / 13-60 Ah ×2,0 / 61-125 Ah ×1,7 / >130 Ah ×1,5).
    // Prix de vente = prix d'achat × coefficient — modifiable à tout niveau.
    items: [
      { id: "bat-7", label: "7 Ah", prixAchat: 30, coef: 2.1, prix: 63 },
      { id: "bat-12", label: "12 Ah", prixAchat: 45, coef: 2.1, prix: 94.5 },
      { id: "bat-17", label: "17 Ah", prixAchat: 60, coef: 2.0, prix: 120 },
      { id: "bat-24", label: "24 Ah", prixAchat: 85, coef: 2.0, prix: 170 },
      { id: "bat-33", label: "33 Ah", prixAchat: 150, coef: 2.0, prix: 300 },
      { id: "bat-38", label: "38 Ah", prixAchat: 190, coef: 2.0, prix: 380 },
      { id: "bat-55", label: "55 Ah", prixAchat: 230, coef: 2.0, prix: 460 },
      { id: "bat-65", label: "65 Ah", prixAchat: 280, coef: 1.7, prix: 476 },
      { id: "bat-100", label: "100 Ah", prixAchat: 420, coef: 1.7, prix: 714 },
      { id: "bat-150", label: "150 Ah", prixAchat: 650, coef: 1.5, prix: 975 },
      { id: "bat-200", label: "200 Ah", prixAchat: 850, coef: 1.5, prix: 1275 },
    ],
  },
  onduleur: {
    label: "Onduleurs",
    unite: "u",
    items: [
      { id: "ond-simplifie", label: "Onduleur simplifié", prix: 300 },
      { id: "ond-complet", label: "Onduleur", prix: 600 },
    ],
  },
};

// ---------------------------------------------------------------------------
// FIREPRO — extinction par aérosol (norme EN 15276-1/2:2019), repris du
// fichier QUO-CHIFFRAGE FIREPRO. Masse nécessaire (g) = Volume (m3) × DAD
// (g/m3, densité d'application selon la classe de feu) × coefficient de
// remplissage × coefficient de confinement.
// ---------------------------------------------------------------------------

const DEFAULT_FIREPRO_CLASSES = [
  { id: "a-pvc", label: "Classe A - PVC", dad: 59.8 },
  { id: "a-bois", label: "Classe A - Bois", dad: 96.2 },
  { id: "b-liquide", label: "Classe B - Liquide", dad: 67.6 },
  { id: "c-gaz", label: "Classe C - Gaz", dad: 39 },
  { id: "f-huiles", label: "Classe F - Huiles", dad: 98.8 },
];

const DEFAULT_FIREPRO_CONFINEMENT = {
  bon: { label: "Bon", coef: 1 },
  standard: { label: "Standard", coef: 1.15 },
  mauvais: { label: "Mauvais", coef: 1.4 },
};

// Masse = grammes d'aérosol délivrés par le générateur. Prix HT fournisseur.
const DEFAULT_FIREPRO_GENERATEURS = [
  { id: "fp-20t", code: "10620", label: "FP-20T", masse: 12, prix: 105.6 },
  { id: "fp-20th", code: "10649", label: "FP-20TH", masse: 12, prix: 96.8 },
  { id: "fp-40t", code: "10609", label: "FP-40T", masse: 24.4, prix: 193.8 },
  { id: "fp-80t", code: "10617", label: "FP-80T", masse: 47.2, prix: 283.1 },
  { id: "fp-100s", code: "10140", label: "FP-100S", masse: 61, prix: 421.8 },
  { id: "fp-200s", code: "10142", label: "FP-200S", masse: 118, prix: 488.3 },
  { id: "fp-500s", code: "10145", label: "FP-500S", masse: 330, prix: 678.6 },
  { id: "fp-1200ts", code: "10622", label: "FP-1200TS", masse: 756, prix: 1359 },
  { id: "fp-2000ts", code: "10623", label: "FP-2000TS", masse: 1200, prix: 1569.6 },
  { id: "fp-3000ts", code: "10624", label: "FP-3000TS", masse: 1830, prix: 1825.2 },
  { id: "fp-4200ts", code: "10644", label: "FP-4200TS", masse: 2520, prix: 2394 },
  { id: "fp-5700ts", code: "10625", label: "FP-5700TS", masse: 3363, prix: 2893.4 },
  { id: "fp-100t-2gex", code: "11080", label: "FP-100T (Zone 1-2 ATEX)", masse: 61, prix: 511.1 },
  { id: "fp-200t-2gex", code: "11081", label: "FP-200T (Zone 1-2 ATEX)", masse: 118, prix: 583.3 },
  { id: "fp-500t-2gex", code: "11082", label: "FP-500T (Zone 1-2 ATEX)", masse: 330, prix: 781.2 },
  { id: "fp-100ex", code: "11028", label: "FP-100EX (Zone 0 ATEX)", masse: 61, prix: 808.2 },
  { id: "fp-200ex", code: "11073", label: "FP-200EX (Zone 0 ATEX)", masse: 118, prix: 867.6 },
  { id: "fp-500ex", code: "11074", label: "FP-500EX (Zone 0 ATEX)", masse: 330, prix: 1083.6 },
  { id: "fp-1200ex", code: "11075", label: "FP-1200EX (Zone 0 ATEX)", masse: 756, prix: 2289.6 },
  { id: "fp-2000ex", code: "11076", label: "FP-2000EX (Zone 0 ATEX)", masse: 1200, prix: 2509.2 },
  { id: "fp-3000ex", code: "11077", label: "FP-3000EX (Zone 0 ATEX)", masse: 1830, prix: 2601 },
  { id: "fp-4200ex", code: "11078", label: "FP-4200EX (Zone 0 ATEX)", masse: 2520, prix: 3216.4 },
  { id: "fp-5700ex", code: "11079", label: "FP-5700EX (Zone 0 ATEX)", masse: 3363, prix: 3828.4 },
];

const DEFAULT_FIREPRO_ACCESSOIRES = {
  centrales: {
    label: "Centrales",
    items: [
      { id: "c-fpc4r", code: "11007", label: "Minicentrale FPC-4R", prix: 50 },
      { id: "c-fpc2", code: "11416", label: "Minicentrale FPC-2", prix: 421.8 },
      { id: "c-sigmaxt", code: "10279", label: "Centrale Sigma XT", prix: 1067.4 },
      { id: "c-activateur", code: "10173", label: "Activateur séquentiel", prix: 131.1 },
      { id: "c-alim24", code: "AT-REC01", label: "Coffret alimentation 24V", prix: 990 },
    ],
  },
  detection: {
    label: "Détection",
    items: [
      { id: "d-bta57", code: "11272", label: "BTA V3 Cylindrique 57°C", prix: 115.9 },
      { id: "d-bta68", code: "11273", label: "BTA V3 Cylindrique 68°C", prix: 115.9 },
      { id: "d-bta79", code: "11274", label: "BTA V3 Cylindrique 79°C", prix: 115.9 },
      { id: "d-bta93", code: "11275", label: "BTA V3 Cylindrique 93°C", prix: 115.9 },
      { id: "d-bta141", code: "11276", label: "BTA V3 Cylindrique 141°C", prix: 115.9 },
      { id: "d-bta182", code: "11277", label: "BTA V3 Cylindrique 182°C", prix: 115.9 },
      { id: "d-btam57", code: "11284", label: "BTA V3 Mécanique 57°C", prix: 62.5 },
      { id: "d-btam68", code: "11285", label: "BTA V3 Mécanique 68°C", prix: 62.5 },
      { id: "d-btam79", code: "11286", label: "BTA V3 Mécanique 79°C", prix: 62.5 },
      { id: "d-btam93", code: "11287", label: "BTA V3 Mécanique 93°C", prix: 62.5 },
      { id: "d-btam141", code: "11288", label: "BTA V3 Mécanique 141°C", prix: 62.5 },
      { id: "d-btam182", code: "11289", label: "BTA V3 Mécanique 182°C", prix: 62.5 },
      { id: "d-btaadapt", code: "10585", label: "BTA Adaptateur", prix: 9 },
      { id: "d-fpc5-20-60", code: "11113", label: "FPC-5 V2 pour FP-20 — 60°C", prix: 92.4 },
      { id: "d-fpc5-20-70", code: "11114", label: "FPC-5 V2 pour FP-20 — 70°C", prix: 92.4 },
      { id: "d-fpc5-20-80", code: "11115", label: "FPC-5 V2 pour FP-20 — 80°C", prix: 92.4 },
      { id: "d-fpc5-20-100", code: "11116", label: "FPC-5 V2 pour FP-20 — 100°C", prix: 92.4 },
      { id: "d-fpc5-20-lhd", code: "11117", label: "FPC-5 V2 pour FP-20 — LHD", prix: 92.4 },
      { id: "d-fpc5-4080-60", code: "11118", label: "FPC-5 V2 pour FP-40/80 — 60°C", prix: 92.4 },
      { id: "d-fpc5-4080-70", code: "11119", label: "FPC-5 V2 pour FP-40/80 — 70°C", prix: 92.4 },
      { id: "d-fpc5-4080-80", code: "11120", label: "FPC-5 V2 pour FP-40/80 — 80°C", prix: 92.4 },
      { id: "d-fpc5-4080-100", code: "11121", label: "FPC-5 V2 pour FP-40/80 — 100°C", prix: 92.4 },
      { id: "d-fpc5-4080-lhd", code: "11122", label: "FPC-5 V2 pour FP-40/80 — LHD", prix: 92.4 },
      { id: "d-fpc5-100500-60", code: "11123", label: "FPC-5 V2 pour FP-100/500 — 60°C", prix: 88 },
      { id: "d-fpc5-100500-70", code: "11124", label: "FPC-5 V2 pour FP-100/500 — 70°C", prix: 88 },
      { id: "d-fpc5-100500-80", code: "11125", label: "FPC-5 V2 pour FP-100/500 — 80°C", prix: 88 },
      { id: "d-fpc5-100500-100", code: "11126", label: "FPC-5 V2 pour FP-100/500 — 100°C", prix: 88 },
      { id: "d-fpc5-100500-lhd", code: "11127", label: "FPC-5 V2 pour FP-100/500 — LHD", prix: 88 },
      { id: "d-cordon68", code: "10244", label: "Cordon thermique 68°C", prix: 22.5 },
      { id: "d-cordon88", code: "10245", label: "Cordon thermique 88°C", prix: 22.5 },
      { id: "d-cordon105", code: "10246", label: "Cordon thermique 105°C", prix: 22.5 },
      { id: "d-cordon185", code: "10607", label: "Cordon thermique 185°C", prix: 25 },
      { id: "d-termlhd", code: "10830", label: "Terminaison LHD", prix: 9 },
      { id: "d-baseopt", code: "10990", label: "Base détecteur optique", prix: 18 },
      { id: "d-detchal", code: "10992", label: "Détecteur de chaleur", prix: 47.5 },
      { id: "d-detfum", code: "10991", label: "Détecteur de fumée", prix: 47.5 },
    ],
  },
  relais: {
    label: "Contacts, relayage",
    items: [
      { id: "r-thermo", code: "10994", label: "Contact thermorupteur", prix: 12 },
      { id: "r-kittemp", code: "KIT-TEMP", label: "Kit temporisation extinction", prix: 500 },
      { id: "r-kitsign", code: "KIT-SIGN", label: "Kit signalisation local", prix: 1000 },
    ],
  },
  fixation: {
    label: "Fixations",
    items: [
      { id: "f-bracket", code: "10492", label: "Bracket FP-100/500", prix: 21 },
      { id: "f-aimants", code: "10640", label: "Aimants", prix: 18 },
      { id: "f-vis", code: "FIX-VIS", label: "Vis", prix: 2 },
    ],
  },
  mainOeuvre: {
    label: "Main d'œuvre & transport",
    items: [
      { id: "mo-mesa", code: "MES-A", label: "Mise en service (Aérosol)", prix: 380 },
      { id: "mo-mesc", code: "MES-C", label: "Mise en service (Centrale)", prix: 380 },
      { id: "mo-instac", code: "INST-AC", label: "Installation (Aérosol Cylindrique)", prix: 250 },
      { id: "mo-instab", code: "INST-AB", label: "Installation (Aérosol Box)", prix: 550 },
      { id: "mo-instc", code: "INST-C", label: "Installation (Centrale)", prix: 450 },
      { id: "mo-instd", code: "INST-D", label: "Installation (Détecteur)", prix: 250 },
      { id: "mo-instkt", code: "INST-KT", label: "Installation (Temporisation)", prix: 300 },
      { id: "mo-instks", code: "INST-KS", label: "Installation (Signalisation locale)", prix: 1000 },
      { id: "mo-transp", code: "TRANSP", label: "Transport", prix: 75 },
    ],
  },
};

function nouvelleZoneFirePro(n) {
  return {
    id: uid(),
    nom: `Local ${n}`,
    largeur: 0,
    profondeur: 0,
    hauteur: 0,
    volumeManuel: 0, // si > 0, prioritaire sur largeur × profondeur × hauteur
    classeId: "a-pvc",
    coefRemplissage: 1,
    confinementId: "standard",
    generateurs: {}, // { genId: quantite }
    accessoires: {}, // { itemId: quantite }, toutes familles confondues
  };
}

function volumeZoneFirePro(zone) {
  if (zone.volumeManuel > 0) return zone.volumeManuel;
  return (zone.largeur || 0) * (zone.profondeur || 0) * (zone.hauteur || 0);
}

function computeZoneFirePro(zone, classes, confinements, generateurs, accessoires) {
  const classe = classes.find((c) => c.id === zone.classeId) || classes[0];
  const confinement = confinements[zone.confinementId] || confinements.standard;
  const volume = volumeZoneFirePro(zone);
  const masseNecessaire = volume * classe.dad * (zone.coefRemplissage || 0) * confinement.coef;
  let masseEffective = 0;
  let montantGenerateurs = 0;
  generateurs.forEach((g) => {
    const qte = zone.generateurs[g.id] || 0;
    masseEffective += qte * g.masse;
    montantGenerateurs += qte * g.prix;
  });
  let montantAccessoires = 0;
  Object.values(accessoires).forEach((famille) => {
    famille.items.forEach((item) => {
      const qte = zone.accessoires[item.id] || 0;
      montantAccessoires += qte * item.prix;
    });
  });
  return {
    volume,
    classe,
    confinement,
    masseNecessaire,
    masseEffective,
    montantGenerateurs,
    montantAccessoires,
    montantTotal: montantGenerateurs + montantAccessoires,
    suffisant: masseEffective >= masseNecessaire && masseNecessaire > 0,
  };
}

// Familles pour les "lignes libres" (batteries, composants, poste manuel)
const FAMILLES_LIBRES = [
  { id: "maindoeuvre", label: "Main-d'œuvre", type: "manuel" },
  { id: "composants", label: "Composants", type: "coef" },
  { id: "soustraitance", label: "Sous-traitance", type: "coef" },
  { id: "manuel", label: "Poste libre (autre)", type: "manuel" },
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
  onduleur: "Onduleurs",
  maindoeuvre: "Main-d'œuvre",
  soustraitance: "Sous-traitance",
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
  // c.max peut valoir null si une ancienne sauvegarde a transformé Infinity
  // en null (JSON ne sait pas encoder Infinity) — on le traite alors comme
  // "sans limite haute" plutôt que de faire échouer la comparaison.
  return catalogue.categories.find((c) => valeur >= c.min && valeur <= (c.max ?? Infinity)) || catalogue.categories[0];
}

function nouveauPosteEquipement(n) {
  return { id: uid(), nom: `Poste ${n}`, niveauTechnicien: "expert", typeJournee: "semaine", quantites: {}, niveaux: {} };
}

// ---------------------------------------------------------------------------

function SectionCard({ title, icon: Icon, children, right, subtitle, bg }) {
  return (
    <div style={{ background: bg || "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="overflow-hidden">
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
  const migrerCatalogueDirect = (savedDirect) => {
    const base = { ...DEFAULT_CATALOGUE_DIRECT, ...(savedDirect || {}) };
    if (!base.batteries || !base.batteries.items || base.batteries.items[0]?.prixAchat == null) {
      base.batteries = DEFAULT_CATALOGUE_DIRECT.batteries;
    }
    return base;
  };

  const [tarifs, setTarifs] = useState(tarifsValides(saved.tarifs) ? saved.tarifs : DEFAULT_TARIFS);
  const [majorations, setMajorations] = useState(saved.majorations || DEFAULT_MAJORATIONS);
  const [degressivite, setDegressivite] = useState(saved.degressivite || DEFAULT_DEGRESSIVITE);
  const [coefContrat, setCoefContrat] = useState(saved.coefContrat || DEFAULT_COEF_CONTRAT);
  const [heuresJour, setHeuresJour] = useState(saved.heuresJour ?? 7);

  const [catalogueTemps, setCatalogueTemps] = useState(catalogueTempsValide(saved.catalogueTemps) ? saved.catalogueTemps : DEFAULT_CATALOGUE_TEMPS);
  const [catalogueCoef, setCatalogueCoef] = useState({ ...DEFAULT_CATALOGUE_COEF, ...(saved.catalogueCoef || {}) });
  const [catalogueDirect, setCatalogueDirect] = useState(migrerCatalogueDirect(saved.catalogueDirect));

  const capaciteAh = (label) => {
    const m = String(label).match(/[\d.,]+/);
    return m ? parseFloat(m[0].replace(",", ".")) : Infinity;
  };
  const ajouterLigneCatalogueDirect = (familleKey) =>
    setCatalogueDirect((s) => {
      if (familleKey === "batteries") {
        const nouvelItem = { id: uid(), label: "Nouvelle capacité", prixAchat: 0, coef: 2, prix: 0 };
        const items = [...s.batteries.items, nouvelItem].sort((a, b) => capaciteAh(a.label) - capaciteAh(b.label));
        return { ...s, batteries: { ...s.batteries, items } };
      }
      const nouvelItem = { id: uid(), label: "Nouvel article", prix: 0 };
      return { ...s, [familleKey]: { ...s[familleKey], items: [...s[familleKey].items, nouvelItem] } };
    });
  const supprimerLigneCatalogueDirect = (familleKey, itemId) =>
    setCatalogueDirect((s) => ({ ...s, [familleKey]: { ...s[familleKey], items: s[familleKey].items.filter((it) => it.id !== itemId) } }));
  const renommerLigneCatalogueDirect = (familleKey, itemId, label) =>
    setCatalogueDirect((s) => ({ ...s, [familleKey]: { ...s[familleKey], items: s[familleKey].items.map((it) => (it.id === itemId ? { ...it, label } : it)) } }));

  // FirePro — ajout/suppression/renommage des générateurs
  const ajouterGenerateurFirePro = () =>
    setFireproGenerateurs((arr) => [...arr, { id: uid(), code: "", label: "Nouveau générateur", masse: 0, prix: 0 }]);
  const supprimerGenerateurFirePro = (id) => setFireproGenerateurs((arr) => arr.filter((g) => g.id !== id));
  const renommerGenerateurFirePro = (id, label) => setFireproGenerateurs((arr) => arr.map((g) => (g.id === id ? { ...g, label } : g)));

  // FirePro — ajout/suppression/renommage des accessoires (par famille)
  const ajouterAccessoireFirePro = (famId) =>
    setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: [...s[famId].items, { id: uid(), code: "", label: "Nouvel article", prix: 0 }] } }));
  const supprimerAccessoireFirePro = (famId, itemId) =>
    setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: s[famId].items.filter((it) => it.id !== itemId) } }));
  const renommerAccessoireFirePro = (famId, itemId, label) =>
    setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: s[famId].items.map((it) => (it.id === itemId ? { ...it, label } : it)) } }));

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
      zonesFirePro: [nouvelleZoneFirePro(1)],
    };
  }

  const [affaire, setAffaire] = useState(saved.affaire || devisVierge().affaire);
  const [postesEquipement, setPostesEquipement] = useState(saved.postesEquipement || devisVierge().postesEquipement);
  const [lignesLibres, setLignesLibres] = useState(saved.lignesLibres || []);
  const [zonesFirePro, setZonesFirePro] = useState(saved.zonesFirePro || devisVierge().zonesFirePro);
  const [nbAAjouter, setNbAAjouter] = useState(1);

  // Paramètres FirePro — partagés entre chiffrages, comme tarifs/catalogues.
  const [fireproClasses, setFireproClasses] = useState(saved.fireproClasses || DEFAULT_FIREPRO_CLASSES);
  const [fireproConfinement, setFireproConfinement] = useState(saved.fireproConfinement || DEFAULT_FIREPRO_CONFINEMENT);
  const [fireproGenerateurs, setFireproGenerateurs] = useState(saved.fireproGenerateurs || DEFAULT_FIREPRO_GENERATEURS);
  const [fireproAccessoires, setFireproAccessoires] = useState(saved.fireproAccessoires || DEFAULT_FIREPRO_ACCESSOIRES);
  const [fireproCoefAjustement, setFireproCoefAjustement] = useState(saved.fireproCoefAjustement ?? 1);

  const addZoneFirePro = () => setZonesFirePro((zs) => [...zs, nouvelleZoneFirePro(zs.length + 1)]);
  const dupliquerZoneFirePro = (id) =>
    setZonesFirePro((zs) => {
      const src = zs.find((z) => z.id === id);
      if (!src) return zs;
      return [...zs, { ...src, id: uid(), nom: `${src.nom} (copie)`, generateurs: { ...src.generateurs }, accessoires: { ...src.accessoires } }];
    });
  const updateZoneFirePro = (id, patch) => setZonesFirePro((zs) => zs.map((z) => (z.id === id ? { ...z, ...patch } : z)));
  const removeZoneFirePro = (id) => setZonesFirePro((zs) => (zs.length > 1 ? zs.filter((z) => z.id !== id) : zs));
  const setGenerateurZone = (zoneId, genId, v) =>
    setZonesFirePro((zs) => zs.map((z) => (z.id === zoneId ? { ...z, generateurs: { ...z.generateurs, [genId]: Math.max(0, v) } } : z)));
  const setAccessoireZone = (zoneId, itemId, v) =>
    setZonesFirePro((zs) => zs.map((z) => (z.id === zoneId ? { ...z, accessoires: { ...z.accessoires, [itemId]: Math.max(0, v) } } : z)));

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
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false); // évite un flash de l'écran de connexion pendant la toute première vérification
  const [authTimedOut, setAuthTimedOut] = useState(false);
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);

  // Liste des chiffrages (devis) disponibles, et celui actuellement affiché.
  const [devisList, setDevisList] = useState(saved.devisList || []);
  const [currentDevisId, setCurrentDevisId] = useState(saved.currentDevisId || null);
  const [devisReady, setDevisReady] = useState(false); // évite d'écraser un chiffrage avant la fin du chargement initial

  // Accès protégé par email + mot de passe (Firebase Authentication). Pas de
  // connexion automatique : tant que personne n'est identifié, l'app affiche
  // un écran de connexion au lieu du chiffrage.
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (user) => {
      if (user && user.isAnonymous) {
        // Reliquat de l'ancienne connexion anonyme automatique : on l'efface,
        // ça ne doit pas être considéré comme une connexion valide.
        signOut(auth);
        setCurrentUser(null);
        setAuthReady(false);
        setAuthChecked(true);
        return;
      }
      setCurrentUser(user);
      setAuthReady(!!user);
      setAuthChecked(true);
    });
    const timeout = setTimeout(() => setAuthTimedOut(true), 6000);
    return () => {
      unsubscribe();
      clearTimeout(timeout);
    };
  }, []);

  const seConnecter = async (e) => {
    e.preventDefault();
    setLoginError("");
    setLoginLoading(true);
    try {
      await signInWithEmailAndPassword(auth, loginEmail.trim(), loginPassword);
    } catch (err) {
      setLoginError("Adresse e-mail ou mot de passe incorrect.");
    } finally {
      setLoginLoading(false);
    }
  };

  const seDeconnecter = () => {
    signOut(auth);
  };

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
          setCatalogueCoef({ ...DEFAULT_CATALOGUE_COEF, ...(data.catalogueCoef || {}) });
          setCatalogueDirect(migrerCatalogueDirect(data.catalogueDirect));
          if (data.fireproClasses) setFireproClasses(data.fireproClasses);
          if (data.fireproConfinement) setFireproConfinement(data.fireproConfinement);
          if (data.fireproGenerateurs) setFireproGenerateurs(data.fireproGenerateurs);
          if (data.fireproAccessoires) setFireproAccessoires(data.fireproAccessoires);
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
        setZonesFirePro(target.zonesFirePro || devisVierge().zonesFirePro);
        setFireproCoefAjustement(target.fireproCoefAjustement ?? 1);
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
          setZonesFirePro(data.zonesFirePro || devisVierge().zonesFirePro);
          setFireproCoefAjustement(data.fireproCoefAjustement ?? 1);
        }
        setDevisReady(true);
      })
      .catch(() => setDevisReady(true));
  };

  const nouveauChiffrage = async () => {
    const blank = devisVierge();
    setDevisReady(false);
    try {
      const payload = { ...blank, fireproCoefAjustement: 1, updatedAt: Date.now() };
      const ref = await addDoc(collection(db, DEVIS_COLLECTION), payload);
      setDevisList((l) => [{ id: ref.id, reference: blank.affaire.reference, client: "", updatedAt: payload.updatedAt }, ...l]);
      setCurrentDevisId(ref.id);
      setAffaire(blank.affaire);
      setPostesEquipement(blank.postesEquipement);
      setLignesLibres(blank.lignesLibres);
      setZonesFirePro(blank.zonesFirePro);
      setFireproCoefAjustement(1);
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
        zonesFirePro: zonesFirePro.map((z) => ({ ...z, generateurs: { ...z.generateurs }, accessoires: { ...z.accessoires } })),
        fireproCoefAjustement,
        updatedAt: Date.now(),
      };
      const ref = await addDoc(collection(db, DEVIS_COLLECTION), payload);
      setDevisList((l) => [{ id: ref.id, reference: payload.affaire.reference, client: payload.affaire.client, updatedAt: payload.updatedAt }, ...l]);
      setCurrentDevisId(ref.id);
      setAffaire(payload.affaire);
      setPostesEquipement(payload.postesEquipement);
      setLignesLibres(payload.lignesLibres);
      setZonesFirePro(payload.zonesFirePro);
      setFireproCoefAjustement(payload.fireproCoefAjustement);
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
    const payload = {
      tarifs,
      majorations,
      degressivite,
      coefContrat,
      heuresJour,
      catalogueTemps,
      catalogueCoef,
      catalogueDirect,
      fireproClasses,
      fireproConfinement,
      fireproGenerateurs,
      fireproAccessoires,
    };
    const t = setTimeout(() => {
      setDoc(doc(db, FIRESTORE_DOC), payload)
        .then(() => setSyncState("synced"))
        .catch(() => setSyncState("error"));
    }, 1000);
    return () => clearTimeout(t);
  }, [
    authReady,
    tarifs,
    majorations,
    degressivite,
    coefContrat,
    heuresJour,
    catalogueTemps,
    catalogueCoef,
    catalogueDirect,
    fireproClasses,
    fireproConfinement,
    fireproGenerateurs,
    fireproAccessoires,
  ]);

  // Enregistre automatiquement le chiffrage courant (affaire + postes +
  // lignes libres) dans son propre document, sans toucher aux autres.
  useEffect(() => {
    if (!authReady || !devisReady || !currentDevisId) return;
    setSyncState("syncing");
    const t = setTimeout(() => {
      const payload = { affaire, postesEquipement, lignesLibres, zonesFirePro, fireproCoefAjustement, updatedAt: Date.now() };
      setDoc(doc(db, DEVIS_COLLECTION, currentDevisId), payload)
        .then(() => {
          setSyncState("synced");
          setDevisList((l) => l.map((d) => (d.id === currentDevisId ? { ...d, reference: affaire.reference, client: affaire.client, updatedAt: payload.updatedAt } : d)));
        })
        .catch(() => setSyncState("error"));
    }, 1000);
    return () => clearTimeout(t);
  }, [authReady, devisReady, currentDevisId, affaire, postesEquipement, lignesLibres, zonesFirePro, fireproCoefAjustement]);

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
          zonesFirePro,
          fireproClasses,
          fireproConfinement,
          fireproGenerateurs,
          fireproAccessoires,
          fireproCoefAjustement,
        })
      );
    } catch {
      // stockage local indisponible (navigation privée, quota dépassé...) — on continue sans bloquer
    }
  }, [
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
    zonesFirePro,
    fireproClasses,
    fireproConfinement,
    fireproGenerateurs,
    fireproAccessoires,
    fireproCoefAjustement,
  ]);

  // Enregistre immédiatement les paramètres/catalogues, sans attendre le
  // délai automatique (utile pour une confirmation explicite à l'utilisateur).
  const enregistrerParametresMaintenant = () => {
    setSyncState("syncing");
    setDoc(doc(db, FIRESTORE_DOC), { tarifs, majorations, degressivite, coefContrat, heuresJour, catalogueTemps, catalogueCoef, catalogueDirect })
      .then(() => setSyncState("synced"))
      .catch(() => setSyncState("error"));
  };

  // Annule les modifications non enregistrées : recharge les dernières
  // valeurs réellement enregistrées dans le cloud (pas les valeurs d'usine).
  const annulerModifications = () => {
    if (!window.confirm("Annuler vos modifications non enregistrées et recharger les dernières valeurs enregistrées ?")) return;
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
          setCatalogueCoef({ ...DEFAULT_CATALOGUE_COEF, ...(data.catalogueCoef || {}) });
          setCatalogueDirect(migrerCatalogueDirect(data.catalogueDirect));
          if (data.fireproClasses) setFireproClasses(data.fireproClasses);
          if (data.fireproConfinement) setFireproConfinement(data.fireproConfinement);
          if (data.fireproGenerateurs) setFireproGenerateurs(data.fireproGenerateurs);
          if (data.fireproAccessoires) setFireproAccessoires(data.fireproAccessoires);
        }
        setSyncState("synced");
      })
      .catch(() => setSyncState("error"));
  };

  const reinitialiserParametres = () => {
    if (!window.confirm("Réinitialiser tous les paramètres et catalogues aux valeurs D'USINE (ça remplacera aussi la dernière version enregistrée) ?")) return;
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
    if (!window.confirm("Remettre à zéro le chiffrage en cours (infos affaire, équipements saisis, lignes libres, zones FirePro) ? Cette action est irréversible.")) return;
    const blank = devisVierge();
    setAffaire(blank.affaire);
    setPostesEquipement(blank.postesEquipement);
    setLignesLibres(blank.lignesLibres);
    setZonesFirePro(blank.zonesFirePro);
    setFireproCoefAjustement(1);
  };


  const addLigneLibre = (n = 1) =>
    setLignesLibres((ls) => [
      ...ls,
      ...Array.from({ length: Math.max(1, Math.round(n)) }, () => ({
        id: uid(),
        famille: "composants",
        quantite: 1,
        niveauTechnicien: "expert",
        typeJournee: "semaine",
        prixAchatUnitaire: 20,
        categorieId: "",
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
      const categorie = cat.manuelCategorie
        ? cat.categories.find((c) => c.id === l.categorieId) || cat.categories[0]
        : findCoefCategory(cat, l.prixAchatUnitaire);
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
            // Décomposition réelle des heures par type de temps (simple / complexe /
            // préparation), qu'importe si le niveau choisi est "1-4 complet" ou non —
            // sert au détail par niveau dans le récapitulatif.
            const joursSimple = (niveau === "complexe" ? 0 : item.heuresSimple / heuresJour) * qte;
            const joursComplexe = (niveau === "simple" ? 0 : item.heuresComplexe / heuresJour) * qte;
            const joursPrepa = (item.heuresPrepa / heuresJour) * qte;
            out.push({
              id: `${poste.id}-${item.id}`,
              posteId: poste.id,
              posteNom: poste.nom,
              famille: famId,
              label: item.label,
              qte,
              joursHomme,
              joursSimple,
              joursComplexe,
              joursPrepa,
              niveau,
              montant: montantUnitaire * qte,
            });
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

  // Jours-hommes des équipements du catalogue, détaillés par type de temps
  // réel (simple / complexe / préparation) — même quand le niveau choisi est
  // "1-4 complet", on voit la part de chaque composante, pas juste un total.
  const joursHommeParNiveau = { simple: 0, complexe: 0, prepa: 0 };
  lignesCatalogue.forEach((l) => {
    joursHommeParNiveau.simple += l.joursSimple || 0;
    joursHommeParNiveau.complexe += l.joursComplexe || 0;
    joursHommeParNiveau.prepa += l.joursPrepa || 0;
  });
  const joursHommeLignesLibres = lignesLibresCalc.reduce((s, l) => s + l.joursHomme, 0);

  const totalAvantCoef =
    lignesCatalogue.reduce((s, l) => s + l.montant, 0) + lignesLibresCalc.reduce((s, l) => s + l.montant, 0);

  const palierDegressif =
    degressivite.find((d) => totalEquipements > d.min && totalEquipements <= (d.max ?? Infinity)) || degressivite[degressivite.length - 1];
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

  // ---- Calculs FirePro (dimensionnement + chiffrage par zone) ----
  const zonesFireProCalc = useMemo(
    () => zonesFirePro.map((zone) => ({ zone, ...computeZoneFirePro(zone, fireproClasses, fireproConfinement, fireproGenerateurs, fireproAccessoires) })),
    [zonesFirePro, fireproClasses, fireproConfinement, fireproGenerateurs, fireproAccessoires]
  );
  const totalFireProAvantCoef = zonesFireProCalc.reduce((s, z) => s + z.montantTotal, 0);
  const totalFirePro = totalFireProAvantCoef * (fireproCoefAjustement || 1);

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

  // Écran de connexion : affiché tant que personne n'est identifié (ou
  // pendant la toute première vérification, pour éviter un flash).
  if (!authChecked || !currentUser) {
    return (
      <div style={{ background: PAPER, minHeight: "100%", fontFamily: "Inter, system-ui, sans-serif", colorScheme: "light" }} className="w-full flex items-center justify-center" >
        <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10, minWidth: 320, maxWidth: 380 }} className="p-6 mx-4">
          {!authChecked && authTimedOut && (
            <div style={{ fontSize: 13, color: INK }} className="text-center">
              <div style={{ fontWeight: 600, marginBottom: 6 }}>La vérification de connexion ne répond pas.</div>
              <div style={{ color: MUTED, fontSize: 12.5 }}>
                Ça peut venir d'un bloqueur de contenu, d'une version de navigateur ancienne, ou d'un souci réseau. Essayez de recharger la page, ou un autre navigateur.
              </div>
            </div>
          )}
          {authChecked && (
            <>
              <div className="flex items-center gap-3 mb-5">
                <div style={{ background: AMBER, width: 34, height: 34, borderRadius: 8 }} className="flex items-center justify-center shrink-0">
                  <Zap size={18} color={INK} strokeWidth={2.5} />
                </div>
                <div>
                  <div style={{ color: INK, fontWeight: 700, fontSize: 16 }}>HT Maintenance</div>
                  <div style={{ color: MUTED, fontSize: 12 }}>Outil de chiffrage — accès protégé</div>
                </div>
              </div>
              <form onSubmit={seConnecter} className="flex flex-col gap-3">
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Adresse e-mail</label>
                  <TextField value={loginEmail} onChange={setLoginEmail} placeholder="vous@exemple.com" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Mot de passe</label>
                  <input
                    type="password"
                    value={loginPassword}
                    onChange={(e) => setLoginPassword(e.target.value)}
                    style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: "6px 9px", fontSize: 13, color: INK, background: "#fff", colorScheme: "light", width: "100%" }}
                  />
                </div>
                {loginError && <div style={{ color: "#B0473E", fontSize: 12.5 }}>{loginError}</div>}
                <button
                  type="submit"
                  disabled={loginLoading}
                  className="w-full py-2 rounded-md text-sm font-medium mt-1"
                  style={{ background: AMBER, color: INK, opacity: loginLoading ? 0.6 : 1 }}
                >
                  {loginLoading ? "Connexion…" : "Se connecter"}
                </button>
              </form>
            </>
          )}
        </div>
      </div>
    );
  }

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
            <button onClick={seDeconnecter} style={{ color: "#9AA6B2", fontSize: 11 }} className="underline">
              Déconnexion
            </button>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
        <div className="max-w-6xl mx-auto flex px-4 md:px-6 overflow-x-auto">
          {tabBtn("chiffrage", "Chiffrage", ClipboardList)}
          {tabBtn("recap", "Récapitulatif", FileText)}
          {tabBtn("firepro", "FirePro", Flame)}
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
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Référence</label>
                  <TextField value={affaire.reference} onChange={(v) => setAffaire((a) => ({ ...a, reference: v }))} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Client</label>
                  <TextField value={affaire.client} onChange={(v) => setAffaire((a) => ({ ...a, client: v }))} placeholder="Nom du client" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Site</label>
                  <TextField value={affaire.site} onChange={(v) => setAffaire((a) => ({ ...a, site: v }))} placeholder="Site / adresse" />
                </div>
              </div>
              <div className="flex items-end gap-6">
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Cadre contractuel</label>
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

            <div className="grid grid-cols-1 md:grid-cols-2 gap-5 items-start">
              {postesEquipement.map((poste, index) => (
                <SectionCard
                  key={poste.id}
                  title={poste.nom}
                  subtitle="Renseignez une quantité pour autant de types d'équipements que nécessaire — pas de limite"
                  icon={ClipboardList}
                  bg={index % 2 === 1 ? "#F3F4F5" : "#fff"}
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
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Nom du poste</label>
                      <TextField value={poste.nom} onChange={(v) => updatePosteEquipement(poste.id, { nom: v })} />
                    </div>
                    <div>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Journée</label>
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
            </div>

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
                          <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Type</label>
                          <Select value={l.famille} onChange={(v) => updateLigneLibre(l.id, { famille: v })} options={FAMILLES_LIBRES.map((f) => ({ value: f.id, label: f.label }))} />
                        </div>

                        {famille.type === "coef" && (
                          <div className="md:col-span-2">
                            <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>{catalogueCoef[l.famille].critereLabel}</label>
                            <NumberField value={l.prixAchatUnitaire} onChange={(v) => updateLigneLibre(l.id, { prixAchatUnitaire: v })} width="100%" />
                          </div>
                        )}

                        {famille.type === "coef" && catalogueCoef[l.famille].manuelCategorie && (
                          <div className="md:col-span-1">
                            <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Catégorie</label>
                            <Select
                              value={l.categorieId || catalogueCoef[l.famille].categories[0].id}
                              onChange={(v) => updateLigneLibre(l.id, { categorieId: v })}
                              options={catalogueCoef[l.famille].categories.map((c) => ({ value: c.id, label: c.label }))}
                            />
                          </div>
                        )}

                        {famille.type === "manuel" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Libellé</label>
                            <TextField value={l.libelleManuel} onChange={(v) => updateLigneLibre(l.id, { libelleManuel: v })} placeholder="Description du poste" />
                          </div>
                        )}

                        <div className="md:col-span-1">
                          <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Qté</label>
                          <NumberField value={l.quantite} onChange={(v) => updateLigneLibre(l.id, { quantite: v })} width="100%" />
                        </div>

                        {famille.type === "manuel" && (
                          <div className="md:col-span-1">
                            <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>j/u</label>
                            <NumberField value={l.joursManuel} onChange={(v) => updateLigneLibre(l.id, { joursManuel: v })} width="100%" />
                          </div>
                        )}

                        {famille.type === "manuel" && (
                          <>
                            <div className="md:col-span-2">
                              <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Technicien (préremplit le prix)</label>
                              <Select
                                value={l.niveauTechnicien}
                                onChange={(v) => updateLigneLibre(l.id, { niveauTechnicien: v, prixJour: tarifs[v].jour })}
                                options={Object.entries(tarifs).map(([k, v]) => ({ value: k, label: v.label }))}
                              />
                            </div>
                            <div className="md:col-span-1">
                              <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Prix / jour</label>
                              <NumberField value={l.prixJour} onChange={(v) => updateLigneLibre(l.id, { prixJour: v })} suffix="€" width="100%" />
                            </div>
                            <div className="md:col-span-1">
                              <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Journée</label>
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

            {postesEquipement.map((poste) => {
              const lignesPoste = lignesCatalogue.filter((l) => l.posteId === poste.id);
              const totalPoste = lignesPoste.reduce((s, l) => s + l.montant, 0);
              if (lignesPoste.length === 0) return null;
              return (
                <SectionCard key={poste.id} title={poste.nom} subtitle={`${lignesPoste.length} équipement${lignesPoste.length > 1 ? "s" : ""}`} icon={ClipboardList}>
                  <div className="overflow-x-auto">
                    <table className="w-full" style={{ fontSize: 13 }}>
                      <thead>
                        <tr style={{ color: MUTED, textAlign: "left" }}>
                          <th className="pb-2 font-medium">Équipement</th>
                          <th className="pb-2 font-medium text-right">Qté</th>
                          <th className="pb-2 font-medium text-right">Montant HT</th>
                        </tr>
                      </thead>
                      <tbody>
                        {lignesPoste.map((l) => (
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
                        <tr style={{ borderTop: `2px solid ${INK}` }}>
                          <td className="pt-2" style={{ fontWeight: 700, color: INK }} colSpan={2}>
                            Sous-total {poste.nom}
                          </td>
                          <td className="pt-2 text-right" style={{ fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>
                            {euros(totalPoste)}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </SectionCard>
              );
            })}

            <SectionCard title="Lignes libres" subtitle="Batteries, composants, postes manuels — non rattachés à un poste d'équipements" icon={ClipboardList}>
              <div className="overflow-x-auto">
                <table className="w-full" style={{ fontSize: 13 }}>
                  <thead>
                    <tr style={{ color: MUTED, textAlign: "left" }}>
                      <th className="pb-2 font-medium">Désignation</th>
                      <th className="pb-2 font-medium text-right">Qté</th>
                      <th className="pb-2 font-medium text-right">Montant HT</th>
                    </tr>
                  </thead>
                  <tbody>
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
                    {lignesLibresCalc.length === 0 && (
                      <tr>
                        <td colSpan={3} className="py-4 text-center" style={{ color: MUTED }}>
                          Aucune ligne libre.
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
                    <td className="py-2 pl-3" style={{ color: MUTED, fontSize: 12.5 }}>— Niveau 1-2 (simple)</td>
                    <td className="py-2 text-right" style={{ color: MUTED, fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>{joursHommeParNiveau.simple.toFixed(2)} j</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2 pl-3" style={{ color: MUTED, fontSize: 12.5 }}>— Niveau 3-4 (complexe)</td>
                    <td className="py-2 text-right" style={{ color: MUTED, fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>{joursHommeParNiveau.complexe.toFixed(2)} j</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2 pl-3" style={{ color: MUTED, fontSize: 12.5 }}>— Préparation</td>
                    <td className="py-2 text-right" style={{ color: MUTED, fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>{joursHommeParNiveau.prepa.toFixed(2)} j</td>
                  </tr>
                  <tr style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-2 pl-3" style={{ color: MUTED, fontSize: 12.5 }}>— Lignes libres</td>
                    <td className="py-2 text-right" style={{ color: MUTED, fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>{joursHommeLignesLibres.toFixed(2)} j</td>
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

        {/* ---------------- ONGLET FIREPRO ---------------- */}
        {tab === "firepro" && (
          <>
            {zonesFireProCalc.map(({ zone, volume, masseNecessaire, masseEffective, montantTotal, suffisant }, index) => (
              <SectionCard
                key={zone.id}
                title={zone.nom}
                subtitle="Dimensionnement EN 15276 : masse nécessaire = volume × densité de la classe (g/m³) × coef. remplissage × coef. confinement"
                icon={Flame}
                bg={index % 2 === 1 ? "#F3F4F5" : "#fff"}
                right={
                  <div className="flex items-center gap-2">
                    <button onClick={() => dupliquerZoneFirePro(zone.id)} title="Dupliquer ce local" style={{ color: INK_2 }} className="p-1.5">
                      <Copy size={16} />
                    </button>
                    {zonesFirePro.length > 1 && (
                      <button onClick={() => removeZoneFirePro(zone.id)} title="Supprimer ce local" style={{ color: "#B0473E" }} className="p-1.5">
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                }
              >
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Nom du local</label>
                    <TextField value={zone.nom} onChange={(v) => updateZoneFirePro(zone.id, { nom: v })} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Classe de feu</label>
                    <Select
                      value={zone.classeId}
                      onChange={(v) => updateZoneFirePro(zone.id, { classeId: v })}
                      options={fireproClasses.map((c) => ({ value: c.id, label: `${c.label} (${c.dad} g/m³)` }))}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Confinement</label>
                    <Select
                      value={zone.confinementId}
                      onChange={(v) => updateZoneFirePro(zone.id, { confinementId: v })}
                      options={Object.entries(fireproConfinement).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-1">
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Largeur (m)</label>
                    <NumberField value={zone.largeur} onChange={(v) => updateZoneFirePro(zone.id, { largeur: v })} width="100%" />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Profondeur (m)</label>
                    <NumberField value={zone.profondeur} onChange={(v) => updateZoneFirePro(zone.id, { profondeur: v })} width="100%" />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Hauteur (m)</label>
                    <NumberField value={zone.hauteur} onChange={(v) => updateZoneFirePro(zone.id, { hauteur: v })} width="100%" />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Coef. remplissage</label>
                    <NumberField value={zone.coefRemplissage} onChange={(v) => updateZoneFirePro(zone.id, { coefRemplissage: v })} suffix="×" width="100%" />
                  </div>
                </div>
                <div className="mb-4">
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Volume forcé (m³) — optionnel, prioritaire sur largeur × profondeur × hauteur</label>
                  <NumberField value={zone.volumeManuel} onChange={(v) => updateZoneFirePro(zone.id, { volumeManuel: v })} suffix="m³" width={140} />
                </div>

                <div
                  style={{ background: suffisant ? "#EAF6EC" : "#FDEEEC", border: `1px solid ${suffisant ? "#8FBF98" : "#E3A79E"}`, borderRadius: 8 }}
                  className="p-3 mb-5 flex flex-wrap gap-x-6 gap-y-1 items-center justify-between"
                >
                  <span style={{ fontSize: 13, color: INK }}>Volume : <b>{volume.toFixed(2)} m³</b></span>
                  <span style={{ fontSize: 13, color: INK }}>Masse nécessaire : <b>{Math.round(masseNecessaire)} g</b></span>
                  <span style={{ fontSize: 13, color: INK }}>Masse effective : <b>{Math.round(masseEffective)} g</b></span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: suffisant ? "#2F7D3C" : "#B0473E" }}>{suffisant ? "✓ Suffisant" : "✗ Insuffisant"}</span>
                </div>

                <div style={{ fontWeight: 600, color: INK_2, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>Générateurs</div>
                <div className="overflow-x-auto mb-5">
                  <table className="w-full" style={{ fontSize: 13 }}>
                    <thead>
                      <tr style={{ color: MUTED, textAlign: "left" }}>
                        <th className="pb-1.5 font-medium">Désignation</th>
                        <th className="pb-1.5 font-medium text-right">Masse</th>
                        <th className="pb-1.5 font-medium text-right">Prix</th>
                        <th className="pb-1.5 font-medium text-right">Qté</th>
                      </tr>
                    </thead>
                    <tbody>
                      {fireproGenerateurs.map((g) => {
                        const qte = zone.generateurs[g.id] || 0;
                        return (
                          <tr key={g.id} style={{ borderTop: `1px solid ${LINE}`, background: qte > 0 ? "#FBF3E4" : "transparent" }}>
                            <td className="py-1.5 pr-3" style={{ color: INK }}>{g.label}</td>
                            <td className="py-1.5 text-right" style={{ color: MUTED }}>{g.masse} g</td>
                            <td className="py-1.5 text-right" style={{ color: MUTED, fontVariantNumeric: "tabular-nums" }}>{euros(g.prix)}</td>
                            <td className="py-1.5 text-right" style={{ width: 90 }}>
                              <NumberField value={qte} onChange={(v) => setGenerateurZone(zone.id, g.id, v)} suffix="u" width={64} />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {Object.entries(fireproAccessoires).map(([famId, fam]) => (
                  <div key={famId} className="mb-5">
                    <div style={{ fontWeight: 600, color: INK_2, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>{fam.label}</div>
                    <div className="overflow-x-auto">
                      <table className="w-full" style={{ fontSize: 13 }}>
                        <tbody>
                          {fam.items.map((item) => {
                            const qte = zone.accessoires[item.id] || 0;
                            return (
                              <tr key={item.id} style={{ borderBottom: `1px solid ${LINE}`, background: qte > 0 ? "#FBF3E4" : "transparent" }}>
                                <td className="py-1.5 pr-3" style={{ color: INK }}>{item.label}</td>
                                <td className="py-1.5 text-right" style={{ color: MUTED, fontVariantNumeric: "tabular-nums", width: 90 }}>{euros(item.prix)}</td>
                                <td className="py-1.5 text-right" style={{ width: 90 }}>
                                  <NumberField value={qte} onChange={(v) => setAccessoireZone(zone.id, item.id, v)} suffix="u" width={64} />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))}

                <div className="flex items-center justify-between pt-3" style={{ borderTop: `2px solid ${INK}` }}>
                  <span style={{ fontWeight: 700, color: INK }}>Sous-total {zone.nom}</span>
                  <span style={{ fontWeight: 800, color: INK, fontSize: 18, fontVariantNumeric: "tabular-nums" }}>{euros(montantTotal)}</span>
                </div>
              </SectionCard>
            ))}

            <button
              onClick={addZoneFirePro}
              className="flex items-center justify-center gap-1.5 px-4 py-3 rounded-lg text-sm font-medium border-2 border-dashed"
              style={{ borderColor: LINE, color: INK_2 }}
            >
              <Plus size={16} /> Ajouter un local
            </button>

            <SectionCard title="Récapitulatif FirePro" icon={FileText}>
              <div className="overflow-x-auto">
                <table className="w-full" style={{ fontSize: 13 }}>
                  <thead>
                    <tr style={{ color: MUTED, textAlign: "left" }}>
                      <th className="pb-2 font-medium">Local</th>
                      <th className="pb-2 font-medium text-right">Masse nécess.</th>
                      <th className="pb-2 font-medium text-right">Masse effective</th>
                      <th className="pb-2 font-medium text-right">Montant HT</th>
                    </tr>
                  </thead>
                  <tbody>
                    {zonesFireProCalc.map(({ zone, masseNecessaire, masseEffective, montantTotal, suffisant }) => (
                      <tr key={zone.id} style={{ borderTop: `1px solid ${LINE}` }}>
                        <td className="py-2" style={{ color: INK }}>
                          {zone.nom} {!suffisant && <span style={{ color: "#B0473E", fontSize: 11.5 }}>(insuffisant)</span>}
                        </td>
                        <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{Math.round(masseNecessaire)} g</td>
                        <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{Math.round(masseEffective)} g</td>
                        <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{euros(montantTotal)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex flex-col gap-2.5 mt-5" style={{ fontSize: 13.5 }}>
                <div className="flex items-center justify-between">
                  <span style={{ color: INK }}>Montant HT avant coefficient</span>
                  <span style={{ fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{euros(totalFireProAvantCoef)}</span>
                </div>
                <div className="flex items-center justify-between gap-3">
                  <span style={{ color: INK }}>Coefficient d'ajustement</span>
                  <NumberField value={fireproCoefAjustement} onChange={setFireproCoefAjustement} suffix="×" width={90} />
                </div>
                <div style={{ borderTop: `2px solid ${INK}`, marginTop: 6, paddingTop: 12 }} className="flex items-center justify-between">
                  <span style={{ fontWeight: 700, color: INK, fontSize: 16 }}>TOTAL HT FIREPRO</span>
                  <span style={{ fontWeight: 800, color: INK, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>{euros(totalFirePro)}</span>
                </div>
              </div>
            </SectionCard>
          </>
        )}

        {/* ---------------- ONGLET PARAMETRES ---------------- */}
        {tab === "parametres" && (
          <>
            <div className="flex items-center justify-between flex-wrap gap-2">
              <span style={{ fontSize: 12, color: MUTED }}>Vos modifications sont enregistrées automatiquement dans ce navigateur.</span>
              <div className="flex items-center gap-3 flex-wrap">
                <button onClick={enregistrerParametresMaintenant} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                  Enregistrer les valeurs
                </button>
                <button onClick={annulerModifications} className="text-sm underline" style={{ color: INK_2 }}>
                  Annuler mes modifications
                </button>
                <button onClick={reinitialiserParametres} className="text-sm underline" style={{ color: "#B0473E" }}>
                  Réinitialiser aux valeurs d'usine
                </button>
              </div>
            </div>

            <SectionCard title="Prix jour technicien" icon={Settings2}>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {Object.entries(tarifs).map(([key, t]) => (
                  <div key={key} style={{ border: `1px solid ${LINE}`, borderRadius: 8 }} className="p-4">
                    <div style={{ fontWeight: 600, color: INK, marginBottom: 10 }}>{t.label}</div>
                    <div className="flex flex-col gap-2">
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Prix / jour</label>
                      <NumberField value={t.jour} onChange={(v) => setTarifs((s) => ({ ...s, [key]: { ...s[key], jour: v } }))} suffix="€" width="100%" />
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Prix / demi-journée (jour ÷ 2 × 1,2)</label>
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
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>{m.label}</label>
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
                                <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Simple (1-2)</label>
                                <NumberField value={item.heuresSimple} onChange={(v) => majItem({ heuresSimple: v })} suffix="h" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Complexe (3-4)</label>
                                <NumberField value={item.heuresComplexe} onChange={(v) => majItem({ heuresComplexe: v })} suffix="h" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Préparation</label>
                                <NumberField value={item.heuresPrepa} onChange={(v) => majItem({ heuresPrepa: v })} suffix="h" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Amortissement</label>
                                <NumberField value={item.amort} onChange={(v) => majItem({ amort: v })} suffix="€" width="100%" />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Niveau par défaut</label>
                                <Select
                                  value={item.niveauDefaut || "complet"}
                                  onChange={(v) => majItem({ niveauDefaut: v })}
                                  options={Object.entries(NIVEAUX_PRESTATION).map(([k, label]) => ({ value: k, label }))}
                                />
                              </div>
                              <div>
                                <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Prix de revient (défaut)</label>
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

            <SectionCard title="Catalogue — composants & sous-traitance (coefficient sur prix d'achat)" icon={ClipboardList}>
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

            <SectionCard title="Analyses d'huile, TGBT, Batteries & Onduleurs (prix directs)" icon={ClipboardList}>
              <div className="flex flex-col gap-5">
                {Object.entries(catalogueDirect).map(([key, cat]) => (
                  <div key={key}>
                    <div className="flex items-center justify-between mb-2">
                      <div style={{ fontWeight: 600, color: INK, fontSize: 13 }}>{cat.label}</div>
                      <button onClick={() => ajouterLigneCatalogueDirect(key)} className="flex items-center gap-1 text-xs font-medium" style={{ color: INK_2 }}>
                        <Plus size={13} /> Ajouter une ligne
                      </button>
                    </div>
                    {key === "batteries" ? (
                      <div className="flex flex-col gap-1.5">
                        <div className="grid grid-cols-6 gap-2" style={{ fontSize: 10.5, color: MUTED, textTransform: "uppercase" }}>
                          <span></span>
                          <span>Prix d'achat</span>
                          <span>Coefficient</span>
                          <span colSpan={2}>Prix de vente</span>
                          <span></span>
                        </div>
                        {cat.items.map((item, i) => {
                          const majBatterie = (patch) => {
                            const next = { ...item, ...patch };
                            next.prix = next.prixAchat * next.coef;
                            setCatalogueDirect((s) => ({ ...s, batteries: { ...s.batteries, items: s.batteries.items.map((it, j) => (j === i ? next : it)) } }));
                          };
                          return (
                            <div key={item.id} className="grid grid-cols-6 gap-2 items-center">
                              <TextField value={item.label} onChange={(v) => renommerLigneCatalogueDirect(key, item.id, v)} />
                              <NumberField value={item.prixAchat} onChange={(v) => majBatterie({ prixAchat: v })} suffix="€" width="100%" />
                              <NumberField value={item.coef} onChange={(v) => majBatterie({ coef: v })} suffix="×" width="100%" />
                              <span style={{ fontSize: 13, fontWeight: 600, color: INK, fontVariantNumeric: "tabular-nums", gridColumn: "span 2" }}>{euros(item.prix)}</span>
                              <button onClick={() => supprimerLigneCatalogueDirect(key, item.id)} style={{ color: "#B0473E" }} className="flex justify-end">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="flex flex-col gap-1.5">
                        {cat.items.map((item, i) => (
                          <div key={item.id} className="grid grid-cols-4 gap-3 items-center">
                            <div style={{ gridColumn: "span 2" }}>
                              <TextField value={item.label} onChange={(v) => renommerLigneCatalogueDirect(key, item.id, v)} />
                            </div>
                            <NumberField
                              value={item.prix}
                              onChange={(v) => setCatalogueDirect((s) => ({ ...s, [key]: { ...s[key], items: s[key].items.map((it, j) => (j === i ? { ...it, prix: v } : it)) } }))}
                              suffix="€"
                            />
                            <button onClick={() => supprimerLigneCatalogueDirect(key, item.id)} style={{ color: "#B0473E" }} className="flex justify-end">
                              <Trash2 size={14} />
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="FirePro — Classes de feu (densité DAD, g/m³)" subtitle="Masse nécessaire = volume × DAD × coef. remplissage × coef. confinement" icon={Flame}>
              <div className="flex flex-col gap-1.5">
                {fireproClasses.map((c, i) => (
                  <div key={c.id} className="grid grid-cols-3 gap-3 items-center">
                    <span style={{ fontSize: 13, color: INK, gridColumn: "span 2" }}>{c.label}</span>
                    <NumberField value={c.dad} onChange={(v) => setFireproClasses((arr) => arr.map((x, j) => (j === i ? { ...x, dad: v } : x)))} suffix="g/m³" />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="FirePro — Coefficients de confinement" icon={Flame}>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {Object.entries(fireproConfinement).map(([key, c]) => (
                  <div key={key}>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>{c.label}</label>
                    <NumberField value={c.coef} onChange={(v) => setFireproConfinement((s) => ({ ...s, [key]: { ...s[key], coef: v } }))} suffix="×" width="100%" />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard
              title="FirePro — Générateurs (désignation, masse, prix)"
              icon={Flame}
              right={
                <button onClick={ajouterGenerateurFirePro} className="flex items-center gap-1 text-xs font-medium" style={{ color: INK_2 }}>
                  <Plus size={13} /> Ajouter un générateur
                </button>
              }
            >
              <div className="overflow-x-auto">
                <div className="flex flex-col gap-1.5">
                  <div className="grid grid-cols-5 gap-3" style={{ fontSize: 10.5, color: MUTED, textTransform: "uppercase" }}>
                    <span style={{ gridColumn: "span 2" }}></span>
                    <span>Masse (g)</span>
                    <span>Prix (€)</span>
                    <span></span>
                  </div>
                  {fireproGenerateurs.map((g, i) => (
                    <div key={g.id} className="grid grid-cols-5 gap-3 items-center">
                      <div style={{ gridColumn: "span 2" }}>
                        <TextField value={g.label} onChange={(v) => renommerGenerateurFirePro(g.id, v)} />
                      </div>
                      <NumberField value={g.masse} onChange={(v) => setFireproGenerateurs((arr) => arr.map((x, j) => (j === i ? { ...x, masse: v } : x)))} width="100%" />
                      <NumberField value={g.prix} onChange={(v) => setFireproGenerateurs((arr) => arr.map((x, j) => (j === i ? { ...x, prix: v } : x)))} suffix="€" width="100%" />
                      <button onClick={() => supprimerGenerateurFirePro(g.id)} style={{ color: "#B0473E" }} className="flex justify-end">
                        <Trash2 size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </SectionCard>

            <SectionCard title="FirePro — Centrales, détection, relayage, fixations, main d'œuvre" icon={Flame}>
              <div className="flex flex-col gap-5">
                {Object.entries(fireproAccessoires).map(([famId, fam]) => (
                  <div key={famId}>
                    <div className="flex items-center justify-between mb-2">
                      <div style={{ fontWeight: 600, color: INK, fontSize: 13 }}>{fam.label}</div>
                      <button onClick={() => ajouterAccessoireFirePro(famId)} className="flex items-center gap-1 text-xs font-medium" style={{ color: INK_2 }}>
                        <Plus size={13} /> Ajouter une ligne
                      </button>
                    </div>
                    <div className="flex flex-col gap-1.5">
                      {fam.items.map((item, i) => (
                        <div key={item.id} className="grid grid-cols-4 gap-3 items-center">
                          <div style={{ gridColumn: "span 2" }}>
                            <TextField value={item.label} onChange={(v) => renommerAccessoireFirePro(famId, item.id, v)} />
                          </div>
                          <NumberField
                            value={item.prix}
                            onChange={(v) =>
                              setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: s[famId].items.map((it, j) => (j === i ? { ...it, prix: v } : it)) } }))
                            }
                            suffix="€"
                            width="100%"
                          />
                          <button onClick={() => supprimerAccessoireFirePro(famId, item.id)} style={{ color: "#B0473E" }} className="flex justify-end">
                            <Trash2 size={14} />
                          </button>
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
