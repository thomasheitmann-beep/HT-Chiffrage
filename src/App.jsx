import React, { useState, useEffect, useMemo, useRef } from "react";
import { Plus, Trash2, Settings2, FileText, ClipboardList, Download, Copy, Flame, Upload } from "lucide-react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, WidthType, ShadingType } from "docx";
import { doc, getDoc, setDoc, collection, getDocs, addDoc, deleteDoc, writeBatch } from "firebase/firestore";
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
const FIREPRO_DEVIS_COLLECTION = "ht-chiffrage-firepro-devis";
// Collection racine partagée entre TOUTES les apps HT Maintenance (pas propre
// à cette app) : un document par contact, même forme et mêmes fonctions que
// sur les autres apps (devis-facture, consignation...).
const CONTACTS_COLLECTION = "contacts-ht-maintenance";

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

// ---------------------------------------------------------------------------
// Contacts — base partagée entre toutes les apps HT Maintenance (même projet
// Firebase). Un document par contact dans "contacts-ht-maintenance" ; modifier
// un contact ici le modifie partout.
// ---------------------------------------------------------------------------

async function loadContactsShared() {
  const snap = await getDocs(collection(db, CONTACTS_COLLECTION));
  return snap.docs.map((d) => ({ id: d.id, ...d.data() }));
}

// Ne réécrit que ce qui a changé (ajouts/modifications/suppressions), par
// lots de 450 opérations max (limite Firestore : 500 par batch).
async function syncContactsToShared(newClients, previousClients) {
  const prevMap = new Map(previousClients.map((c) => [c.id, c]));
  const newIds = new Set(newClients.map((c) => c.id));
  const ops = [];
  newClients.forEach((c) => {
    const prev = prevMap.get(c.id);
    if (!prev || JSON.stringify(prev) !== JSON.stringify(c)) ops.push({ type: "set", id: c.id, data: c });
  });
  previousClients.forEach((c) => {
    if (!newIds.has(c.id)) ops.push({ type: "delete", id: c.id });
  });
  for (let i = 0; i < ops.length; i += 450) {
    const batch = writeBatch(db);
    ops.slice(i, i + 450).forEach((op) => {
      const ref = doc(db, CONTACTS_COLLECTION, op.id);
      if (op.type === "set") {
        const { id, ...data } = op.data;
        batch.set(ref, data);
      } else {
        batch.delete(ref);
      }
    });
    await batch.commit();
  }
}

// Libellé d'une ligne libre : libellé saisi (poste manuel / main d'œuvre) ou nom de la famille.
function libelleLigneLibre(l) {
  if (l.famille === "manuel" || l.famille === "maindoeuvre") return l.libelleManuel || LABEL_FAMILLE[l.famille] || "Poste libre";
  return LABEL_FAMILLE[l.famille] || l.famille;
}

function nouveauContactVierge() {
  return { id: uid(), societe: "", contact: "", adresse: "", cp: "", ville: "", pays: "", email: "", telephones: [] };
}

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

// Extrait le premier nombre trouvé dans un libellé (ex. "45 Ah" -> 45, "60°C"
// -> 60) pour trier les catalogues par ordre croissant. Renvoie Infinity si
// aucun nombre n'est trouvé (l'élément reste en fin de liste).
function extraireNombreLabel(label) {
  const m = String(label).match(/[\d.,]+/);
  return m ? parseFloat(m[0].replace(",", ".")) : Infinity;
}
function trierParLabel(items) {
  return [...items].sort((a, b) => extraireNombreLabel(a.label) - extraireNombreLabel(b.label));
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

// Petits équivalents locaux des styles / helpers de l'application de maintenance, pour réutiliser
// le même composant de catalogue à l'identique.
const CAT_BLUE = "#0A5DA8";
const catInput = { border: "1px solid #DCE0E3", borderRadius: 6, background: "#fff", color: "#1B2733", outline: "none" };
const catBtn = (c) => ({ display: "inline-flex", alignItems: "center", gap: 5, padding: "5px 10px", borderRadius: 6, border: `1px solid ${c}`, color: c, background: "transparent", fontSize: 12, fontWeight: 600, cursor: "pointer" });
const catNum = (v) => { if (v === null || v === undefined || String(v).trim() === "") return null; const n = Number(String(v).replace(",", ".")); return Number.isFinite(n) ? n : null; };
// ---------------------------------------------------------------------------------------------
// Catalogue d'aide au choix des pièces d'usure (condensateurs AC / DC, cartes condensateurs,
// ventilateurs) — extrait du fichier de codification fourni par l'utilisateur (codes internes du
// fichier d'origine non repris). r = [[marque, référence], …équivalences] ; c/cv = capacité (texte /
// µF) ; t/tv = tension (texte / V) ; d = dimensions (Ø×L ou L×l×h) ; dd/dl = Ø et longueur (mm) ;
// cx = connexion (I = insert, M = tige filetée, B = bornes, Faston) ; fx = fixation ; n = remarque.
// Cartes : m/r = marque/référence de la carte, v = variante, c = condensateurs montés (q × catalogue dc[i]).
const CATALOGUE_PIECES = {"ac":[{"r":[["KEMET","C44AMLP5200ZU0J"]],"c":"20µF","cv":20.0,"t":"850Vdc","tv":850.0,"d":"65*101","dd":65.0,"dl":101.0,"cx":"M6","fx":"M12"},{"r":[["DUCATI","DUC416101564"]],"c":"10µF","cv":10.0,"t":"400Vdc","tv":400.0,"d":"36*58","dd":36.0,"dl":58.0,"cx":"double Faston","fx":"M8"},{"r":[["ICAR","LNFP3Y10030"]],"c":"100µF","cv":100.0,"n":"ICAR liquidation"},{"r":[["DUCATI","DUC416843290"]],"c":"100µF","cv":100.0,"t":"330Vac","tv":330.0,"d":"55*150","dd":55.0,"dl":150.0,"cx":"M5","fx":"M8"},{"r":[["ELECTRONICON","E62G12104G10"]],"c":"100µF","cv":100.0,"t":"420V","tv":420.0,"d":"50*124","dd":50.0,"dl":124.0,"cx":"M5","fx":"M8"},{"r":[["ICAR","MFR21000"]],"c":"100µF","cv":100.0,"t":"400V","tv":400.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["EPCOS","B32362A3107J030"]],"c":"100µF","cv":100.0,"t":"330Vac","tv":330.0,"d":"75*117","dd":75.0,"dl":117.0,"cx":"M10","fx":"M12"},{"r":[["ITALFARAD","MFR2MT"]],"c":"100µF","cv":100.0,"t":"450V","tv":450.0,"d":"60*138","dd":60.0,"dl":138.0},{"r":[["KEMET","C44AFGP6100ZG0J"]],"c":"100µF","cv":100.0,"t":"250Vac","tv":250.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"M6","fx":"M12"},{"r":[["KEMET","C930ZG36100ZA0X"]],"c":"100µF","cv":100.0,"t":"330Vac","tv":330.0,"d":"60*133","dd":60.0,"dl":133.0,"cx":"double Faston","fx":"M12"},{"r":[["DUCATI","DUC416840068"]],"c":"100µF","cv":100.0,"t":"330Vac","tv":330.0,"d":"55*150","dd":55.0,"dl":150.0,"cx":"M6","fx":"M12"},{"r":[["ICAR","MC2100"]],"c":"100µF","cv":100.0,"t":"250V","tv":250.0,"d":"60*122","dd":60.0,"dl":122.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["KEMET","C44PLGR6100AASJ"]],"c":"100µF","cv":100.0,"t":"330Vac","tv":330.0,"d":"65*115","dd":65.0,"dl":115.0,"cx":"M10","fx":"M12","n":"doublon CPS100330 plus petit"},{"r":[["ARCOTRONICS","C44HFGP6100ZASJ"],["EPCOS","B32361A2107J050"]],"c":"100µF","cv":100.0,"t":"300Vac","tv":300.0,"d":"63,5*10","dd":63.5,"dl":10.0,"cx":"M6","fx":"M12"},{"r":[["DUCATI","DUC416845618"]],"c":"100µF","cv":100.0,"t":"550Vac","tv":550.0,"d":"75*210","dd":75.0,"dl":210.0,"cx":"M10","fx":"M12","n":"Mise a jour des données suite mail de Christophe 02/03/21"},{"r":[["ICAR","CRM2511A550415"]],"c":"101.7µF","cv":101.7,"t":"415V","tv":415.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["KEMET","C44PZGR6120ZCSJ"],["EPCOS","B32362A3127J030"]],"c":"120µF","cv":120.0,"t":"300Vac","tv":300.0,"d":"75*117","dd":75.0,"dl":117.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44AFGP6130ZA0J"]],"c":"130µF","cv":130.0,"t":"250Vac","tv":250.0,"d":"65*138","dd":65.0,"dl":138.0,"cx":"M6","fx":"M12"},{"r":[["EPCOS","B32361A2137J050"]],"c":"130µF","cv":130.0,"t":"250Vac","tv":250.0,"d":"65*110","dd":65.0,"dl":110.0},{"r":[["ICAR","LNFP3Y15030"]],"c":"150µF","cv":150.0,"t":"300V","tv":300.0,"d":"75*140","dd":75.0,"dl":140.0,"n":"ICAR liquidation"},{"r":[["EPCOS","B32361A2157J050"]],"c":"150µF","cv":150.0,"t":"250Vac","tv":250.0,"d":"63,5*117","dd":63.5,"dl":117.0,"cx":"M6","fx":"M12"},{"r":[["KEMET","C44AFGP6150ZA0J"]],"c":"150µF","cv":150.0,"t":"250Vac","tv":250.0,"d":"70*138","dd":70.0,"dl":138.0,"cx":"M6","fx":"M12"},{"r":[["ICAR","WB40160F"]],"c":"16µF","cv":16.0,"t":"450V","tv":450.0,"d":"35*71","dd":35.0,"dl":71.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["ICAR","WB4025F"]],"c":"2.5µF","cv":2.5,"t":"400V","tv":400.0,"d":"30*51","dd":30.0,"dl":51.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["KEMET","C44AMLP5200ZU"]],"c":"20µF","cv":20.0,"t":"850Vdc","tv":850.0},{"r":[["ICAR","WB40200F"]],"c":"20µF","cv":20.0,"t":"450V","tv":450.0,"d":"40*71","dd":40.0,"dl":71.0,"cx":"double Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["DUCATI","DUC416102564"]],"c":"20µF","cv":20.0,"t":"450V","tv":450.0,"d":"40*70","dd":40.0,"dl":70.0},{"r":[["COMAR","8252385"]],"c":"200µF","cv":200.0,"t":"250V","tv":250.0,"d":"75*165","dd":75.0,"dl":165.0},{"r":[["DUCATI","DUC416843637"]],"c":"200µF","cv":200.0,"t":"330Vac","tv":330.0,"d":"75*180","dd":75.0,"dl":180.0,"cx":"M10","fx":"M12"},{"r":[["ICAR","LNFP3B20030"]],"c":"200µF","cv":200.0,"n":"ICAR liquidation"},{"r":[["KEMET","C44PLGR6200ZASJ"]],"c":"200µF","cv":200.0,"t":"330Vac","tv":330.0,"d":"65*145","dd":65.0,"dl":145.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44AFGP6200ZE0J"]],"c":"200µF","cv":200.0,"t":"250Vac","tv":250.0,"d":"76*138","dd":76.0,"dl":138.0,"cx":"M6","fx":"M12"},{"r":[["EPCOS","B32361B2207J050"]],"c":"200µF","cv":200.0,"t":"250Vac","tv":250.0,"d":"65*142","dd":65.0,"dl":142.0,"cx":"M6","fx":"M12"},{"r":[["KEMET","C44EFGR6200ZB0J"]],"c":"200µF","cv":200.0,"t":"250Vac","tv":250.0,"d":"85*145","dd":85.0,"dl":145.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44AFGR6200ZE0J"]],"c":"200µF","cv":200.0,"t":"250Vac","tv":250.0,"d":"76*138","dd":76.0,"dl":138.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44PKGR6200AASJ"]],"c":"200µF","cv":200.0,"t":"440Vac","tv":440.0,"d":"75*247","dd":75.0,"dl":247.0,"cx":"M10","fx":"M12"},{"r":[["DUCATI","DUC416740007"]],"c":"200µF","cv":200.0,"t":"250Vac","tv":250.0},{"r":[["DUCATI","DUC416747054"]],"c":"200µF","cv":200.0,"t":"330Vac","tv":330.0,"d":"75*180","dd":75.0,"dl":180.0,"cx":"M10","fx":"M12"},{"r":[["ICAR","bioenergy A11B24017"]],"c":"240µF","cv":240.0,"t":"170V","tv":170.0,"d":"80*249","dd":80.0,"dl":249.0,"cx":"Mx","fx":"M12","n":"ICAR liquidation"},{"r":[["ICAR","WB40250"]],"c":"25µF","cv":25.0,"t":"450V","tv":450.0,"d":"45*71","dd":45.0,"dl":71.0,"cx":"double Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["KEMET","C44AFGR6250ZA0J"]],"c":"250µF","cv":250.0,"t":"250Vac","tv":250.0,"d":"75*150","dd":75.0,"dl":150.0,"cx":"M10","fx":"M12"},{"r":[["VISHAY","PhMKP4403281084"]],"c":"3*154µF","t":"440V","tv":440.0},{"r":[["VISHAY","PhMKP4003250084"]],"c":"3*165.8µF","t":"400Vac","tv":400.0,"d":"85*265","dd":85.0,"dl":265.0,"cx":"B25","fx":"M12"},{"r":[["EPCOS","B32366A4497J320"],["EPCOS","C9TS5LD6166AA0J"]],"c":"3*166µF","t":"400V","tv":400.0,"d":"85*290","dd":85.0,"dl":290.0},{"r":[["VISHAY","PhMKPg4206003"]],"c":"3*192.5µF","t":"420Vac","tv":420.0,"d":"87*H","dd":87.0,"cx":"M5","fx":"M12"},{"r":[["KEMET","C9TS5LD620AA0J"]],"c":"3*200µF","t":"400V","tv":400.0},{"r":[["EPCOS","B25836B5117A305"]],"c":"3*32.1µF","t":"525V","tv":525.0},{"r":[["KEMET","C44EKMQ5999ZB0J"]],"c":"3*33µF","t":"750V","tv":750.0,"cx":"M8","fx":"Mx"},{"r":[["EPCOS","B25667C5966A375"]],"c":"3*38.5µF","t":"525V","tv":525.0},{"r":[["FRAKO","LKT141440D52"]],"c":"3*77.3µF"},{"r":[["EPCOS","B25667C5237A375"]],"c":"3*77µF","t":"525Vac","tv":525.0,"d":"116*200","dd":116.0,"dl":200.0,"cx":"B25","fx":"M12"},{"r":[["VISHAY","GTO53030R"]],"c":"30µF","cv":30.0,"t":"530V","tv":530.0,"d":"87*62","dd":87.0,"dl":62.0,"cx":"Mx","fx":"Mx"},{"r":[["ICAR","WB40300F"]],"c":"30µF","cv":30.0,"t":"400V","tv":400.0,"d":"45*71","dd":45.0,"dl":71.0,"cx":"Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["ICAR","MLR25L45300"]],"c":"30µF","cv":30.0,"t":"400V","tv":400.0,"d":"45*128","dd":45.0,"dl":128.0,"cx":"double Faston","n":"ICAR liquidation"},{"r":[["VISHAY","GTO 530-30R"]],"c":"30µF","cv":30.0,"t":"530V","tv":530.0,"d":"87*62","dd":87.0,"dl":62.0,"cx":"Mx","fx":"Mx"},{"r":[["KEMET","C44EFGR6300ZA0J"]],"c":"300µF","cv":300.0,"t":"250Vac","tv":250.0,"d":"85*230","dd":85.0,"dl":230.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44PLGR6300ZASJ"]],"c":"300µF","cv":300.0,"t":"330Vac","tv":330.0,"d":"65*247","dd":65.0,"dl":247.0,"cx":"Mx","fx":"M12"},{"r":[["DUCATI","DUC416843754"]],"c":"300µF","cv":300.0,"t":"330Vac","tv":330.0,"d":"85*180","dd":85.0,"dl":180.0,"cx":"M10","fx":"M12"},{"r":[["DUCATI","DUC416335164"]],"c":"31.5µF","cv":31.5,"t":"475V","tv":475.0,"d":"45*122","dd":45.0,"dl":122.0,"cx":"double Faston","fx":"Mx"},{"r":[["COMAR","MKAS31450500"]],"c":"31µF","cv":31.0,"t":"450V","tv":450.0},{"r":[["ICAR","CRM2511D166400"]],"c":"33.3µF","cv":33.3,"t":"400V","tv":400.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"M12","n":"ICAR liquidation"},{"r":[["ITALFARAD","MFA33344"]],"c":"33.3µF","cv":33.3,"t":"400V","tv":400.0,"d":"55*128","dd":55.0,"dl":128.0},{"r":[["VISHAY","KMKP60033IBR"]],"c":"33µF","cv":33.0,"t":"600V","tv":600.0,"d":"64*135","dd":64.0,"dl":135.0},{"r":[["EPCOS","B32362A2337J050"]],"c":"330µF","cv":330.0,"t":"250Vac","tv":250.0,"d":"75*190","dd":75.0,"dl":190.0},{"r":[["ICAR","CRM2511333550"]],"c":"35µF","cv":35.0,"t":"550V","tv":550.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["DUCATI","DUC416463310"]],"c":"3x131µF","t":"450V","tv":450.0,"d":"116*255","dd":116.0,"dl":255.0},{"r":[["DUCATI","DUC416463360"]],"c":"3x157µF","t":"450Vac","tv":450.0,"d":"116*290","dd":116.0,"dl":290.0},{"r":[["ICAR","CRTE13620840052"]],"c":"3x154µF","d":"136*208","dd":136.0,"dl":208.0,"n":"ICAR liquidation"},{"r":[["ICAR","WB40400F"]],"c":"40µF","cv":40.0,"t":"400V","tv":400.0,"d":"45*91","dd":45.0,"dl":91.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["DUCATI","DUC416843994"],["DUCATI","DUC416843998"]],"c":"400µF","cv":400.0,"t":"330Vac","tv":330.0,"d":"90*210","dd":90.0,"dl":210.0,"cx":"M10","fx":"M12"},{"r":[["DUCATI","DUC416747053"],["EPCOS","B32362A3407J030"]],"c":"400µF","cv":400.0,"t":"330Vac","tv":330.0,"d":"85*280","dd":85.0,"dl":280.0,"cx":"M10","fx":"M12"},{"r":[["EPCOS","B32362A2407J050"]],"c":"400µF","cv":400.0,"t":"250Vac","tv":250.0,"d":"85*197","dd":85.0,"dl":197.0,"cx":"M10","fx":"Mx"},{"r":[["KEMET","C44PLGR6400ZASJ"]],"c":"400µF","cv":400.0,"t":"330Vac","tv":330.0,"d":"65*265","dd":65.0,"dl":265.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44EFGR6400ZA0J"]],"c":"400µF","cv":400.0,"t":"250Vac","tv":250.0,"d":"85*270","dd":85.0,"dl":270.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44HHG35450ZASJ"]],"c":"45µF","cv":45.0,"t":"330Vac","tv":330.0,"d":"50*90","dd":50.0,"dl":90.0,"cx":"double Faston","fx":"M12"},{"r":[["KEMET","C44HMG35450ZASJ"]],"c":"45µF","cv":45.0,"t":"800V","tv":800.0,"d":"50*90","dd":50.0,"dl":90.0,"cx":"double Faston","fx":"M12"},{"r":[["VISHAY","KMKP60047IBR"]],"c":"47µF","cv":47.0,"t":"600V","tv":600.0,"d":"84*125","dd":84.0,"dl":125.0},{"r":[["ICAR","CRM251125400"]],"c":"50µF","cv":50.0,"t":"400V","tv":400.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["COMAR","MKAS50550"]],"c":"50µF","cv":50.0,"t":"550V","tv":550.0},{"r":[["ITALFARAD","MFR1MT"]],"c":"50µF","cv":50.0,"t":"280Vac","tv":280.0,"d":"50*103","dd":50.0,"dl":103.0,"cx":"double Faston","fx":"M8"},{"r":[["EPCOS","B32332I6506J082"]],"c":"50µF","cv":50.0,"t":"450Vac","tv":450.0,"d":"45*103","dd":45.0,"dl":103.0,"cx":"double Faston","fx":"M8"},{"r":[["EPCOS","B32361A2506J050"]],"c":"50µF","cv":50.0,"t":"250Vac","tv":250.0,"d":"63,5*70","dd":63.5,"dl":70.0,"cx":"M6","fx":"M12"},{"r":[["EPCOS","B32361A2506J051"]],"c":"50µF","cv":50.0,"t":"250Vac","tv":250.0,"d":"63,5*70","dd":63.5,"dl":70.0,"cx":"M6","fx":"M12"},{"r":[["ICAR","MLR25L45500"]],"c":"50µF","cv":50.0,"t":"450V","tv":450.0,"d":"55*128","dd":55.0,"dl":128.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["ITALFARAD","MFR2MT"]],"c":"50µF","cv":50.0,"t":"450V","tv":450.0,"d":"50*103","dd":50.0,"dl":103.0,"cx":"double Faston","fx":"Mx"},{"r":[["ICAR","WB40500F"],["DUCATI","DUC416103064"]],"c":"50µF","cv":50.0,"t":"450V","tv":450.0,"d":"45*116","dd":45.0,"dl":116.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["KEMET","C44AFFP5500ZE0J"]],"c":"50µF","cv":50.0,"t":"250Vac","tv":250.0,"d":"50*100","dd":50.0,"dl":100.0,"cx":"M6","fx":"M8"},{"r":[["ITALFARAD","MFR2"]],"c":"50µF","cv":50.0,"t":"400V","tv":400.0,"d":"50*128","dd":50.0,"dl":128.0,"cx":"double Faston","fx":"Mx"},{"r":[["DUCATI","DUC416843158"]],"c":"50µF","cv":50.0,"t":"330Vac","tv":330.0,"d":"55*115","dd":55.0,"dl":115.0,"cx":"M6","fx":"M12"},{"r":[["DUCATI","DUC416272464"]],"c":"50µF","cv":50.0,"t":"330Vac","tv":330.0,"d":"45*98","dd":45.0,"dl":98.0,"cx":"Faston","fx":"M8"},{"r":[["EPCOS","B32360A3506J030"],["EPCOS","B32360S3506J030"]],"c":"50µF","cv":50.0,"t":"330Vac","tv":330.0,"d":"53*107","dd":53.0,"dl":107.0,"cx":"double Faston","fx":"M8"},{"r":[["KEMET","C44PFGR6500AA0J"]],"c":"500µF","cv":500.0,"t":"250Vac","tv":250.0,"d":"85*270","dd":85.0,"dl":270.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44PLGR6500ZASJ"]],"c":"500µF","cv":500.0,"t":"330Vac","tv":330.0,"d":"75*247","dd":75.0,"dl":247.0,"cx":"Mx","fx":"M12"},{"r":[["ITALFARAD","MFA54850"]],"c":"54.8µF","cv":54.8,"t":"440Vac","tv":440.0,"d":"60*138","dd":60.0,"dl":138.0},{"r":[["DUCATI","DUC416108264"]],"c":"6µF","cv":6.0,"t":"450V","tv":450.0,"d":"30*55","dd":30.0,"dl":55.0},{"r":[["ICAR","MLR25L40400"]],"c":"7.5µF","cv":7.5,"t":"400Vac","tv":400.0,"n":"ICAR liquidation"},{"r":[["ICAR","MLR25L45600"]],"c":"60µF","cv":60.0,"t":"400V","tv":400.0,"d":"55*128","dd":55.0,"dl":128.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["ICAR","WB40600F"]],"c":"60µF","cv":60.0,"t":"400V","tv":400.0,"d":"45*116","dd":45.0,"dl":116.0,"cx":"double Faston","fx":"M8","n":"ICAR liquidation"},{"r":[["ITALFARAD","MFR4"]],"c":"60µF","cv":60.0,"t":"450V","tv":450.0,"d":"55*128","dd":55.0,"dl":128.0,"cx":"double Faston","fx":"M8"},{"r":[["ITALFARAD","MFR4"]],"c":"60µF","cv":60.0,"t":"450V","tv":450.0,"d":"55*128","dd":55.0,"dl":128.0,"cx":"double Faston","fx":"M12"},{"r":[["DUCATI","DUC416335864"]],"c":"60µF","cv":60.0,"t":"475V","tv":475.0,"d":"55*122","dd":55.0,"dl":122.0,"cx":"double Faston","fx":"M12"},{"r":[["DUCATI","DUC416175664"]],"c":"60µF","cv":60.0,"t":"400V","tv":400.0,"d":"50*117","dd":50.0,"dl":117.0,"cx":"double Faston","fx":"Mx"},{"r":[["ICAR","CRM2511D55550"],["ICAR","CRM2511A obsolète"]],"c":"57.8µF","cv":57.8,"t":"550Vac","tv":550.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"M12","n":"ICAR liquidation"},{"r":[["KEMET","C44PFGR6600AA0J"]],"c":"600µF","cv":600.0,"t":"250Vac","tv":250.0,"d":"95*280","dd":95.0,"dl":280.0,"cx":"M10","fx":"M12"},{"r":[["KEMET","C44PLGR6600ZASJ"]],"c":"600µF","cv":600.0,"t":"330Vac","tv":330.0,"d":"85*270","dd":85.0,"dl":270.0,"cx":"M10","fx":"M12","n":"modiF L a la plce de F"},{"r":[["ITALFARAD","MFA66644"]],"c":"66.6µF","cv":66.6,"t":"440V","tv":440.0,"d":"60*138","dd":60.0,"dl":138.0},{"r":[["ICAR","CRM251133400"]],"c":"66.6µF","cv":66.6,"t":"400V","tv":400.0,"d":"60*138","dd":60.0,"dl":138.0,"cx":"double Faston","fx":"M12","n":"ICAR liquidation"},{"r":[["COMAR","MKAS666415540"]],"c":"66.6µF","cv":66.6,"t":"440V","tv":440.0,"d":"60*185","dd":60.0,"dl":185.0,"cx":"Faston","fx":"M12"},{"r":[["ICAR","MLR25U"]],"c":"70µF","cv":70.0,"t":"280V","tv":280.0,"d":"55*128","dd":55.0,"dl":128.0,"cx":"double Faston","fx":"Mx","n":"ICAR liquidation"},{"r":[["KEMET","C44AFGP5750ZA0J"]],"c":"75µF","cv":75.0,"t":"250V","tv":250.0,"d":"60*100","dd":60.0,"dl":100.0,"cx":"M6","fx":"M12"},{"r":[["COMAR","MKAS77500"]],"c":"77µF","cv":77.0,"t":"500V","tv":500.0},{"r":[["COMAR","DMP77 550 code 8490767"]],"c":"77µF","cv":77.0,"t":"550Vac","tv":550.0,"d":"65*185","dd":65.0,"dl":185.0,"cx":"double Faston","fx":"M12"},{"r":[["COMAR","MCZS"]],"c":"77µF","cv":77.0,"t":"500V","tv":500.0},{"r":[["EPCOS","B32340C4052A000"]],"c":"99µF","cv":99.0,"t":"400Vac","tv":400.0,"d":"63,5*142","dd":63.5,"dl":142.0,"cx":"double Faston","fx":"M12"},{"r":[["ICAR","CRT66V5525"]],"n":"ICAR liquidation"},{"r":[["FRAKO","LKT10525DPL"]],"t":"525V","tv":525.0,"d":"70*225","dd":70.0,"dl":225.0},{"r":[["FRAKO","LKT72480DL"]],"t":"480V","tv":480.0,"d":"60*225","dd":60.0,"dl":225.0},{"r":[["EPCOS","B32362C3207J030"]],"c":"200µF","cv":200.0,"t":"330Vac","tv":330.0,"d":"85*152","dd":85.0,"dl":152.0,"cx":"M10","fx":"M12","n":"B32362x3207j030 => x version du composant"},{"r":[["EPCOS","B32362C3307J030"]],"c":"300µF","cv":300.0,"t":"330Vac","tv":330.0,"d":"85*197","dd":85.0,"dl":197.0,"cx":"M10","fx":"M12","n":"B32362x3307j030 => x version du composant"},{"r":[["DUCATI","DUC416843991"]],"c":"500µF","cv":500.0,"t":"330Vac","tv":330.0,"cx":"M10","fx":"M12"}],"dc":[{"r":[["KENDEIL","K012501030HM0J143"]],"c":"10000µF","cv":10000.0,"t":"250V","tv":250.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014001030HM0J220"]],"c":"10000µF","cv":10000.0,"t":"400V","tv":400.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014001030HM0J214"]],"c":"10000µF","cv":10000.0,"t":"400V","tv":400.0,"d":"76*214","dd":76.0,"dl":214.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014501030HM0J143"]],"c":"10000µF","cv":10000.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014501030HM0J220"]],"c":"10000µF","cv":10000.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR102M450BC1"]],"c":"1000µF","cv":1000.0,"t":"450V","tv":450.0,"d":"51*105","dd":51.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR332M350DF1"],["KENDEIL","K013503320HM0J143"]],"c":"3300µF","cv":3300.0,"t":"350V","tv":350.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014502220HM0J093"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"76*93","dd":76.0,"dl":93.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR222M450DC1"],["KENDEIL","K014502220HM0J105"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"76*105","dd":76.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K013503320HM0J105"]],"c":"3300µF","cv":3300.0,"t":"350V","tv":350.0,"d":"76*105","dd":76.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K013505620HM0J143"]],"c":"5600µF","cv":5600.0,"t":"350V","tv":350.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014501520HM0H105"]],"c":"1500µF","cv":1500.0,"t":"450V","tv":450.0,"d":"63*105","dd":63.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K013502220HM0H105"]],"c":"2200µF","cv":2200.0,"t":"350V","tv":350.0,"d":"63*105","dd":63.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","ALS30A222MF400"]],"c":"2200µF","cv":2200.0,"t":"400V","tv":400.0,"d":"66*105","dd":66.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR222M450CC1"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"63*105","dd":63.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K013502220HM1M115"]],"c":"2200µF","cv":2200.0,"t":"350V","tv":350.0,"d":"72,5*115","dd":72.5,"dl":115.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K015002220HM0J143"]],"c":"2200µF","cv":2200.0,"t":"500V","tv":500.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014503320HM0J105"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*105","dd":76.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014503320HM0J093"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*93","dd":76.0,"dl":93.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR332M450DF1"],["KENDEIL","K014503320HM0J143"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K025003920HM0J143"]],"c":"3900µF","cv":3900.0,"t":"500V","tv":500.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014503320HM0J143"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K012504720HM0H105"]],"c":"4700µF","cv":4700.0,"t":"250V","tv":250.0,"d":"63*105","dd":63.0,"dl":105.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014504720HM0J143"]],"c":"4700µF","cv":4700.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K012504720HM0J093"]],"c":"4700µF","cv":4700.0,"t":"250V","tv":250.0,"d":"76*93","dd":76.0,"dl":93.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014505620HM0J143"]],"c":"5600µF","cv":5600.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR682M450DF1"],["KENDEIL","K014506820HM0J143"]],"c":"6800µF","cv":6800.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYXHR682X450EF1"],["KENDEIL","K014506820HM0L145"]],"c":"6800µF","cv":6800.0,"t":"450V","tv":450.0,"d":"90*149","dd":90.0,"dl":149.0,"cx":"Ix","fx":"Fond plat"},{"r":[["KENDEIL","K014506820HM0J220"]],"c":"6800µF","cv":6800.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYUXHR822M450DJ1"]],"c":"8200µF","cv":8200.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Fond plat"},{"r":[["ITELCOND","AYUXHR103M350DJ1"]],"c":"10000µF","cv":10000.0,"t":"350V","tv":350.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR103M250DF1"],["KENDEIL","K01250103SHM0J143"]],"c":"10000µF","cv":10000.0,"t":"250V","tv":250.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01400103SHM0J220"]],"c":"10000µF","cv":10000.0,"t":"400V","tv":400.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450103SHM0J143"]],"c":"10000µF","cv":10000.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450103SHM0J220"]],"c":"10000µF","cv":10000.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR102X450BC1"]],"c":"1000µF","cv":1000.0,"t":"450V","tv":450.0,"d":"51*105","dd":51.0,"dl":105.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","ARUX332X350DF1"],["KENDEIL","K01350332SHM0J143"]],"c":"3300µF","cv":3300.0,"t":"350V","tv":350.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450222SHM0J093"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"76*93","dd":76.0,"dl":93.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450222SHM0J143"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR222M450DC1"],["KENDEIL","K01450222SHM0J105"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"76*105","dd":76.0,"dl":105.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01420442SHM0J143"]],"c":"4400µF","cv":4400.0,"t":"420V","tv":420.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR332M350DC1"],["KENDEIL","K01350332SHM0J105"]],"c":"3300µF","cv":3300.0,"t":"350V","tv":350.0,"d":"76*105","dd":76.0,"dl":105.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450502SHM0J220"]],"c":"5000µF","cv":5000.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR682M350DF1"],["KENDEIL","K01350682SHM0J143"]],"c":"6800µF","cv":6800.0,"t":"350V","tv":350.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01350562SHM0J143"]],"c":"5600µF","cv":5600.0,"t":"350V","tv":350.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01350222SHM0H105"]],"c":"2200µF","cv":2200.0,"t":"350V","tv":350.0,"d":"63*105","dd":63.0,"dl":105.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","APUX222M500DF1"],["KENDEIL","K01500222SHM0J143"]],"c":"2200µF","cv":2200.0,"t":"500V","tv":500.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450332SHM0J105"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*105","dd":76.0,"dl":105.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450332SHM0J093"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*93","dd":76.0,"dl":93.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450332SHM0J143"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["NIPPON CHEMICON NCC","U32F11e14L05 code E32F501CPS392MEE3M"]],"c":"3900µF","cv":3900.0,"t":"500V","tv":500.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUX392M500DF1"],["KENDEIL","K01500392SHM0J143"]],"c":"3900µF","cv":3900.0,"t":"500V","tv":500.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","APUX392M500DF1"],["KENDEIL","K02500392SHM0J143"]],"c":"3900µF","cv":3900.0,"t":"500V","tv":500.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx","n":"T=105°c"},{"r":[["KENDEIL","K01450332SHM0J143"]],"c":"3300µF","cv":3300.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR472M350DJ1"],["KENDEIL","K01350472SHM0J143"]],"c":"4700µF","cv":4700.0,"t":"350V","tv":350.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450472SHM0J143"]],"c":"4700µF","cv":4700.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01500472SHM0L145"]],"c":"4700µF","cv":4700.0,"t":"500V","tv":500.0,"d":"90*143","dd":90.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450562SHM0J220"]],"c":"5600µF","cv":5600.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450562SHM0J143"]],"c":"5600µF","cv":5600.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01500562SHQ0J220"]],"c":"5600µF","cv":5600.0,"t":"500V","tv":500.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01500562SHQ0J214"]],"c":"5600µF","cv":5600.0,"t":"500V","tv":500.0,"d":"76*214","dd":76.0,"dl":214.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450602SHM0J220"]],"c":"6000µF","cv":6000.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450602SHM0J143"]],"c":"6000µF","cv":6000.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Mx"},{"r":[["NIPPON CHEMICON NCC","E36F501CSS602QEM9M"],["KENDEIL","EPCOS B43584S6608Q1"]],"c":"6000µF","cv":6000.0,"t":"500V","tv":500.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx","n":"T=85°c"},{"r":[["ITELCOND","AYUXHR602M500DJ1"],["KENDEIL","K01500602SHM0J220"]],"c":"6000µF","cv":6000.0,"t":"500V","tv":500.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450682SHM0L145"]],"c":"6800µF","cv":6800.0,"t":"450V","tv":450.0,"d":"90*145","dd":90.0,"dl":145.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AYUXHR682M450DF1"]],"c":"6800µF","cv":6800.0,"t":"450V","tv":450.0,"d":"76*147","dd":76.0,"dl":147.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01450682SHM0J220"]],"c":"6800µF","cv":6800.0,"t":"450V","tv":450.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["KENDEIL","K01500682SHMOJ220"]],"c":"6800µF","cv":6800.0,"t":"500V","tv":500.0,"d":"76*220","dd":76.0,"dl":220.0,"cx":"Ix","fx":"Mx"},{"r":[["ITELCOND","AKS102M450PG1"],["KENDEIL","K054501025PM0F097"]],"c":"1000µF","cv":1000.0,"t":"450V","tv":450.0,"d":"40*97","dd":40.0,"dl":97.0,"cx":"Ix"},{"r":[["ITELCOND","ACS152M450PG1"],["KENDEIL","K054501524PM0F097"]],"c":"1500µF","cv":1500.0,"t":"450V","tv":450.0,"d":"40*97","dd":40.0,"dl":97.0,"cx":"Ix"},{"r":[["KENDEIL","K054501525PM0F097"]],"c":"1500µF","cv":1500.0,"t":"450V","tv":450.0,"d":"40*97","dd":40.0,"dl":97.0,"cx":"Ix"},{"r":[["ITELCOND","ACS222M450QG1"],["KENDEIL","K164502224PM0N105"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"45*105","dd":45.0,"dl":105.0,"cx":"Ix"},{"r":[["ITELCOND","ACC471M450NC1"],["KENDEIL","K054504710PM0E050"]],"c":"470µF","cv":470.0,"t":"450V","tv":450.0,"d":"35*50","dd":35.0,"dl":50.0,"cx":"Ix"},{"r":[["ITELCOND","AKS471M450PE1"]],"c":"470µF","cv":470.0,"t":"450V","tv":450.0,"d":"40*75","dd":40.0,"dl":75.0,"cx":"Ix"},{"r":[["ITELCOND","AYUXHR222M450DF1"],["KENDEIL","K014502220HM0J143"]],"c":"2200µF","cv":2200.0,"t":"450V","tv":450.0,"d":"76*143","dd":76.0,"dl":143.0,"cx":"Ix","fx":"Fond plat"}],"ventil":[{"r":[["EBM PAPST","9956"]],"t":"230V","tv":230.0,"d":"119*119*25"},{"r":[["SUNON","DP200A"],["EBM PAPST","4650N"],["NMB","4715MS23TB5AD00"]],"t":"230V","tv":230.0,"d":"119*119*38"},{"r":[["EBM PAPST","W2K121AA0103"]],"t":"230V","tv":230.0,"d":"127*127*38"},{"r":[["EBM PAPST","5950"]],"t":"230V","tv":230.0,"d":"127*127*38"},{"r":[["EBM PAPST","W1G180AB3101"]],"t":"230V","tv":230.0},{"r":[["EBM PAPST","W2G130AA3301"]],"t":"230V","tv":230.0},{"r":[["SAVIOR","SA17251H2BX"],["EBM PAPST","W2S130AA0301"]],"t":"230V","tv":230.0,"n":"EBM PAPST - W2S130AA0301 (Anciennement Papst 7855 ES)"},{"r":[["EBM PAPST","5656S"]],"t":"230V","tv":230.0,"d":"135*135*38"},{"r":[["SUNON","A2175HBTTCGN"]],"t":"230V","tv":230.0,"d":"171*151*51","n":"(marque HBL)"},{"r":[["EBM PAPST","W2E142BB0101"],["NMB","5915PC23TB30A00"]],"t":"230V","tv":230.0,"d":"172*150*38"},{"r":[["EBM PAPST","W2E143AB0901"]],"t":"230V","tv":230.0,"d":"DIAM 143 172*172*51"},{"r":[["EBM PAPST","W2E143AA0901"]],"t":"230V","tv":230.0,"d":"DIAM 143 172*172*52"},{"r":[["SAVIOR","SA16062H2B"]],"t":"220V","tv":220.0,"d":"160*160*62"},{"r":[["ECOFIT","2VGCu25200V"],["EBM PAPST","W2E200HK3801"]],"t":"230V","tv":230.0,"d":"DIAM 200 225*225*80","n":"ref ecofit  (C49 A9p)"},{"r":[["EBM PAPST","W2E200HK3813"]],"t":"230V","tv":230.0,"d":"DIAM 200 225*225*80"},{"r":[["ORION","OA180APL222"]],"t":"230V","tv":230.0,"d":"180*180*65"},{"r":[["EBM PAPST","W2E250HL0601"]],"t":"230V","tv":230.0},{"r":[["EBM PAPST","W2E300CP0237"]],"t":"230V","tv":230.0,"d":"DIAM 326 397*80"},{"r":[["EBM PAPST","3950"],["EBM PAPST","3956"]],"t":"230V","tv":230.0,"d":"92*92*25"},{"r":[["EBM PAPST","4412F"]],"t":"12V","tv":12.0,"d":"119*119*25"},{"r":[["EBM PAPST","4414F"]],"t":"24V","tv":24.0,"d":"119*119*25"},{"r":[["EBM PAPST","4314M"],["NMB","4712KL05WB20"]],"t":"24V","tv":24.0,"d":"119*119*32"},{"r":[["EBM PAPST","4114NH3"],["NMB","4715SL05WB60"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["NIDEC","G1238B24BAZZ32"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["SAN ACE 120","9G1224G102"],["EBM PAPST","4414HR"],["NMB","4715KL05WB50"]],"t":"24V","tv":24.0,"d":"119*119*38","q":"137CFM","n":"sans capteur - SAN ACE 120 => Sanyo Denki"},{"r":[["EBM PAPST","4414M"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["DC BRUSHLESS","AFB1212VHEF00"]],"t":"12V","tv":12.0,"d":"120*120*38"},{"r":[["SUNON","EEC0251B2000UA99"]],"t":"12V","tv":12.0,"d":"120*120*25","q":"93CFM"},{"r":[["NMB","4715KL04WB40P00"]],"t":"12V","tv":12.0,"d":"119*119*38"},{"r":[["SUNON","KD1212PMB16A"]],"t":"12V","tv":12.0,"d":"120*120*25","q":"108CFM"},{"r":[["SAN ACE 120","9G1212E1D011"]],"t":"12V","tv":12.0,"d":"120*120*38","n":"With lock sensor - SAN ACE 120 => Sanyo Denki"},{"r":[["SAN ACE 120","9G1212G101"]],"t":"12V","tv":12.0,"d":"120*120*38","n":"With pulse sensor - SAN ACE 120 => Sanyo Denki"},{"r":[["SAN ACE 120","9G1212G1D01"]],"t":"12V","tv":12.0,"d":"120*120*38","n":"With lock sensor - SAN ACE 120 => Sanyo Denki"},{"r":[["NMB","4715KL05WB40P00"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["EBM PAPST","4184N17X"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["EBM PAPST","4414H"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["EBM PAPST","4414HH"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["NIDEC","D1238B24B9ZZ61"],["EBM PAPST","4114NH4"]],"t":"24V","tv":24.0,"d":"119*119*38"},{"r":[["DELTA","AFB1224SHE"]],"t":"24V","tv":24.0,"d":"120*120*38"},{"r":[["DELTA","AFB1224HHE"]],"t":"24V","tv":24.0,"d":"120*120*38"},{"r":[["ARX","FD2413A1251G"]],"t":"24V","tv":24.0,"d":"127*127*38","n":"PP (Polarity protection) + ARR  (Auto Restart) +3rd wire (Frequency Generator output)"},{"r":[["EBM PAPST","DV52142N"]],"t":"24V","tv":24.0,"d":"127*127*38"},{"r":[["EBM PAPST","W2E130BM1501"]],"t":"110V","tv":110.0,"d":"150*55*172"},{"r":[["DELTA","FFB1248EHER00"]],"t":"48V","tv":48.0,"d":"120*120*38"},{"r":[["DELTA","FFB1248EHEF00"]],"t":"48V","tv":48.0,"d":"120*120*38","n":"frequency generator signal"},{"r":[["SAN ACE 120L","9LG1248P1S001"]],"t":"48V","tv":48.0,"d":"120*120*38","n":"With pulse sensor with PWM control function  - SAN ACE 120L => Sanyo Denki\nj'ai inverser 2 lettres pour trouver une corespondence 9GL1248P1S001"},{"r":[["NIDEC","G1238B48BBZP00"]],"t":"48V","tv":48.0,"d":"120x120x38"},{"r":[["NIDEC","G1751M48B8ZZ00"]],"t":"48V","tv":48.0},{"r":[["EBM PAPST","DV6224"]],"t":"24V","tv":24.0,"d":"172x160x51"},{"r":[["NMB","3110KL05WB50"]],"t":"24V","tv":24.0,"d":"172*150*25"},{"r":[["EBM PAPST","614NN"],["NMB","2410ML05WB50"]],"t":"24V","tv":24.0,"d":"60*60*25"},{"r":[["NMB","OD602512LB"]],"t":"12V","tv":12.0,"d":"60*60*25"},{"r":[["NMB","3110KL05WB50"]],"t":"24V","tv":24.0,"d":"80*80*25"},{"r":[["EBM PAPST","8412NGMLE"]],"t":"12V","tv":12.0,"d":"80*80*25"},{"r":[["SAN COOLER 80","9A0812S4011"],["DELTA","AFB0812HHF00"]],"t":"12V","tv":12.0,"d":"80*80*25","n":"With pulse sensor - SAN COOLER 80  => Sanyo Denki"},{"r":[["SAN COOLER 80","9A0812G4D011"],["SUNON","EE80251B2000UG99"]],"t":"12V","tv":12.0,"d":"80*80*25","q":"37CFM","n":"With lock sensor - SAN COOLER 80 => Sanyo Denki\na eviter SUNON - EE80251B2000UG99 (Le G c'est lock sensor auto resart)"},{"r":[["NMB","3110SB05WB50E00"]],"t":"24V","tv":24.0,"d":"80*80*25"},{"r":[["DELTA","AFB0812SHR00"]],"t":"12V","tv":12.0,"d":"80*80*25","n":"With lock sensor"},{"r":[["DELTA","AFB0824M"]],"t":"24V","tv":24.0,"d":"80*80*25","n":"With lock sensor"},{"r":[["ADDA","AD0924MBA70GL"],["EBM PAPST","3414NL"]],"t":"24V","tv":24.0,"d":"92*92*25"},{"r":[["EBM PAPST","3412NGME"]],"t":"12V","tv":12.0,"d":"92*92*25"},{"r":[["YaLn","D90BH12"],["NMB","3610KL04WB50"]],"t":"12V","tv":12.0,"d":"92*92*25"},{"r":[["ARX","FD1290A1151E"]],"t":"12V","tv":12.0,"d":"92*92*25","n":"frequency generator signal"},{"r":[["NMB","3610KL04WB40"]],"t":"12V","tv":12.0,"d":"92*92*25"},{"r":[["EBM PAPST","3610KL05WB50"]],"t":"24V","tv":24.0,"d":"92*92*25"},{"r":[["DELTA","AFB0924HH"]],"t":"24V","tv":24.0,"d":"92*92*25"},{"r":[["MARQUE","REF_1"]],"t":"12V","tv":12.0,"d":"92x92x38"},{"r":[["SUNON","PMD1209PMB3A"],["NMB","3615RL04WB39"],["NMB","09238RA12LFL"]],"t":"12V","tv":12.0,"d":"92x92x38"},{"r":[["SAN COOLER 92","9A0912S4D04"]],"t":"12V","tv":12.0,"d":"92*92*25","n":"With  ? sensor - SAN COOLER 92 => Sanyo Denki"},{"r":[["DELTA","AFB0912HH"]],"t":"12V","tv":12.0,"d":"92*92*25"},{"r":[["EBM PAPST","R2E220AA4005"]],"t":"230V","tv":230.0,"d":"220*220*44"},{"r":[["ECOFIT","2RRE45250x56R"],["EBM PAPST","R2E250AC0925"],["RUCK","RE2502B100507 (arrêté)"]],"t":"230V","tv":230.0,"d":"DIAM 250 252*252*56","n":"ref ecofit (H06 17)"},{"r":[["ZIEHL","RH28M 2DK3F2R"]],"t":"230V","tv":230.0,"d":"DIAM 284*90"},{"r":[["EBM PAPST","R4E280AA0105"]],"t":"230V","tv":230.0,"d":"DIAM 280*102"},{"r":[["EBM PAPST","A2E300AC4701"]],"t":"230V","tv":230.0},{"r":[["EBM PAPST","R2S133AB0313"],["EBM PAPST","R2E133RA0301"],["EBM PAPST","R2E133BH6614"]],"t":"230V","tv":230.0,"n":"EBM PAPST R2E133RA0301 + adaptation mécanique\n(R2E133BH66 14 pas achetable et Remplace R2S133 AB03 13)"},{"r":[["EBM PAPST","R4E225AI0110"]],"t":"230V","tv":230.0},{"r":[["EBM PAPST","R2E250AS4726"]],"t":"230V","tv":230.0,"d":"DIAM 250 252*252*56"},{"r":[["EBM PAPST","R2E225RA9209"]],"t":"230V","tv":230.0,"d":"225*99"},{"r":[["EBM PAPST","R2E250RA5001 + KIT ALU"]],"t":"230V","tv":230.0,"d":"DIAM 250 252*252*56","n":"EBM PAPST - R2E250RA5001 + KIT ALU"},{"r":[["EBM PAPST","R2E250RA5001"]],"t":"230V","tv":230.0,"d":"DIAM 250 252*252*56"},{"r":[["EBM PAPST","R2E220AA4005"],["RUCK","RE2202F120768"]],"t":"230V","tv":230.0,"d":"DIAM 220"},{"r":[["EBM PAPST","R3G190RC0503"]],"t":"230V","tv":230.0},{"r":[["ECOFIT","2RRE25220x45R"]],"t":"230V","tv":230.0,"d":"220*45R","n":"ref ecofit    (C36 A4p)"},{"r":[["EBM PAPST","VC133D KIT05 + D2E133 DM64 70"]],"t":"230V","tv":230.0,"d":"DIAM 133"},{"r":[["EBM PAPST","D2E133DM64I5"]],"t":"230V","tv":230.0,"d":"DIAM 133","n":"EBM PAPST - VC133D KIT05 + D2E133 DM64 70"},{"r":[["ECOFIT","2GDR45Q0730146180L"],["EBM PAPST","D2E146AP4702"],["RUCK","SP108071"]],"t":"230V","tv":230.0,"d":"DIAM 146","n":"RUCK - GD133 2J 108071\n( remplacé par  SP 108071)"},{"r":[["ECOFIT","2GDS25133190R"],["EBM PAPST","D2E133AM4701"],["RUCK","GD1332A100184"]],"t":"230V","tv":230.0,"d":"DIAM 133"},{"r":[["ECOFIT","2GDS35133190LL23A8"]],"t":"230V","tv":230.0,"d":"133*190"},{"r":[["ECOFIT","2GDS35133190LL23A6"]],"t":"230V","tv":230.0,"d":"133*190"},{"r":[["EBM PAPST","R2E190RA2605"]],"t":"230V","tv":230.0,"d":"DIAM 190"},{"r":[["DELTA","AFB1212HHER00"]],"t":"12V","tv":12.0,"d":"120*120*38","q":"120.07","n":"detection d'arret rotation"},{"r":[["DELTA","AFB1212SHER00"]],"t":"12V","tv":12.0,"d":"120*120*38"},{"r":[["EBM PAPST","44142NA"]],"t":"24V","tv":24.0,"d":"119*119*37"},{"r":[["NMB","4715VL05WB46"]],"t":"24V","tv":24.0,"d":"119*119*38"}],"cartes":[{"m":"SOCOMEC","r":"S07058101","v":"S07058x01","c":[{"q":12,"k":"dc","i":72}]},{"m":"SOCOMEC","r":"S07030101","v":"S07030x01","c":[{"q":6,"k":"dc","i":72}]},{"m":"SOCOMEC","r":"S07030102","v":"S07030x02","c":[{"q":12,"k":"dc","i":72}]},{"m":"SOCOMEC","r":"S06949201","v":"S06949x01","c":[{"q":6,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06949202","v":"S06949x02","c":[{"q":12,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06949203","v":"S06949x03","c":[{"q":8,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06966201","v":"S06966x01","c":[{"q":4,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06966202","v":"S06966x02","c":[{"q":6,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06967101","v":"S06967x01","c":[{"q":6,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06968101","v":"S06968x01","c":[{"q":4,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06972101","v":"S06972x01","c":[{"q":12,"k":"dc","i":71}]},{"m":"SOCOMEC","r":"S06968201","v":"S06968x01","c":[]},{"m":"SOCOMEC","r":"S06830201","v":"S06830x01","c":[{"q":6,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06830202","v":"S06830x02","c":[{"q":4,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06830205","v":"S06830x05","c":[{"q":6,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06852101","v":"S06852x01","c":[{"q":12,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06852102","v":"S06852x02","c":[{"q":10,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06852103","v":"S06852x03","c":[{"q":8,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06826301","v":"S06826x01","c":[{"q":12,"k":"dc","i":73}]},{"m":"SOCOMEC","r":"S06826302","v":"S06826x02","c":[{"q":8,"k":"dc","i":73}]},{"m":"SOCOMEC","r":"S06828201","v":"S06828x01","c":[{"q":2,"k":"dc","i":73},{"q":2,"k":"dc","i":69}]},{"m":"SOCOMEC","r":"S06828202","v":"S06828x02","c":[{"q":4,"k":"dc","i":73}]},{"m":"SOCOMEC","r":"S06830203","v":"S06830x03","c":[{"q":6,"k":"dc","i":73}]},{"m":"SOCOMEC","r":"S06830204","v":"S06830x04","c":[{"q":4,"k":"dc","i":73}]},{"m":"MGE","r":"CARP51030161","v":"","c":[{"q":8,"k":"dc","i":72}]},{"m":"MGE","r":"CARP51030159","v":"","c":[{"q":14,"k":"dc","i":72}]}]};
const fmtNum = (v) => (v === null || v === undefined ? "" : String(Math.round(v * 100) / 100).replace(".", ","));
function libelleCondensateur(e, kind) {
  const [m, r] = e.r[0];
  const parts = [e.c, e.t, e.d && (kind === "ventil" ? e.d : "Ø" + String(e.d).replace("*", "×")), e.cx, e.fx].filter(Boolean);
  return `${kind === "ac" ? "Condensateur AC" : kind === "dc" ? "Condensateur DC" : "Ventilateur"} ${parts.join(" · ")}${m ? ` (${m} ${r})` : ` (${r})`}`;
}
function compositionCarte(carte) {
  return carte.c.map((x) => {
    const e = CATALOGUE_PIECES[x.k][x.i];
    return e ? { q: x.q, e } : null;
  }).filter(Boolean);
}
function CataloguePiecesPicker({ kind, onChoisir, onFermer }) {
  const [onglet, setOnglet] = useState(kind === "tous" ? "ac" : kind);
  const [f, setF] = useState({ texte: "", cap: "", tension: "", diam: "", marque: "", dim: "" });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const liste = onglet === "cartes" ? CATALOGUE_PIECES.cartes : CATALOGUE_PIECES[onglet] || [];
  const marques = useMemo(() => [...new Set(liste.flatMap((e) => (onglet === "cartes" ? [e.m] : e.r.map((x) => x[0]))).filter(Boolean))].sort(), [onglet]);
  const norm = (s) => String(s || "").toLowerCase().replace(/[\s\-_.]/g, "");
  const capF = catNum(f.cap), tensionF = catNum(f.tension), diamF = catNum(f.diam);
  const dimF = String(f.dim || "").replace(/[x×X]/g, "*").replace(/\s/g, "");
  const resultats = useMemo(() => {
    const t = norm(f.texte);
    const out = [];
    liste.forEach((e, i) => {
      if (onglet === "cartes") {
        const comp = compositionCarte(e);
        const texte = norm([e.m, e.r, e.v, ...comp.flatMap((x) => x.e.r.flat())].join(" "));
        if (t && !texte.includes(t)) return;
        if (f.marque && e.m !== f.marque) return;
        if (capF !== null && !comp.some((x) => x.e.cv && Math.abs(x.e.cv - capF) / capF < 0.01)) return;
        out.push({ e, i, comp, score: 0 });
        return;
      }
      const texte = norm(e.r.flat().join(" ") + " " + (e.n || ""));
      if (t && !texte.includes(t)) return;
      if (f.marque && !e.r.some((x) => x[0] === f.marque)) return;
      let score = 0, compatible = true;
      const ecarts = [];
      if (onglet !== "ventil") {
        if (capF !== null) { if (!e.cv || Math.abs(e.cv - capF) / capF > 0.01) return; }
        if (tensionF !== null) { if (!e.tv || e.tv < tensionF) return; score += (e.tv - tensionF) / tensionF; if (e.tv > tensionF) ecarts.push("tension supérieure"); }
        if (diamF !== null) { if (!e.dd || Math.abs(e.dd - diamF) > 0.5) return; }
        if (capF === null || tensionF === null) compatible = false;
      } else {
        if (tensionF !== null && (!e.tv || Math.abs(e.tv - tensionF) > 0.5)) return;
        if (dimF && !String(e.d || "").replace(/\s/g, "").includes(dimF)) return;
        if (tensionF === null || !dimF) compatible = false;
      }
      out.push({ e, i, score, compatible, ecarts });
    });
    return out.sort((a, b) => a.score - b.score).slice(0, 60);
  }, [onglet, f.texte, f.marque, capF, tensionF, diamF, dimF]);
  const th = { background: "#1A1F26", color: "#fff", padding: "6px 8px", fontSize: 11, textAlign: "left", fontWeight: 700, whiteSpace: "nowrap" };
  const td = { padding: "6px 8px", borderBottom: "1px solid #E2E6EB", fontSize: 12, verticalAlign: "top" };
  const champ = (label, k, w, ph) => (
    <label style={{ display: "flex", flexDirection: "column", gap: 3, fontSize: 10.5, color: "#5B6B7D", fontWeight: 600 }}>
      {label}
      <input value={f[k]} onChange={(ev) => set(k, ev.target.value)} placeholder={ph} style={{ ...catInput, width: w, padding: "6px 8px", fontSize: 12.5 }} />
    </label>
  );
  const choisirCarte = (carte, comp) => onChoisir({
    marque: carte.m, modele: carte.r,
    description: `Carte condensateurs : ${comp.map((x) => `${x.q} × ${x.e.c} ${x.e.t} Ø${String(x.e.d || "").replace("*", "×")} (${x.e.r[0].filter(Boolean).join(" ")})`).join(" + ") || "composition non renseignée"}`,
  });
  return (
    <div style={{ border: `1.5px solid ${CAT_BLUE}`, borderRadius: 10, padding: 12, marginTop: 8, background: "#F8FAFC" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {(kind === "tous" ? [["ac", "Condensateurs AC"], ["dc", "Condensateurs DC"], ["cartes", "Cartes condensateurs"], ["ventil", "Ventilateurs"]] : kind === "dc" ? [["dc", "Condensateurs DC"], ["cartes", "Cartes condensateurs"]] : kind === "ac" ? [["ac", "Condensateurs AC"]] : [["ventil", "Ventilateurs"]]).map(([k, l]) => (
            <button key={k} onClick={() => setOnglet(k)} style={{ ...catBtn(CAT_BLUE), background: onglet === k ? CAT_BLUE : "transparent", color: onglet === k ? "#fff" : CAT_BLUE }}>{l}</button>
          ))}
        </div>
        <button onClick={onFermer} style={{ background: "none", border: "none", color: "#5B6B7D", cursor: "pointer", fontSize: 12 }}>Fermer ✕</button>
      </div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 10 }}>
        {champ("Recherche (référence, marque…)", "texte", 200, "ex. K0145, 9956…")}
        {onglet !== "ventil" && champ("Capacité (µF)", "cap", 100, "ex. 4700")}
        {(onglet === "ac" || onglet === "dc") && champ("Tension mini (V)", "tension", 100, "ex. 450")}
        {(onglet === "ac" || onglet === "dc") && champ("Diamètre (mm)", "diam", 90, "ex. 76")}
        {onglet === "ventil" && champ("Tension (V)", "tension", 90, "12, 24, 48, 230")}
        {onglet === "ventil" && champ("Dimensions", "dim", 130, "ex. 119*119*38")}
        <label style={{ display: "flex", flexDirection: "column", gap: 3, fontSize: 10.5, color: "#5B6B7D", fontWeight: 600 }}>
          Marque
          <select value={f.marque} onChange={(ev) => set("marque", ev.target.value)} style={{ ...catInput, padding: "6px 8px", fontSize: 12.5 }}>
            <option value="">Toutes</option>
            {marques.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
        </label>
        <span style={{ fontSize: 11.5, color: "#8B96A3" }}>{resultats.length}{resultats.length === 60 ? "+" : ""} résultat(s)</span>
      </div>
      <div style={{ overflowX: "auto", maxHeight: 360, overflowY: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 640 }}>
          <thead>
            {onglet === "cartes" ? (
              <tr><th style={th}>Carte</th><th style={th}>Condensateurs montés</th><th style={th}></th></tr>
            ) : onglet === "ventil" ? (
              <tr><th style={th}>Marque — Référence</th><th style={th}>Équivalences</th><th style={th}>Tension</th><th style={th}>Dimensions</th><th style={th}>Remarque</th><th style={th}></th></tr>
            ) : (
              <tr><th style={th}>Marque — Référence</th><th style={th}>Équivalence</th><th style={th}>Capacité</th><th style={th}>Tension</th><th style={th}>Ø × L (mm)</th><th style={th}>Connexion</th><th style={th}>Fixation</th><th style={th}></th></tr>
            )}
          </thead>
          <tbody>
            {resultats.map(({ e, i, comp, compatible, ecarts }) => onglet === "cartes" ? (
              <tr key={i}>
                <td style={td}><b>{e.m} {e.r}</b>{e.v && <div style={{ fontSize: 10.5, color: "#8B96A3" }}>variante {e.v}</div>}</td>
                <td style={td}>{comp.length ? comp.map((x, j) => (
                  <div key={j}><b>{x.q} ×</b> {x.e.c} · {x.e.t} · Ø{String(x.e.d || "").replace("*", "×")} {x.e.cx ? `· ${x.e.cx}` : ""} — <span style={{ color: "#5B6B7D" }}>{x.e.r.map((y) => y.filter(Boolean).join(" ")).join(" / ")}</span></div>
                )) : <span style={{ color: "#8B96A3" }}>composition non renseignée</span>}</td>
                <td style={td}><button onClick={() => choisirCarte(e, comp)} style={catBtn(CAT_BLUE)}>Choisir</button></td>
              </tr>
            ) : (
              <tr key={i} style={{ background: compatible ? "#F0FBF6" : undefined }}>
                <td style={td}><b>{e.r[0].filter(Boolean).join(" — ")}</b>{compatible && <div style={{ fontSize: 10.5, color: "#0F8A5F", fontWeight: 700 }}>✓ compatible{ecarts && ecarts.length ? ` (${ecarts.join(", ")})` : ""}</div>}</td>
                <td style={td}>{e.r.slice(1).map((y) => y.filter(Boolean).join(" ")).join(" / ") || "—"}</td>
                {onglet !== "ventil" && <td style={td}>{e.c || "—"}</td>}
                <td style={td}>{e.t || "—"}</td>
                <td style={td}>{e.d ? String(e.d).replace(/\*/g, "×") : "—"}</td>
                {onglet === "ventil" ? <td style={{ ...td, fontSize: 11, color: "#5B6B7D", maxWidth: 260 }}>{e.n || "—"}</td> : <>
                  <td style={td}>{e.cx || "—"}</td><td style={td}>{e.fx || "—"}{e.n && <div style={{ fontSize: 10.5, color: "#8B96A3" }}>{e.n}</div>}</td>
                </>}
                <td style={td}><button onClick={() => onChoisir({ marque: e.r[0][0], modele: e.r[0][1], description: libelleCondensateur(e, onglet) })} style={catBtn(CAT_BLUE)}>Choisir</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {onglet === "dc" && capF !== null && (() => {
        const cartes = CATALOGUE_PIECES.cartes.filter((c) => compositionCarte(c).some((x) => x.e.cv && Math.abs(x.e.cv - capF) / capF < 0.01));
        return cartes.length ? <div style={{ fontSize: 11.5, color: "#5B6B7D", marginTop: 8 }}>Ce condensateur de {fmtNum(capF)} µF est monté sur {cartes.length} carte(s) : {cartes.slice(0, 8).map((c) => `${c.m} ${c.r}`).join(", ")}{cartes.length > 8 ? "…" : ""} — voir l'onglet « Cartes condensateurs ».</div> : null;
      })()}
      <div style={{ fontSize: 10.5, color: "#8B96A3", marginTop: 8 }}>« Compatible » : même capacité, tension égale ou supérieure{onglet === "ventil" ? " — ventilateurs : même tension et mêmes dimensions" : ", même diamètre si renseigné ; vérifier la longueur, la connexion et la fixation avant commande"}. Connexion : I = insert, M = tige filetée, B = bornes.</div>
    </div>
  );
}

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
    base.batteries = { ...base.batteries, items: trierParLabel(base.batteries.items) };
    return base;
  };

  const [tarifs, setTarifs] = useState(tarifsValides(saved.tarifs) ? saved.tarifs : DEFAULT_TARIFS);
  const [majorations, setMajorations] = useState(saved.majorations || DEFAULT_MAJORATIONS);
  const [degressivite, setDegressivite] = useState(saved.degressivite || DEFAULT_DEGRESSIVITE);
  const [coefContrat, setCoefContrat] = useState(saved.coefContrat || DEFAULT_COEF_CONTRAT);

  const [catalogueTemps, setCatalogueTemps] = useState(catalogueTempsValide(saved.catalogueTemps) ? saved.catalogueTemps : DEFAULT_CATALOGUE_TEMPS);
  const [catalogueCoef, setCatalogueCoef] = useState({ ...DEFAULT_CATALOGUE_COEF, ...(saved.catalogueCoef || {}) });
  const [catalogueDirect, setCatalogueDirect] = useState(migrerCatalogueDirect(saved.catalogueDirect));

  // Extrait le premier nombre trouvé dans un libellé (ex. "45 Ah" -> 45,
  // "60°C" -> 60) pour trier les catalogues par ordre croissant. Renvoie
  // Infinity si aucun nombre n'est trouvé (l'élément reste en fin de liste).

  const ajouterLigneCatalogueDirect = (familleKey) =>
    setCatalogueDirect((s) => {
      const nouvelItem =
        familleKey === "batteries" ? { id: uid(), label: "Nouvelle capacité", prixAchat: 0, coef: 2, prix: 0 } : { id: uid(), label: "Nouvel article", prix: 0 };
      const items = trierParLabel([...s[familleKey].items, nouvelItem]);
      return { ...s, [familleKey]: { ...s[familleKey], items } };
    });
  const supprimerLigneCatalogueDirect = (familleKey, itemId) =>
    setCatalogueDirect((s) => ({ ...s, [familleKey]: { ...s[familleKey], items: s[familleKey].items.filter((it) => it.id !== itemId) } }));
  const renommerLigneCatalogueDirect = (familleKey, itemId, label) =>
    setCatalogueDirect((s) => {
      const items = trierParLabel(s[familleKey].items.map((it) => (it.id === itemId ? { ...it, label } : it)));
      return { ...s, [familleKey]: { ...s[familleKey], items } };
    });

  // FirePro — ajout/suppression/renommage des générateurs
  const ajouterGenerateurFirePro = () =>
    setFireproGenerateurs((arr) => [...arr, { id: uid(), code: "", label: "Nouveau générateur", masse: 0, prix: 0 }]);
  const supprimerGenerateurFirePro = (id) => setFireproGenerateurs((arr) => arr.filter((g) => g.id !== id));
  const renommerGenerateurFirePro = (id, label) => setFireproGenerateurs((arr) => arr.map((g) => (g.id === id ? { ...g, label } : g)));
  const majMasseGenerateurFirePro = (id, masse) =>
    setFireproGenerateurs((arr) => [...arr.map((g) => (g.id === id ? { ...g, masse } : g))].sort((a, b) => a.masse - b.masse));

  // FirePro — ajout/suppression/renommage des accessoires (par famille)
  const ajouterAccessoireFirePro = (famId) =>
    setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: trierParLabel([...s[famId].items, { id: uid(), code: "", label: "Nouvel article", prix: 0 }]) } }));
  const supprimerAccessoireFirePro = (famId, itemId) =>
    setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: s[famId].items.filter((it) => it.id !== itemId) } }));
  const renommerAccessoireFirePro = (famId, itemId, label) =>
    setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: trierParLabel(s[famId].items.map((it) => (it.id === itemId ? { ...it, label } : it))) } }));

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
        heuresJour: 7,
      },
      postesEquipement: [nouveauPosteEquipement(1)],
      lignesLibres: [],
    };
  }

  // Contenu vierge d'un devis FirePro — totalement indépendant du devis
  // HT/BT (référence, client, site, locaux propres).
  function devisFireProVierge() {
    return {
      fireproAffaire: {
        client: "",
        site: "",
        reference: "FIR-" + new Date().getFullYear() + "-001",
      },
      zonesFirePro: [nouvelleZoneFirePro(1)],
      fireproCoefAjustement: 1,
    };
  }

  const [affaire, setAffaire] = useState(saved.affaire || devisVierge().affaire);
  const [postesEquipement, setPostesEquipement] = useState(saved.postesEquipement || devisVierge().postesEquipement);
  const [lignesLibres, setLignesLibres] = useState(saved.lignesLibres || []);
  const [nbAAjouter, setNbAAjouter] = useState(1);

  // Paramètres FirePro — partagés entre TOUS les devis (HT/BT et FirePro), comme tarifs/catalogues.
  const [fireproClasses, setFireproClasses] = useState(saved.fireproClasses || DEFAULT_FIREPRO_CLASSES);
  const [fireproConfinement, setFireproConfinement] = useState(saved.fireproConfinement || DEFAULT_FIREPRO_CONFINEMENT);
  const [fireproGenerateurs, setFireproGenerateurs] = useState(saved.fireproGenerateurs || DEFAULT_FIREPRO_GENERATEURS);
  const [fireproAccessoires, setFireproAccessoires] = useState(saved.fireproAccessoires || DEFAULT_FIREPRO_ACCESSOIRES);

  // Devis FirePro courant — totalement indépendant du devis HT/BT (sa propre
  // liste, son propre sélecteur, sa propre référence FIR-...).
  const [fireproAffaire, setFireproAffaire] = useState(saved.fireproAffaire || devisFireProVierge().fireproAffaire);
  const [zonesFirePro, setZonesFirePro] = useState(saved.zonesFirePro || devisFireProVierge().zonesFirePro);
  const [fireproCoefAjustement, setFireproCoefAjustement] = useState(saved.fireproCoefAjustement ?? 1);
  const [fireproDevisList, setFireproDevisList] = useState(saved.fireproDevisList || []);
  const [currentFireproDevisId, setCurrentFireproDevisId] = useState(saved.currentFireproDevisId || null);
  const [fireproDevisReady, setFireproDevisReady] = useState(false);

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

  // Contacts partagés entre toutes les apps HT Maintenance.
  const [contacts, setContacts] = useState(saved.contacts || []);
  const [selectedContactId, setSelectedContactId] = useState("");
  const [editingContact, setEditingContact] = useState(null); // brouillon en cours d'ajout/édition, ou null

  useEffect(() => {
    if (!authReady) return;
    loadContactsShared()
      .then(setContacts)
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [authReady]);

  const choisirContactPourAffaire = (id) => {
    setSelectedContactId(id);
    const c = contacts.find((x) => x.id === id);
    if (c) {
      setAffaire((a) => ({
        ...a,
        client: c.societe || c.contact || a.client,
        site: [c.adresse, c.cp, c.ville].filter(Boolean).join(", ") || a.site,
      }));
    }
  };
  const ouvrirNouveauContact = () => setEditingContact(nouveauContactVierge());
  const ouvrirEditionContact = (id) => {
    const c = contacts.find((x) => x.id === id);
    if (c) setEditingContact({ ...c, telephones: (c.telephones || []).map((t) => ({ ...t })) });
  };
  const annulerEditionContact = () => setEditingContact(null);
  const enregistrerContact = async () => {
    if (!editingContact) return;
    const existe = contacts.some((c) => c.id === editingContact.id);
    const prevContacts = contacts;
    const nextContacts = existe ? contacts.map((c) => (c.id === editingContact.id ? editingContact : c)) : [...contacts, editingContact];
    setContacts(nextContacts);
    setSelectedContactId(editingContact.id);
    setEditingContact(null);
    try {
      await syncContactsToShared(nextContacts, prevContacts);
    } catch {
      // la base partagée reste à jour au prochain enregistrement réussi
    }
  };
  const supprimerContact = async (id) => {
    if (!window.confirm("Supprimer ce client de la base partagée (utilisée par toutes vos apps HT Maintenance) ?")) return;
    const prevContacts = contacts;
    const nextContacts = contacts.filter((c) => c.id !== id);
    setContacts(nextContacts);
    if (selectedContactId === id) setSelectedContactId("");
    try {
      await syncContactsToShared(nextContacts, prevContacts);
    } catch {
      // la base partagée reste à jour au prochain enregistrement réussi
    }
  };
  const ajouterTelephoneDraft = () => setEditingContact((c) => ({ ...c, telephones: [...(c.telephones || []), { type: "Fixe", numero: "" }] }));
  const updateTelephoneDraft = (idx, patch) =>
    setEditingContact((c) => ({ ...c, telephones: c.telephones.map((t, i) => (i === idx ? { ...t, ...patch } : t)) }));
  const supprimerTelephoneDraft = (idx) => setEditingContact((c) => ({ ...c, telephones: c.telephones.filter((_, i) => i !== idx) }));

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

  // ---- Devis FirePro (totalement indépendant du devis HT/BT ci-dessus) ----
  useEffect(() => {
    if (!authReady) return;
    getDocs(collection(db, FIREPRO_DEVIS_COLLECTION))
      .then(async (snap) => {
        let list = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        list.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
        if (list.length === 0) {
          const blank = devisFireProVierge();
          const payload = { ...blank, updatedAt: Date.now() };
          const ref = await addDoc(collection(db, FIREPRO_DEVIS_COLLECTION), payload);
          list = [{ id: ref.id, ...payload }];
        }
        setFireproDevisList(list.map((d) => ({ id: d.id, reference: d.fireproAffaire?.reference, client: d.fireproAffaire?.client, updatedAt: d.updatedAt })));
        const target = list.find((d) => d.id === currentFireproDevisId) || list[0];
        setCurrentFireproDevisId(target.id);
        setFireproAffaire(target.fireproAffaire || devisFireProVierge().fireproAffaire);
        setZonesFirePro(target.zonesFirePro || devisFireProVierge().zonesFirePro);
        setFireproCoefAjustement(target.fireproCoefAjustement ?? 1);
        setFireproDevisReady(true);
      })
      .catch(() => setSyncState("error"));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [authReady]);

  const chargerDevisFirePro = (id) => {
    setFireproDevisReady(false);
    getDoc(doc(db, FIREPRO_DEVIS_COLLECTION, id))
      .then((snap) => {
        if (snap.exists()) {
          const data = snap.data();
          setCurrentFireproDevisId(id);
          setFireproAffaire(data.fireproAffaire || devisFireProVierge().fireproAffaire);
          setZonesFirePro(data.zonesFirePro || devisFireProVierge().zonesFirePro);
          setFireproCoefAjustement(data.fireproCoefAjustement ?? 1);
        }
        setFireproDevisReady(true);
      })
      .catch(() => setFireproDevisReady(true));
  };

  const nouveauChiffrageFirePro = async () => {
    const blank = devisFireProVierge();
    setFireproDevisReady(false);
    try {
      const payload = { ...blank, updatedAt: Date.now() };
      const ref = await addDoc(collection(db, FIREPRO_DEVIS_COLLECTION), payload);
      setFireproDevisList((l) => [{ id: ref.id, reference: blank.fireproAffaire.reference, client: "", updatedAt: payload.updatedAt }, ...l]);
      setCurrentFireproDevisId(ref.id);
      setFireproAffaire(blank.fireproAffaire);
      setZonesFirePro(blank.zonesFirePro);
      setFireproCoefAjustement(1);
    } finally {
      setFireproDevisReady(true);
    }
  };

  const dupliquerChiffrageFirePro = async () => {
    setFireproDevisReady(false);
    try {
      const payload = {
        fireproAffaire: { ...fireproAffaire, reference: fireproAffaire.reference + " (copie)" },
        zonesFirePro: zonesFirePro.map((z) => ({ ...z, generateurs: { ...z.generateurs }, accessoires: { ...z.accessoires } })),
        fireproCoefAjustement,
        updatedAt: Date.now(),
      };
      const ref = await addDoc(collection(db, FIREPRO_DEVIS_COLLECTION), payload);
      setFireproDevisList((l) => [{ id: ref.id, reference: payload.fireproAffaire.reference, client: payload.fireproAffaire.client, updatedAt: payload.updatedAt }, ...l]);
      setCurrentFireproDevisId(ref.id);
      setFireproAffaire(payload.fireproAffaire);
      setZonesFirePro(payload.zonesFirePro);
      setFireproCoefAjustement(payload.fireproCoefAjustement);
    } finally {
      setFireproDevisReady(true);
    }
  };

  const supprimerChiffrageFirePro = async () => {
    if (fireproDevisList.length <= 1) return;
    if (!window.confirm(`Supprimer définitivement le devis FirePro « ${fireproAffaire.reference} » ? Cette action est irréversible.`)) return;
    const idASupprimer = currentFireproDevisId;
    try {
      await deleteDoc(doc(db, FIREPRO_DEVIS_COLLECTION, idASupprimer));
    } catch {
      // on continue quand même côté interface
    }
    const reste = fireproDevisList.filter((d) => d.id !== idASupprimer);
    setFireproDevisList(reste);
    chargerDevisFirePro(reste[0].id);
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

  // Enregistre automatiquement le devis FirePro courant, indépendamment du devis HT/BT.
  useEffect(() => {
    if (!authReady || !fireproDevisReady || !currentFireproDevisId) return;
    setSyncState("syncing");
    const t = setTimeout(() => {
      const payload = { fireproAffaire, zonesFirePro, fireproCoefAjustement, updatedAt: Date.now() };
      setDoc(doc(db, FIREPRO_DEVIS_COLLECTION, currentFireproDevisId), payload)
        .then(() => {
          setSyncState("synced");
          setFireproDevisList((l) =>
            l.map((d) => (d.id === currentFireproDevisId ? { ...d, reference: fireproAffaire.reference, client: fireproAffaire.client, updatedAt: payload.updatedAt } : d))
          );
        })
        .catch(() => setSyncState("error"));
    }, 1000);
    return () => clearTimeout(t);
  }, [authReady, fireproDevisReady, currentFireproDevisId, fireproAffaire, zonesFirePro, fireproCoefAjustement]);

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
          catalogueTemps,
          catalogueCoef,
          catalogueDirect,
          affaire,
          postesEquipement,
          lignesLibres,
          devisList,
          currentDevisId,
          zonesFirePro,
          fireproAffaire,
          fireproDevisList,
          currentFireproDevisId,
          fireproClasses,
          fireproConfinement,
          fireproGenerateurs,
          fireproAccessoires,
          fireproCoefAjustement,
          contacts,
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
    catalogueTemps,
    catalogueCoef,
    catalogueDirect,
    affaire,
    postesEquipement,
    lignesLibres,
    devisList,
    currentDevisId,
    zonesFirePro,
    fireproAffaire,
    fireproDevisList,
    currentFireproDevisId,
    fireproClasses,
    fireproConfinement,
    fireproGenerateurs,
    fireproAccessoires,
    fireproCoefAjustement,
    contacts,
  ]);

  // Enregistre immédiatement les paramètres/catalogues, sans attendre le
  // délai automatique (utile pour une confirmation explicite à l'utilisateur).
  const enregistrerParametresMaintenant = () => {
    setSyncState("syncing");
    setDoc(doc(db, FIRESTORE_DOC), { tarifs, majorations, degressivite, coefContrat, catalogueTemps, catalogueCoef, catalogueDirect })
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

  // RAZ FirePro seul : vide les locaux du devis FirePro courant, sans toucher
  // au devis HT/BT (totalement indépendant).
  const razFirePro = () => {
    if (!window.confirm("Remettre à zéro les locaux de ce devis FirePro ? Cette action est irréversible.")) return;
    setZonesFirePro(devisFireProVierge().zonesFirePro);
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
        description: "",
        devisExterne: "",
      })),
    ]);
  // Import des pièces d'usure exportées par l'application de maintenance (RMP) : une ligne libre
  // « Composants » par pièce, prix d'achat laissé à 0 (à saisir), description détaillée.
  const importPiecesRef = useRef(null);
  const [catalogueLigneId, setCatalogueLigneId] = useState(null); // ligne « Composants » dont le catalogue de pièces est ouvert
  const importerPiecesUsure = (file) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      let data;
      try { data = JSON.parse(ev.target.result); } catch (e) { window.alert("Fichier illisible : ce n'est pas un export JSON valide."); return; }
      if (!data || data.format !== "ht-maintenance/pieces-usure" || !Array.isArray(data.pieces)) {
        window.alert("Ce fichier n'est pas un export « Pièces d'usure » de l'application de maintenance.");
        return;
      }
      if (!data.pieces.length) { window.alert("Le fichier ne contient aucune pièce."); return; }
      const statutTexte = { retard: "échéance dépassée", annee: "à remplacer cette année" };
      const nouvelles = data.pieces.map((p) => {
        const eq = p.equipement || {};
        const eqLabel = [eq.type, eq.repere && `(${eq.repere})`].filter(Boolean).join(" ");
        const piece = [p.marque, p.reference].filter(Boolean).join(" ");
        const echeance = p.echeance ? `échéance ${p.echeance}${statutTexte[p.statut] ? " — " + statutTexte[p.statut] : ""}` : "";
        return {
          id: uid(),
          famille: "composants",
          quantite: Number(p.quantite) > 0 ? Number(p.quantite) : 1,
          niveauTechnicien: "expert",
          typeJournee: "semaine",
          prixAchatUnitaire: 0,
          categorieId: "",
          libelleManuel: "",
          joursManuel: 1,
          prixJour: tarifs.expert.jour,
          description: `${eqLabel} — ${p.categorie || "Pièce d'usure"}${piece ? " : " + piece : ""}${p.description ? " — " + p.description : ""}${echeance ? " (" + echeance + ")" : ""}`,
          devisExterne: "",
        };
      });
      setLignesLibres((ls) => [...ls, ...nouvelles]);
      const s = data.site || {};
      if (s.client || s.site) setAffaire((a) => ({ ...a, client: a.client || s.client || "", site: a.site || s.site || "" }));
      window.alert(`${nouvelles.length} pièce(s) d'usure importée(s) dans les lignes libres « Composants ». Saisissez le prix d'achat unitaire de chacune.`);
    };
    reader.readAsText(file);
  };
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
    const heuresJourActif = affaire.heuresJour || 7; // repli pour les devis enregistrés avant ce réglage
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
            const joursHomme = (heures / heuresJourActif) * qte;
            const montantUnitaire = (heures / heuresJourActif) * tarif.jour * majoration.coef + item.amort;
            // Décomposition réelle des heures par type de temps (simple / complexe /
            // préparation), qu'importe si le niveau choisi est "1-4 complet" ou non —
            // sert au détail par niveau dans le récapitulatif.
            const joursSimple = (niveau === "complexe" ? 0 : item.heuresSimple / heuresJourActif) * qte;
            const joursComplexe = (niveau === "simple" ? 0 : item.heuresComplexe / heuresJourActif) * qte;
            const joursPrepa = (item.heuresPrepa / heuresJourActif) * qte;
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
  }, [postesEquipement, catalogueTemps, catalogueDirect, tarifs, majorations, affaire.heuresJour]);

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

    // --- Tableau détail équipements (6 colonnes) ---
    const equipW = colWidths([24, 13, 17, 15, 9, 22]);
    const headerCell = (text, i) =>
      new TableCell({ width: { size: equipW[i], type: WidthType.DXA }, shading: { type: ShadingType.SOLID, fill: "1B2733" }, children: [para(text, { bold: true, color: "FFFFFF" })] });
    const dataCell = (text, i, opts = {}) => new TableCell({ width: { size: equipW[i], type: WidthType.DXA }, children: [para(text, opts)] });

    const equipHeader = new TableRow({ children: ["Désignation", "Poste", "Famille", "Niveau", "Qté", "Montant HT"].map(headerCell) });
    const equipDataRows = [
      ...lignesCatalogue.map(
        (l) =>
          new TableRow({
            children: [
              dataCell(l.label, 0),
              dataCell(l.posteNom, 1),
              dataCell(LABEL_FAMILLE[l.famille] || l.famille, 2),
              dataCell(l.niveau ? NIVEAUX_PRESTATION[l.niveau] : "—", 3),
              dataCell(l.qte, 4, { right: true }),
              dataCell(euros(l.montant), 5, { right: true }),
            ],
          })
      ),
      ...lignesLibresCalc.map(
        ({ ligne: l, montant }) =>
          new TableRow({
            children: [
              dataCell(
                [libelleLigneLibre(l), l.description, l.devisExterne ? `Devis externe : ${l.devisExterne}` : ""].filter(Boolean).join(" — "),
                0
              ),
              dataCell("—", 1),
              dataCell(LABEL_FAMILLE[l.famille] || l.famille, 2),
              dataCell("—", 3),
              dataCell(l.quantite, 4, { right: true }),
              dataCell(euros(montant), 5, { right: true }),
            ],
          })
      ),
    ];
    if (equipDataRows.length === 0) {
      equipDataRows.push(new TableRow({ children: [new TableCell({ columnSpan: 6, children: [para("Aucun équipement renseigné")] })] }));
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

  // ---- Export Word (.docx) du récapitulatif FirePro ----
  async function exportFireProWord() {
    const dateStr = new Date().toLocaleDateString("fr-FR");
    const noBorder = { style: "none", size: 0, color: "FFFFFF" };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
    const PAGE_WIDTH = 9360;
    const colWidths = (percents) => percents.map((p) => Math.round((PAGE_WIDTH * p) / 100));
    const txt = (text, opts = {}) => new TextRun({ text: String(text), bold: !!opts.bold, color: opts.color });
    const para = (text, opts = {}) => new Paragraph({ alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [txt(text, opts)] });

    const infoW = colWidths([30, 70]);
    const infoTable = new Table({
      width: { size: PAGE_WIDTH, type: WidthType.DXA },
      columnWidths: infoW,
      rows: [
        ["Référence", fireproAffaire.reference || "—"],
        ["Client", fireproAffaire.client || "—"],
        ["Site", fireproAffaire.site || "—"],
        ["Date", dateStr],
      ].map(
        ([label, value]) =>
          new TableRow({
            children: [
              new TableCell({ width: { size: infoW[0], type: WidthType.DXA }, borders: noBorders, children: [para(label, { bold: true })] }),
              new TableCell({ width: { size: infoW[1], type: WidthType.DXA }, borders: noBorders, children: [para(value)] }),
            ],
          })
      ),
    });

    const detailW = colWidths([40, 40, 20]);
    const headerCell = (text) =>
      new TableCell({ width: { size: detailW[0], type: WidthType.DXA }, shading: { type: ShadingType.SOLID, fill: "1B2733" }, children: [para(text, { bold: true, color: "FFFFFF" })] });

    const zonesChildren = [];
    zonesFireProCalc.forEach(({ zone, classe, volume, masseNecessaire, masseEffective, montantTotal, suffisant }) => {
      zonesChildren.push(new Paragraph({ text: `${zone.nom}${suffisant ? "" : " — masse insuffisante"}`, heading: HeadingLevel.HEADING_2, spacing: { before: 240 } }));
      zonesChildren.push(
        para(
          `Classe de feu : ${classe.label} — Volume : ${volume.toFixed(2)} m³ — Masse nécessaire : ${Math.round(masseNecessaire)} g — Masse effective : ${Math.round(
            masseEffective
          )} g`
        )
      );

      const lignesGen = fireproGenerateurs
        .map((g) => ({ label: g.label, qte: zone.generateurs[g.id] || 0, prix: g.prix }))
        .filter((l) => l.qte > 0);
      const lignesAcc = [];
      Object.values(fireproAccessoires).forEach((fam) => {
        fam.items.forEach((item) => {
          const qte = zone.accessoires[item.id] || 0;
          if (qte > 0) lignesAcc.push({ label: item.label, qte, prix: item.prix });
        });
      });
      const toutesLignes = [...lignesGen, ...lignesAcc];

      const rows = [
        new TableRow({
          children: [headerCell("Désignation"), headerCell("Prix unitaire"), headerCell("Qté")],
        }),
        ...(toutesLignes.length > 0
          ? toutesLignes.map(
              (l) =>
                new TableRow({
                  children: [
                    new TableCell({ width: { size: detailW[0], type: WidthType.DXA }, children: [para(l.label)] }),
                    new TableCell({ width: { size: detailW[1], type: WidthType.DXA }, children: [para(euros(l.prix), { right: true })] }),
                    new TableCell({ width: { size: detailW[2], type: WidthType.DXA }, children: [para(l.qte, { right: true })] }),
                  ],
                })
            )
          : [new TableRow({ children: [new TableCell({ columnSpan: 3, children: [para("Aucun élément sélectionné")] })] })]),
      ];
      zonesChildren.push(new Table({ width: { size: PAGE_WIDTH, type: WidthType.DXA }, columnWidths: detailW, rows }));
      zonesChildren.push(
        new Paragraph({
          spacing: { before: 100, after: 100 },
          alignment: AlignmentType.RIGHT,
          children: [txt(`Sous-total ${zone.nom} : ${euros(montantTotal)}`, { bold: true })],
        })
      );
    });

    const recapW = colWidths([70, 30]);
    const recapRows = [
      ["Montant HT avant coefficient", euros(totalFireProAvantCoef), false],
      [`Coefficient d'ajustement (×${fireproCoefAjustement})`, euros(totalFirePro), false],
      ["TOTAL HT FIREPRO", euros(totalFirePro), true],
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

    const docWord = new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: "Récapitulatif FirePro — HT Maintenance", heading: HeadingLevel.HEADING_1 }),
            infoTable,
            new Paragraph({ text: "", spacing: { after: 100 } }),
            new Paragraph({ text: "Détail par local", heading: HeadingLevel.HEADING_2, spacing: { before: 100 } }),
            ...zonesChildren,
            new Paragraph({ text: "", spacing: { after: 100 } }),
            new Paragraph({ text: "Récapitulatif financier", heading: HeadingLevel.HEADING_2, spacing: { before: 100 } }),
            recapTable,
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(docWord);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `FirePro_${(fireproAffaire.reference || "FirePro").replace(/[^a-zA-Z0-9_-]/g, "_")}.docx`;
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
                <img src="/logo-ht-maintenance.png" alt="HT Maintenance" style={{ width: 34, height: 34, borderRadius: 8 }} className="shrink-0" />
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
            <img src="/logo-ht-maintenance.png" alt="HT Maintenance" style={{ width: 34, height: 34, borderRadius: 8 }} className="shrink-0" />
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
              <div className="flex items-end gap-6 flex-wrap">
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Cadre contractuel</label>
                  <Select
                    value={affaire.contrat}
                    onChange={(v) => setAffaire((a) => ({ ...a, contrat: v }))}
                    options={Object.entries(coefContrat).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                    style={{ maxWidth: 320 }}
                  />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Heures travaillées / jour</label>
                  <NumberField value={affaire.heuresJour || 7} onChange={(v) => setAffaire((a) => ({ ...a, heuresJour: v }))} suffix="h" />
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

            <SectionCard
              title="Client"
              subtitle="Base de contacts partagée avec toutes vos apps HT Maintenance"
              icon={FileText}
              right={
                <button onClick={ouvrirNouveauContact} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                  <Plus size={15} /> Nouveau client
                </button>
              }
            >
              <div className="flex items-end gap-3 flex-wrap mb-2">
                <div style={{ flex: 1, minWidth: 220 }}>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Sélectionner un client existant</label>
                  <Select
                    value={selectedContactId}
                    onChange={choisirContactPourAffaire}
                    options={[{ value: "", label: "— Choisir —" }, ...contacts.map((c) => ({ value: c.id, label: c.societe || c.contact || "(sans nom)" }))]}
                  />
                </div>
                {selectedContactId && (
                  <>
                    <button onClick={() => ouvrirEditionContact(selectedContactId)} className="px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: INK_2 }}>
                      Modifier
                    </button>
                    <button onClick={() => supprimerContact(selectedContactId)} className="px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: "#B0473E" }}>
                      Supprimer
                    </button>
                  </>
                )}
              </div>

              {editingContact && (
                <div style={{ border: `1px solid ${LINE}`, borderRadius: 8 }} className="p-4 mt-3">
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-3">
                    <div>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Société</label>
                      <TextField value={editingContact.societe} onChange={(v) => setEditingContact((c) => ({ ...c, societe: v }))} />
                    </div>
                    <div>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Contact</label>
                      <TextField value={editingContact.contact} onChange={(v) => setEditingContact((c) => ({ ...c, contact: v }))} placeholder="Nom de l'interlocuteur" />
                    </div>
                    <div>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Email</label>
                      <TextField value={editingContact.email} onChange={(v) => setEditingContact((c) => ({ ...c, email: v }))} />
                    </div>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
                    <div style={{ gridColumn: "span 2" }}>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Adresse</label>
                      <TextField value={editingContact.adresse} onChange={(v) => setEditingContact((c) => ({ ...c, adresse: v }))} />
                    </div>
                    <div>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Code postal</label>
                      <TextField value={editingContact.cp} onChange={(v) => setEditingContact((c) => ({ ...c, cp: v }))} />
                    </div>
                    <div>
                      <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Ville</label>
                      <TextField value={editingContact.ville} onChange={(v) => setEditingContact((c) => ({ ...c, ville: v }))} />
                    </div>
                  </div>
                  <div className="mb-4" style={{ maxWidth: 260 }}>
                    <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Pays (vide = France)</label>
                    <TextField value={editingContact.pays} onChange={(v) => setEditingContact((c) => ({ ...c, pays: v }))} />
                  </div>

                  <div className="flex items-center justify-between mb-2">
                    <label style={{ fontSize: 11, color: MUTED }}>Téléphones</label>
                    <button onClick={ajouterTelephoneDraft} className="flex items-center gap-1 text-xs font-medium" style={{ color: INK_2 }}>
                      <Plus size={13} /> Ajouter un numéro
                    </button>
                  </div>
                  <div className="flex flex-col gap-2 mb-4">
                    {(editingContact.telephones || []).map((tel, idx) => (
                      <div key={idx} className="grid grid-cols-4 gap-2 items-center">
                        <Select value={tel.type} onChange={(v) => updateTelephoneDraft(idx, { type: v })} options={[{ value: "Fixe", label: "Fixe" }, { value: "Mobile", label: "Mobile" }]} />
                        <div style={{ gridColumn: "span 2" }}>
                          <TextField value={tel.numero} onChange={(v) => updateTelephoneDraft(idx, { numero: v })} placeholder="Numéro" />
                        </div>
                        <button onClick={() => supprimerTelephoneDraft(idx)} style={{ color: "#B0473E" }} className="flex justify-end">
                          <Trash2 size={14} />
                        </button>
                      </div>
                    ))}
                    {(editingContact.telephones || []).length === 0 && <div style={{ fontSize: 12.5, color: MUTED }}>Aucun numéro renseigné.</div>}
                  </div>

                  <div className="flex items-center gap-2">
                    <button onClick={enregistrerContact} className="px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                      Enregistrer le contact
                    </button>
                    <button onClick={annulerEditionContact} className="px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: INK_2 }}>
                      Annuler
                    </button>
                  </div>
                </div>
              )}
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
                  <input
                    ref={importPiecesRef}
                    type="file"
                    accept=".json,application/json"
                    style={{ display: "none" }}
                    onChange={(e) => { if (e.target.files && e.target.files[0]) importerPiecesUsure(e.target.files[0]); e.target.value = ""; }}
                  />
                  <button
                    onClick={() => importPiecesRef.current && importPiecesRef.current.click()}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium"
                    style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}
                    title="Importer le fichier « Pièces d'usure » exporté depuis l'application de maintenance (prix à saisir)"
                  >
                    <Upload size={15} /> Importer pièces d'usure
                  </button>
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
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-3">
                        <div className="md:col-span-2">
                          <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Description (type de la ligne)</label>
                          <TextField
                            value={l.description || ""}
                            onChange={(v) => updateLigneLibre(l.id, { description: v })}
                            placeholder="Ex. relais Schneider 24V, peinture local transfo, location nacelle…"
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Devis externe (réf. ou lien)</label>
                          <TextField value={l.devisExterne || ""} onChange={(v) => updateLigneLibre(l.id, { devisExterne: v })} placeholder="N° de devis ou https://…" />
                        </div>
                      </div>
                      {l.famille === "composants" && (
                        <div className="mt-3">
                          <button onClick={() => setCatalogueLigneId(catalogueLigneId === l.id ? null : l.id)} style={catBtn(CAT_BLUE)} title="Condensateurs AC / DC, cartes condensateurs, ventilateurs : caractéristiques, dimensions, équivalences">
                            Catalogue pièces (condensateurs, ventilateurs)
                          </button>
                          {catalogueLigneId === l.id && (
                            <CataloguePiecesPicker
                              kind="tous"
                              onFermer={() => setCatalogueLigneId(null)}
                              onChoisir={(choix) => {
                                // Conserve le début d'une description issue de l'import (équipement — catégorie) et y ajoute la pièce choisie.
                                const avant = (l.description || "").split(" : ")[0];
                                const prefixe = avant && avant.includes(" — ") ? avant + " : " : "";
                                updateLigneLibre(l.id, { description: `${prefixe}${[choix.marque, choix.modele].filter(Boolean).join(" ")} — ${choix.description}` });
                                setCatalogueLigneId(null);
                              }}
                            />
                          )}
                        </div>
                      )}
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
                          {libelleLigneLibre(l)}
                          {l.description && <div style={{ color: MUTED, fontSize: 11.5 }}>{l.description}</div>}
                          {l.devisExterne && (
                            <div style={{ color: MUTED, fontSize: 11.5 }}>
                              Devis externe :{" "}
                              {/^https?:\/\//i.test(l.devisExterne) ? (
                                <a href={l.devisExterne} target="_blank" rel="noreferrer" style={{ color: INK_2, textDecoration: "underline" }}>
                                  {l.devisExterne}
                                </a>
                              ) : (
                                l.devisExterne
                              )}
                            </div>
                          )}
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

            <div className="flex items-center gap-3 mt-4 mb-1">
              <div style={{ flex: 1, height: 2, background: "linear-gradient(to right, transparent, #E0684A)" }} />
              <div className="flex items-center gap-2 px-3 py-1.5 rounded-full" style={{ background: "#FBEAE5" }}>
                <Flame size={15} color="#C6472E" strokeWidth={2.5} />
                <span style={{ fontSize: 13, fontWeight: 700, color: "#C6472E", letterSpacing: 0.3 }}>FIREPRO</span>
              </div>
              <div style={{ flex: 1, height: 2, background: "linear-gradient(to left, transparent, #E0684A)" }} />
            </div>

            <div style={{ background: "#FFFBF7", border: `1px solid #F0DCCF`, borderRadius: 10 }} className="p-4 mt-2 flex flex-wrap items-center gap-3 justify-between">
              <div className="flex items-center gap-2 flex-wrap">
                <span style={{ fontSize: 11, color: MUTED }}>Devis FirePro</span>
                <Select
                  value={currentFireproDevisId || ""}
                  onChange={(id) => chargerDevisFirePro(id)}
                  options={fireproDevisList.map((d) => ({ value: d.id, label: `${d.reference || "Sans référence"}${d.client ? " — " + d.client : ""}` }))}
                  style={{ minWidth: 240, width: "auto" }}
                />
              </div>
              <button onClick={exportFireProWord} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                <Download size={15} /> Exporter en Word (FirePro)
              </button>
            </div>

            {zonesFireProCalc.map(({ zone, montantTotal }) => {
              const lignesGen = fireproGenerateurs.map((g) => ({ id: g.id, label: g.label, qte: zone.generateurs[g.id] || 0, prix: g.prix })).filter((l) => l.qte > 0);
              const lignesAcc = [];
              Object.entries(fireproAccessoires).forEach(([famId, fam]) => {
                fam.items.forEach((item) => {
                  const qte = zone.accessoires[item.id] || 0;
                  if (qte > 0) lignesAcc.push({ id: `${famId}-${item.id}`, label: item.label, qte, prix: item.prix });
                });
              });
              const toutesLignes = [...lignesGen, ...lignesAcc];
              return (
                <SectionCard
                  key={zone.id}
                  title={zone.nom}
                  subtitle={`${toutesLignes.length} élément${toutesLignes.length > 1 ? "s" : ""} sélectionné${toutesLignes.length > 1 ? "s" : ""}`}
                  icon={Flame}
                  bg="#FFFBF7"
                >
                  <div className="overflow-x-auto">
                    <table className="w-full" style={{ fontSize: 13 }}>
                      <thead>
                        <tr style={{ color: MUTED, textAlign: "left" }}>
                          <th className="pb-2 font-medium">Désignation</th>
                          <th className="pb-2 font-medium text-right">Prix unitaire</th>
                          <th className="pb-2 font-medium text-right">Qté</th>
                          <th className="pb-2 font-medium text-right">Montant HT</th>
                        </tr>
                      </thead>
                      <tbody>
                        {toutesLignes.map((l) => (
                          <tr key={l.id} style={{ borderTop: `1px solid ${LINE}` }}>
                            <td className="py-1.5" style={{ color: INK }}>{l.label}</td>
                            <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{euros(l.prix)}</td>
                            <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{l.qte}</td>
                            <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{euros(l.prix * l.qte)}</td>
                          </tr>
                        ))}
                        {toutesLignes.length === 0 && (
                          <tr>
                            <td colSpan={4} className="py-4 text-center" style={{ color: MUTED }}>Aucun élément sélectionné.</td>
                          </tr>
                        )}
                        <tr style={{ borderTop: `2px solid ${INK}` }}>
                          <td className="pt-2" style={{ fontWeight: 700, color: INK }} colSpan={3}>Sous-total {zone.nom}</td>
                          <td className="pt-2 text-right" style={{ fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{euros(montantTotal)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </SectionCard>
              );
            })}

            <SectionCard title="Récapitulatif FirePro" icon={FileText} bg="#FFFBF7">
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

        {/* ---------------- ONGLET FIREPRO ---------------- */}
        {tab === "firepro" && (
          <>
            <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="p-4 flex flex-wrap items-center gap-3 justify-between">
              <div className="flex items-center gap-2 flex-wrap">
                <span style={{ fontSize: 11, color: MUTED }}>FirePro</span>
                <Select
                  value={currentFireproDevisId || ""}
                  onChange={(id) => chargerDevisFirePro(id)}
                  options={fireproDevisList.map((d) => ({ value: d.id, label: `${d.reference || "Sans référence"}${d.client ? " — " + d.client : ""}` }))}
                  style={{ minWidth: 240, width: "auto" }}
                />
              </div>
              <div className="flex items-center gap-2">
                <button onClick={nouveauChiffrageFirePro} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: AMBER, color: INK }}>
                  <Plus size={15} /> Nouveau
                </button>
                <button onClick={dupliquerChiffrageFirePro} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: INK_2 }}>
                  <Copy size={14} /> Dupliquer
                </button>
                {fireproDevisList.length > 1 && (
                  <button onClick={supprimerChiffrageFirePro} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: "#B0473E" }}>
                    <Trash2 size={14} /> Supprimer
                  </button>
                )}
                <button onClick={razFirePro} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ border: `1px solid ${LINE}`, color: "#B0473E" }}>
                  <Trash2 size={14} /> RAZ
                </button>
              </div>
            </div>

            <SectionCard title="Informations de l'affaire FirePro" icon={FileText}>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Référence</label>
                  <TextField value={fireproAffaire.reference} onChange={(v) => setFireproAffaire((a) => ({ ...a, reference: v }))} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Client</label>
                  <TextField value={fireproAffaire.client} onChange={(v) => setFireproAffaire((a) => ({ ...a, client: v }))} placeholder="Nom du client" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: MUTED, display: "block", marginBottom: 2 }}>Site</label>
                  <TextField value={fireproAffaire.site} onChange={(v) => setFireproAffaire((a) => ({ ...a, site: v }))} placeholder="Site / adresse" />
                </div>
              </div>
            </SectionCard>

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
              <div className="mt-4" style={{ fontSize: 11.5, color: MUTED }}>
                Les heures travaillées par jour se règlent désormais par devis (onglet Chiffrage → Informations de l'affaire), puisqu'elles peuvent différer d'un chiffrage à l'autre.
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
                        const prixRevient = (heuresDefaut / (affaire.heuresJour || 7)) * tarifs.expert.jour + item.amort;
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
                        {cat.items.map((item) => {
                          const majBatterie = (patch) =>
                            setCatalogueDirect((s) => ({
                              ...s,
                              batteries: {
                                ...s.batteries,
                                items: s.batteries.items.map((it) => {
                                  if (it.id !== item.id) return it;
                                  const next = { ...it, ...patch };
                                  next.prix = next.prixAchat * next.coef;
                                  return next;
                                }),
                              },
                            }));
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
                        {cat.items.map((item) => (
                          <div key={item.id} className="grid grid-cols-4 gap-3 items-center">
                            <div style={{ gridColumn: "span 2" }}>
                              <TextField value={item.label} onChange={(v) => renommerLigneCatalogueDirect(key, item.id, v)} />
                            </div>
                            <NumberField
                              value={item.prix}
                              onChange={(v) =>
                                setCatalogueDirect((s) => ({ ...s, [key]: { ...s[key], items: s[key].items.map((it) => (it.id === item.id ? { ...it, prix: v } : it)) } }))
                              }
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
                  {fireproGenerateurs.map((g) => (
                    <div key={g.id} className="grid grid-cols-5 gap-3 items-center">
                      <div style={{ gridColumn: "span 2" }}>
                        <TextField value={g.label} onChange={(v) => renommerGenerateurFirePro(g.id, v)} />
                      </div>
                      <NumberField value={g.masse} onChange={(v) => majMasseGenerateurFirePro(g.id, v)} width="100%" />
                      <NumberField value={g.prix} onChange={(v) => setFireproGenerateurs((arr) => arr.map((x) => (x.id === g.id ? { ...x, prix: v } : x)))} suffix="€" width="100%" />
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
                      {fam.items.map((item) => (
                        <div key={item.id} className="grid grid-cols-4 gap-3 items-center">
                          <div style={{ gridColumn: "span 2" }}>
                            <TextField value={item.label} onChange={(v) => renommerAccessoireFirePro(famId, item.id, v)} />
                          </div>
                          <NumberField
                            value={item.prix}
                            onChange={(v) =>
                              setFireproAccessoires((s) => ({ ...s, [famId]: { ...s[famId], items: s[famId].items.map((it) => (it.id === item.id ? { ...it, prix: v } : it)) } }))
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
