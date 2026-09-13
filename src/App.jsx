import React, { useState, useMemo } from "react";
import { Plus, Trash2, Settings2, FileText, ClipboardList, ChevronDown, ChevronUp, Zap } from "lucide-react";

// ---------------------------------------------------------------------------
// HT MAINTENANCE — OUTIL DE CHIFFRAGE
// Palette : ardoise/graphite + ambre (couleurs d'atelier électrique HT/BT)
// ---------------------------------------------------------------------------

const INK = "#1B2733";      // graphite / ardoise
const INK_2 = "#2E3D4E";
const AMBER = "#E8A33D";    // ambre — accent unique
const PAPER = "#F5F6F4";
const LINE = "#DCE0E3";

const uid = () => Math.random().toString(36).slice(2, 10);

// ---------------------------------------------------------------------------
// Données de référence par défaut (issues du fichier QUO-CHIFFRAGE HT-BT
// d'origine, adaptées et simplifiées — tout est modifiable dans l'onglet
// Paramètres / Catalogue)
// ---------------------------------------------------------------------------

const DEFAULT_TARIFS = {
  technicien: { label: "Technicien", jour: 850, demiJour: 480 },
  expert: { label: "Expert HT/BT", jour: 1050, demiJour: 600 },
  ingenieur: { label: "Ingénieur", jour: 1250, demiJour: 750 },
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

// Catalogue "temps" (jours-hommes par équipement) — HTA / BT / Onduleurs
const DEFAULT_CATALOGUE_TEMPS = {
  hta: {
    label: "Cellules HTA",
    marquesRef: "ABB, Schneider Electric (SM6, FLUOKIT, PIX...), Siemens, CEM Gardy, Ormazabal, Pommier",
    items: [
      { id: "hta-inter-sect", label: "Interrupteur-sectionneur (IM, IS)", jours: 0.3 },
      { id: "hta-inter-fusible", label: "Interrupteur-fusibles (QM)", jours: 0.35 },
      { id: "hta-disjoncteur", label: "Cellule disjoncteur (DM)", jours: 0.45 },
      { id: "hta-arrivee-depart", label: "Arrivée / Départ transfo", jours: 0.3 },
    ],
  },
  bt: {
    label: "Disjoncteurs BT",
    marquesRef: "Schneider (Masterpact, Compact NS/NSX), Eaton/Moeller, ABB SACE, Siemens, Legrand, GE",
    items: [
      { id: "bt-injection-primaire", label: "Injection primaire (débrochable)", jours: 0.5 },
      { id: "bt-injection-secondaire", label: "Injection secondaire (valise de test)", jours: 0.4 },
    ],
  },
  onduleur: {
    label: "Onduleurs",
    marquesRef: "TECO et autres constructeurs",
    items: [
      { id: "ond-teco", label: "Onduleur TECO < 5 kVA", jours: 0.25 },
      { id: "ond-autre", label: "Onduleur autre marque", jours: 0.35 },
    ],
  },
};

// Catalogue "fournitures à coefficient" — batteries / composants (prix d'achat x coef)
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

// Analyses d'huile — prix directs (tout inclus)
const DEFAULT_ANALYSES_HUILE = [
  { id: "ah-pa", label: "Analyse d'huile N° PA (TP1)", prix: 257.4 },
  { id: "ah-b", label: "Analyse d'huile N° B (TP3)", prix: 378.4 },
  { id: "ah-bdf", label: "Analyse d'huile N° B + DF (TP4)", prix: 552.2 },
  { id: "ah-cdf", label: "Analyse d'huile N° C + DF (Expert)", prix: 744 },
  { id: "ah-df", label: "Analyse d'huile N° DF", prix: 286 },
  { id: "ah-soufre", label: "Analyse soufre corrosif", prix: 154 },
  { id: "ah-tangente", label: "Analyse tangente delta", prix: 165 },
  { id: "ah-pcb", label: "Recherche PCB", prix: 156 },
  { id: "ah-prelevement", label: "Prélèvement d'huile", prix: 160 },
];

const FAMILLES = [
  { id: "hta", label: "Cellules HTA", type: "temps" },
  { id: "bt", label: "Disjoncteurs BT", type: "temps" },
  { id: "onduleur", label: "Onduleurs", type: "temps" },
  { id: "batteries", label: "Batteries", type: "coef" },
  { id: "composants", label: "Composants", type: "coef" },
  { id: "analyseHuile", label: "Analyses d'huile", type: "direct" },
  { id: "manuel", label: "Poste libre (manuel)", type: "manuel" },
];

function euros(n) {
  if (!isFinite(n)) return "0 €";
  return n.toLocaleString("fr-FR", { maximumFractionDigits: 0 }) + " €";
}

function findCoefCategory(catalogue, valeur) {
  return catalogue.categories.find((c) => valeur >= c.min && valeur <= c.max) || catalogue.categories[0];
}

// ---------------------------------------------------------------------------

function SectionCard({ title, icon: Icon, children, right }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 10 }} className="overflow-hidden">
      <div
        className="flex items-center justify-between px-5 py-4"
        style={{ borderBottom: `1px solid ${LINE}` }}
      >
        <div className="flex items-center gap-2">
          {Icon && <Icon size={17} color={INK_2} />}
          <h2 style={{ color: INK, fontWeight: 600, fontSize: 15 }}>{title}</h2>
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
      {suffix && <span style={{ fontSize: 12, color: "#8A93A0" }}>{suffix}</span>}
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
      style={{
        border: `1px solid ${LINE}`,
        borderRadius: 6,
        padding: "6px 9px",
        fontSize: 13,
        color: INK,
        width: "100%",
        ...style,
      }}
    />
  );
}

function Select({ value, onChange, options, style }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        border: `1px solid ${LINE}`,
        borderRadius: 6,
        padding: "6px 9px",
        fontSize: 13,
        color: INK,
        background: "#fff",
        width: "100%",
        ...style,
      }}
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
  const [paramsOpen, setParamsOpen] = useState(false);

  const [tarifs, setTarifs] = useState(DEFAULT_TARIFS);
  const [majorations, setMajorations] = useState(DEFAULT_MAJORATIONS);
  const [degressivite, setDegressivite] = useState(DEFAULT_DEGRESSIVITE);
  const [coefContrat, setCoefContrat] = useState(DEFAULT_COEF_CONTRAT);
  const [heuresJour, setHeuresJour] = useState(7);

  const [catalogueTemps, setCatalogueTemps] = useState(DEFAULT_CATALOGUE_TEMPS);
  const [catalogueCoef, setCatalogueCoef] = useState(DEFAULT_CATALOGUE_COEF);
  const [analysesHuile, setAnalysesHuile] = useState(DEFAULT_ANALYSES_HUILE);

  const [affaire, setAffaire] = useState({
    client: "",
    site: "",
    reference: "QUO-" + new Date().getFullYear() + "-001",
    contrat: "aucun",
  });

  const [postes, setPostes] = useState([
    {
      id: uid(),
      famille: "hta",
      sousType: DEFAULT_CATALOGUE_TEMPS.hta.items[0].id,
      quantite: 1,
      niveauTechnicien: "technicien",
      typeJournee: "semaine",
      joursUnitaire: DEFAULT_CATALOGUE_TEMPS.hta.items[0].jours,
      prixAchatUnitaire: 20,
      libelleManuel: "",
      joursManuel: 1,
      commentaire: "",
    },
  ]);

  // ---- Helpers de mise à jour ----
  const updatePoste = (id, patch) =>
    setPostes((ps) => ps.map((p) => (p.id === id ? { ...p, ...patch } : p)));

  const addPoste = () =>
    setPostes((ps) => [
      ...ps,
      {
        id: uid(),
        famille: "hta",
        sousType: DEFAULT_CATALOGUE_TEMPS.hta.items[0].id,
        quantite: 1,
        niveauTechnicien: "technicien",
        typeJournee: "semaine",
        joursUnitaire: DEFAULT_CATALOGUE_TEMPS.hta.items[0].jours,
        prixAchatUnitaire: 20,
        libelleManuel: "",
        joursManuel: 1,
        commentaire: "",
      },
    ]);

  const removePoste = (id) => setPostes((ps) => ps.filter((p) => p.id !== id));

  // ---- Calcul d'un poste ----
  function computePoste(p) {
    const famille = FAMILLES.find((f) => f.id === p.famille);
    const tarif = tarifs[p.niveauTechnicien];
    const majoration = majorations[p.typeJournee];
    let montant = 0;
    let joursHomme = 0;
    let detailUnitaire = "";

    if (famille.type === "temps") {
      const cat = catalogueTemps[p.famille];
      const item = cat.items.find((i) => i.id === p.sousType) || cat.items[0];
      joursHomme = item.jours * p.quantite;
      montant = joursHomme * tarif.jour * majoration.coef;
      detailUnitaire = `${item.jours} j/u × ${p.quantite} × ${euros(tarif.jour)}/j × ${majoration.coef}`;
    } else if (famille.type === "coef") {
      const cat = catalogueCoef[p.famille];
      const catgorie = findCoefCategory(cat, p.prixAchatUnitaire);
      montant = p.prixAchatUnitaire * catgorie.coef * p.quantite;
      detailUnitaire = `${euros(p.prixAchatUnitaire)} × coef ${catgorie.coef} (${catgorie.label}) × ${p.quantite}`;
    } else if (famille.type === "direct") {
      const item = analysesHuile.find((i) => i.id === p.sousType) || analysesHuile[0];
      montant = item.prix * p.quantite;
      detailUnitaire = `${euros(item.prix)} × ${p.quantite}`;
    } else if (famille.type === "manuel") {
      joursHomme = p.joursManuel * p.quantite;
      montant = joursHomme * tarif.jour * majoration.coef;
      detailUnitaire = `${p.joursManuel} j/u × ${p.quantite} × ${euros(tarif.jour)}/j × ${majoration.coef}`;
    }
    return { montant, joursHomme, detailUnitaire };
  }

  const lignes = useMemo(() => postes.map((p) => ({ poste: p, ...computePoste(p) })), [
    postes,
    tarifs,
    majorations,
    catalogueTemps,
    catalogueCoef,
    analysesHuile,
  ]);

  const totalEquipements = postes.reduce((s, p) => s + (Number(p.quantite) || 0), 0);
  const totalJoursHomme = lignes.reduce((s, l) => s + l.joursHomme, 0);
  const totalAvantCoef = lignes.reduce((s, l) => s + l.montant, 0);

  const palierDegressif =
    degressivite.find((d) => totalEquipements > d.min && totalEquipements <= d.max) ||
    degressivite[degressivite.length - 1];
  const coefContratActif = coefContrat[affaire.contrat] || { coef: 1, label: "—" };

  const totalApresDegressivite = totalAvantCoef * palierDegressif.coef;
  const totalFinal = totalApresDegressivite * coefContratActif.coef;

  const parFamille = useMemo(() => {
    const map = {};
    lignes.forEach((l) => {
      const key = l.poste.famille;
      map[key] = (map[key] || 0) + l.montant;
    });
    return map;
  }, [lignes]);

  const tabBtn = (id, label, Icon) => (
    <button
      onClick={() => setTab(id)}
      className="flex items-center gap-2 px-4 py-2.5 text-sm transition-colors"
      style={{
        color: tab === id ? INK : "#8A93A0",
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
            <div
              style={{ background: AMBER, width: 34, height: 34, borderRadius: 8 }}
              className="flex items-center justify-center"
            >
              <Zap size={18} color={INK} strokeWidth={2.5} />
            </div>
            <div>
              <div style={{ color: "#fff", fontWeight: 700, fontSize: 16, letterSpacing: 0.2 }}>
                HT Maintenance
              </div>
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
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                <div>
                  <label style={{ fontSize: 11, color: "#8A93A0" }}>Référence</label>
                  <TextField value={affaire.reference} onChange={(v) => setAffaire((a) => ({ ...a, reference: v }))} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#8A93A0" }}>Client</label>
                  <TextField value={affaire.client} onChange={(v) => setAffaire((a) => ({ ...a, client: v }))} placeholder="Nom du client" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#8A93A0" }}>Site</label>
                  <TextField value={affaire.site} onChange={(v) => setAffaire((a) => ({ ...a, site: v }))} placeholder="Site / adresse" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#8A93A0" }}>Cadre contractuel</label>
                  <Select
                    value={affaire.contrat}
                    onChange={(v) => setAffaire((a) => ({ ...a, contrat: v }))}
                    options={Object.entries(coefContrat).map(([k, v]) => ({ value: k, label: `${v.label} (×${v.coef})` }))}
                  />
                </div>
              </div>
            </SectionCard>

            <SectionCard
              title={`Postes du chiffrage (${postes.length})`}
              icon={ClipboardList}
              right={
                <button
                  onClick={addPoste}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium"
                  style={{ background: AMBER, color: INK }}
                >
                  <Plus size={15} /> Ajouter un poste
                </button>
              }
            >
              <div className="flex flex-col gap-3">
                {lignes.map(({ poste: p, montant, detailUnitaire }, idx) => {
                  const famille = FAMILLES.find((f) => f.id === p.famille);
                  return (
                    <div
                      key={p.id}
                      style={{ border: `1px solid ${LINE}`, borderRadius: 8 }}
                      className="p-4"
                    >
                      <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
                        <div className="md:col-span-2">
                          <label style={{ fontSize: 11, color: "#8A93A0" }}>Famille</label>
                          <Select
                            value={p.famille}
                            onChange={(v) => {
                              const fam = FAMILLES.find((f) => f.id === v);
                              const patch = { famille: v };
                              if (fam.type === "temps") {
                                patch.sousType = catalogueTemps[v].items[0].id;
                                patch.joursUnitaire = catalogueTemps[v].items[0].jours;
                              } else if (fam.type === "direct") {
                                patch.sousType = analysesHuile[0].id;
                              }
                              updatePoste(p.id, patch);
                            }}
                            options={FAMILLES.map((f) => ({ value: f.id, label: f.label }))}
                          />
                        </div>

                        {famille.type === "temps" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: "#8A93A0" }}>Type d'équipement</label>
                            <Select
                              value={p.sousType}
                              onChange={(v) => updatePoste(p.id, { sousType: v })}
                              options={catalogueTemps[p.famille].items.map((i) => ({ value: i.id, label: i.label }))}
                            />
                          </div>
                        )}

                        {famille.type === "direct" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: "#8A93A0" }}>Type d'analyse</label>
                            <Select
                              value={p.sousType}
                              onChange={(v) => updatePoste(p.id, { sousType: v })}
                              options={analysesHuile.map((i) => ({ value: i.id, label: i.label }))}
                            />
                          </div>
                        )}

                        {famille.type === "coef" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: "#8A93A0" }}>
                              {catalogueCoef[p.famille].critereLabel}
                            </label>
                            <NumberField
                              value={p.prixAchatUnitaire}
                              onChange={(v) => updatePoste(p.id, { prixAchatUnitaire: v })}
                              width="100%"
                            />
                          </div>
                        )}

                        {famille.type === "manuel" && (
                          <div className="md:col-span-3">
                            <label style={{ fontSize: 11, color: "#8A93A0" }}>Libellé</label>
                            <TextField
                              value={p.libelleManuel}
                              onChange={(v) => updatePoste(p.id, { libelleManuel: v })}
                              placeholder="Description du poste"
                            />
                          </div>
                        )}

                        <div className="md:col-span-1">
                          <label style={{ fontSize: 11, color: "#8A93A0" }}>Qté</label>
                          <NumberField value={p.quantite} onChange={(v) => updatePoste(p.id, { quantite: v })} width="100%" />
                        </div>

                        {famille.type === "manuel" && (
                          <div className="md:col-span-1">
                            <label style={{ fontSize: 11, color: "#8A93A0" }}>j/u</label>
                            <NumberField value={p.joursManuel} onChange={(v) => updatePoste(p.id, { joursManuel: v })} width="100%" />
                          </div>
                        )}

                        {(famille.type === "temps" || famille.type === "manuel") && (
                          <>
                            <div className="md:col-span-2">
                              <label style={{ fontSize: 11, color: "#8A93A0" }}>Technicien</label>
                              <Select
                                value={p.niveauTechnicien}
                                onChange={(v) => updatePoste(p.id, { niveauTechnicien: v })}
                                options={Object.entries(tarifs).map(([k, v]) => ({ value: k, label: v.label }))}
                              />
                            </div>
                            <div className="md:col-span-2">
                              <label style={{ fontSize: 11, color: "#8A93A0" }}>Journée</label>
                              <Select
                                value={p.typeJournee}
                                onChange={(v) => updatePoste(p.id, { typeJournee: v })}
                                options={Object.entries(majorations).map(([k, v]) => ({ value: k, label: v.label }))}
                              />
                            </div>
                          </>
                        )}

                        <div className="md:col-span-1 flex justify-end">
                          <button onClick={() => removePoste(p.id)} style={{ color: "#B0473E" }}>
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </div>
                      <div className="flex items-center justify-between mt-3 pt-3" style={{ borderTop: `1px dashed ${LINE}` }}>
                        <span style={{ fontSize: 12, color: "#8A93A0" }}>{detailUnitaire}</span>
                        <span style={{ fontSize: 15, fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>
                          {euros(montant)}
                        </span>
                      </div>
                    </div>
                  );
                })}
                {postes.length === 0 && (
                  <div style={{ color: "#8A93A0", fontSize: 13 }} className="text-center py-6">
                    Aucun poste. Cliquez sur « Ajouter un poste » pour commencer le chiffrage.
                  </div>
                )}
              </div>
            </SectionCard>
          </>
        )}

        {/* ---------------- ONGLET RECAP ---------------- */}
        {tab === "recap" && (
          <>
            <SectionCard title="Résumé de l'affaire" icon={FileText}>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-2">
                <div>
                  <div style={{ fontSize: 11, color: "#8A93A0" }}>Référence</div>
                  <div style={{ fontWeight: 600, color: INK }}>{affaire.reference || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "#8A93A0" }}>Client</div>
                  <div style={{ fontWeight: 600, color: INK }}>{affaire.client || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "#8A93A0" }}>Site</div>
                  <div style={{ fontWeight: 600, color: INK }}>{affaire.site || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "#8A93A0" }}>Nb postes</div>
                  <div style={{ fontWeight: 600, color: INK }}>{postes.length}</div>
                </div>
              </div>
            </SectionCard>

            <SectionCard title="Répartition par famille" icon={ClipboardList}>
              <table className="w-full" style={{ fontSize: 13 }}>
                <thead>
                  <tr style={{ color: "#8A93A0", textAlign: "left" }}>
                    <th className="pb-2 font-medium">Famille</th>
                    <th className="pb-2 font-medium text-right">Montant HT</th>
                    <th className="pb-2 font-medium text-right">Part</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(parFamille).map(([key, montant]) => (
                    <tr key={key} style={{ borderTop: `1px solid ${LINE}` }}>
                      <td className="py-2" style={{ color: INK }}>
                        {FAMILLES.find((f) => f.id === key)?.label || key}
                      </td>
                      <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                        {euros(montant)}
                      </td>
                      <td className="py-2 text-right" style={{ color: "#8A93A0" }}>
                        {totalAvantCoef > 0 ? Math.round((montant / totalAvantCoef) * 100) : 0}%
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </SectionCard>

            <SectionCard title="Récapitulatif financier" icon={FileText}>
              <div className="flex flex-col gap-2.5" style={{ fontSize: 13.5 }}>
                <Row label="Nombre total d'équipements" value={totalEquipements} plain />
                <Row label="Total jours-hommes" value={totalJoursHomme.toFixed(2) + " j"} plain />
                <Row label="Montant HT avant coefficients" value={euros(totalAvantCoef)} />
                <Row
                  label={`Dégressivité volume (${palierDegressif.label}, ×${palierDegressif.coef})`}
                  value={euros(totalApresDegressivite)}
                />
                <Row
                  label={`Coefficient contractuel — ${coefContratActif.label} (×${coefContratActif.coef})`}
                  value={euros(totalFinal)}
                />
                <div style={{ borderTop: `2px solid ${INK}`, marginTop: 6, paddingTop: 12 }} className="flex items-center justify-between">
                  <span style={{ fontWeight: 700, color: INK, fontSize: 16 }}>TOTAL HT OFFRE</span>
                  <span style={{ fontWeight: 800, color: INK, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>
                    {euros(totalFinal)}
                  </span>
                </div>
              </div>
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
                      <label style={{ fontSize: 11, color: "#8A93A0" }}>Prix / jour</label>
                      <NumberField
                        value={t.jour}
                        onChange={(v) => setTarifs((s) => ({ ...s, [key]: { ...s[key], jour: v } }))}
                        suffix="€"
                        width="100%"
                      />
                      <label style={{ fontSize: 11, color: "#8A93A0" }}>Prix / demi-journée</label>
                      <NumberField
                        value={t.demiJour}
                        onChange={(v) => setTarifs((s) => ({ ...s, [key]: { ...s[key], demiJour: v } }))}
                        suffix="€"
                        width="100%"
                      />
                    </div>
                  </div>
                ))}
              </div>
              <div className="mt-4 flex items-center gap-2">
                <label style={{ fontSize: 12, color: "#8A93A0" }}>Heures travaillées par jour</label>
                <NumberField value={heuresJour} onChange={setHeuresJour} suffix="h" />
              </div>
            </SectionCard>

            <SectionCard title="Coefficients de majoration (jour d'intervention)" icon={Settings2}>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                {Object.entries(majorations).map(([key, m]) => (
                  <div key={key}>
                    <label style={{ fontSize: 11, color: "#8A93A0" }}>{m.label}</label>
                    <NumberField
                      value={m.coef}
                      onChange={(v) => setMajorations((s) => ({ ...s, [key]: { ...s[key], coef: v } }))}
                      suffix="×"
                      width="100%"
                    />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Dégressivité selon le volume d'équipements" icon={Settings2}>
              <div className="flex flex-col gap-2">
                {degressivite.map((d, i) => (
                  <div key={i} className="grid grid-cols-3 gap-3 items-center">
                    <span style={{ fontSize: 13, color: INK }}>{d.label}</span>
                    <NumberField
                      value={d.coef}
                      onChange={(v) =>
                        setDegressivite((arr) => arr.map((x, j) => (j === i ? { ...x, coef: v } : x)))
                      }
                      suffix="×"
                    />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Coefficients contractuels" icon={Settings2}>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {Object.entries(coefContrat).map(([key, c]) => (
                  <div key={key} className="flex items-center justify-between gap-3">
                    <span style={{ fontSize: 13, color: INK }}>{c.label}</span>
                    <NumberField
                      value={c.coef}
                      onChange={(v) => setCoefContrat((s) => ({ ...s, [key]: { ...s[key], coef: v } }))}
                      suffix="×"
                    />
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Catalogue équipements — temps de maintenance (jours-hommes)" icon={ClipboardList}>
              <div className="flex flex-col gap-5">
                {Object.entries(catalogueTemps).map(([key, cat]) => (
                  <div key={key}>
                    <div style={{ fontWeight: 600, color: INK, fontSize: 13 }}>{cat.label}</div>
                    <div style={{ fontSize: 11.5, color: "#8A93A0", marginBottom: 8 }}>{cat.marquesRef}</div>
                    <div className="flex flex-col gap-1.5">
                      {cat.items.map((item, i) => (
                        <div key={item.id} className="grid grid-cols-3 gap-3 items-center">
                          <span style={{ fontSize: 13, color: INK, gridColumn: "span 2" }}>{item.label}</span>
                          <NumberField
                            value={item.jours}
                            onChange={(v) =>
                              setCatalogueTemps((s) => ({
                                ...s,
                                [key]: {
                                  ...s[key],
                                  items: s[key].items.map((it, j) => (j === i ? { ...it, jours: v } : it)),
                                },
                              }))
                            }
                            suffix="j/u"
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
                            onChange={(v) =>
                              setCatalogueCoef((s) => ({
                                ...s,
                                [key]: {
                                  ...s[key],
                                  categories: s[key].categories.map((x, j) => (j === i ? { ...x, coef: v } : x)),
                                },
                              }))
                            }
                            suffix="×"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="Analyses d'huile (prix directs)" icon={ClipboardList}>
              <div className="flex flex-col gap-1.5">
                {analysesHuile.map((item, i) => (
                  <div key={item.id} className="grid grid-cols-3 gap-3 items-center">
                    <span style={{ fontSize: 13, color: INK, gridColumn: "span 2" }}>{item.label}</span>
                    <NumberField
                      value={item.prix}
                      onChange={(v) => setAnalysesHuile((arr) => arr.map((x, j) => (j === i ? { ...x, prix: v } : x)))}
                      suffix="€"
                    />
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

function Row({ label, value, plain }) {
  return (
    <div className="flex items-center justify-between">
      <span style={{ color: plain ? "#8A93A0" : INK }}>{label}</span>
      <span style={{ fontWeight: plain ? 500 : 700, color: INK, fontVariantNumeric: "tabular-nums" }}>{value}</span>
    </div>
  );
}
