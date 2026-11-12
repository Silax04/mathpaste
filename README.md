# MathPaste — LaTeX → équations natives Microsoft Word

MathPaste est un complément Word (Office Add-in) qui convertit du LaTeX (ChatGPT, Claude, Overleaf…) en **véritables équations Word (OMML)**, éditables avec l'éditeur d'équations de Word. Jamais d'images, jamais de LaTeX brut.

Deux méthodes, **un seul moteur de conversion** :

| | Méthode A — Panneau | Méthode B — Conversion directe |
|---|---|---|
| Usage | Coller du LaTeX dans le panneau, prévisualiser, **Insérer dans Word** | Sélectionner du LaTeX dans le document, **Convert to Formula** |
| Accès | Onglet **MathPaste › Panneau MathPaste** | Onglet **MathPaste › Convert to Formula**, **clic droit › Convert to Formula**, ou **Ctrl+Alt+Maj+M** |
| Résultat | Équations insérées au curseur | Seules les expressions détectées dans la sélection sont remplacées |

---

## 1. Installation (Word Desktop Windows)

Prérequis : Microsoft 365 / Word 2021+ sur Windows, [Node.js](https://nodejs.org) 18+.

```bash
npm install
```

### Certificat HTTPS local (une seule fois)

Word n'accepte que des compléments servis en HTTPS. Ce script installe un certificat de développement `localhost` (Windows vous demande de confirmer) :

```bash
npm run certs
```

### Lancer en développement (sideloading automatique)

```bash
npm start
```

Cette commande démarre le serveur `https://localhost:3000`, enregistre le manifeste auprès de Word (clé de registre développeur) et ouvre Word avec le complément chargé. Un onglet **MathPaste** apparaît dans le ruban.

Arrêter et désenregistrer :

```bash
npm run stop
```

### Sideloading manuel (alternative)

1. Lancer uniquement le serveur : `npm run dev-server`
2. Créer un dossier partagé (ex. `C:\AddinCatalog`, clic droit › Propriétés › Partage › Partager) et y copier `manifest.xml`.
3. Word › Fichier › Options › Centre de gestion de la confidentialité › Paramètres › **Catalogues de compléments approuvés** : ajouter le chemin réseau `\\NOM-PC\AddinCatalog`, cocher **Afficher dans le menu**, redémarrer Word.
4. Accueil › **Compléments** › **Plus de compléments** › **DOSSIER PARTAGÉ** › MathPaste.

---

## 2. Utilisation

### Méthode A — Panneau MathPaste
1. Onglet **MathPaste › Panneau MathPaste**.
2. Collez votre LaTeX (une ou plusieurs expressions ; délimiteurs `$…$`, `$$…$$`, `\[…\]`, `\(…\)`, blocs ```` ```latex ```` acceptés).
3. **Aperçu** : rendu de chaque expression, erreurs signalées individuellement.
4. Choisissez la disposition :
   - **Une équation par ligne** : chaque expression devient une équation centrée sur son propre paragraphe ;
   - **Conserver la disposition originale** : lignes, texte et lignes vides conservés, formules dans le texte insérées en ligne.
5. **Insérer dans Word** (ou Ctrl+Entrée). Si du texte est sélectionné, l'insertion se fait *après* la sélection (rien n'est effacé).

### Méthode B — Convert to Formula
1. Collez du LaTeX directement dans le document, sélectionnez-le.
2. **Convert to Formula** (ruban, clic droit ou Ctrl+Alt+Maj+M).

Comportement :
- une ou plusieurs expressions (une par ligne/paragraphe) → chacune convertie individuellement ;
- expression sur plusieurs lignes → une seule équation ;
- espaces avant/après → conservés, seule la formule est remplacée ;
- texte normal mélangé (« La capacité est `C = \frac{q}{U}` où… ») → seule la formule est remplacée, le texte reste ;
- expression invalide → texte original **intact**, le panneau s'ouvre avec le détail de l'erreur ;
- le texte hors sélection n'est jamais modifié ; la taille de police environnante est reprise.

> Pour le texte mixte sans délimiteurs, la détection est heuristique. Entourez la formule de `$…$` pour un résultat garanti.

---

## 3. Version hébergée (GitHub Pages)

Chaque `git push` sur `main` déclenche `.github/workflows/pages.yml` : tests, build avec l'URL de production, puis publication sur **https://silax04.github.io/mathpaste/**. Le manifeste de production est disponible à l'adresse https://silax04.github.io/mathpaste/manifest.xml.

Installer la version hébergée dans Word (une fois, Word fermé) :

```bash
curl -o "%USERPROFILE%\mathpaste-manifest.xml" https://silax04.github.io/mathpaste/manifest.xml
npx office-addin-dev-settings register "%USERPROFILE%\mathpaste-manifest.xml"
```

Aucun serveur local n'est alors nécessaire. Le code du complément est public, mais vos documents ne quittent jamais Word : la conversion s'exécute localement.

## 4. Build de production (autre hébergement)

```bash
# URL HTTPS où seront hébergés les fichiers de dist/ (terminée par "/")
set ADDIN_URL=https://mon-domaine.exemple/mathpaste/
npm run build
```

`dist/` contient le complément et un `manifest.xml` réécrit avec cette URL. Options d'hébergement gratuites : GitHub Pages, Azure Static Web Apps, ou un serveur local :

```bash
npm run build
npm run serve:prod
```

`serve:prod` sert `dist/` sur `https://localhost:3000` (certificat de dev) — gardez-le actif pendant l'utilisation de Word, ou ajoutez-le au démarrage de Windows. Aucun service cloud n'est nécessaire pour la conversion : tout s'exécute localement dans Word. Seule la bibliothèque `office.js` est chargée depuis le CDN Microsoft (obligatoire pour tout complément Office).

---

## 5. Architecture

```
LaTeX ──Temml──▶ MathML ──OmmlConverter──▶ OMML ──buildPackage──▶ Flat OPC ──insertOoxml──▶ Word
```

| Fichier | Rôle |
|---|---|
| `src/engine/latex.ts` | Normalisation (délimiteurs, `equation`/`align`, `\label`, `\tag`, commentaires) + parsing par **[Temml](https://temml.org)** (LaTeX → MathML, local, éprouvé) + messages d'erreur en français |
| `src/engine/omml.ts` | Convertisseur **MathML → OMML** produisant les structures natives de Word : `m:f`, `m:rad`, `m:sSub/sSup/sSubSup`, `m:d` (parenthèses extensibles, `\middle`), `m:func` (`\ln`, `\sin`, `\lim`…), `m:nary` (∑ ∫ ∏…), `m:acc`, `m:bar`, `m:groupChr`, `m:limLow/limUpp`, `m:m` (matrices, cases), `m:eqArr` (aligned), `m:borderBox` |
| `src/engine/charmaps.ts` | Tables : accents, opérateurs n-aires, espaces (`\,` `\;` `\quad`…), alphabets mathématiques (`\mathbf`, `\mathbb`…) |
| `src/engine/macros.ts` | Macros LaTeX personnalisées (`\R`, `\dd`, `\abs`…) |
| `src/engine/extract.ts` | Détection des expressions dans un texte (délimiteurs, multi-lignes, texte mixte) avec positions exactes |
| `src/engine/ooxml.ts` | Paquet Flat OPC pour `insertOoxml` (`m:oMathPara` pour les équations isolées, `m:oMath` en ligne) |
| `src/office/word.ts` | Intégration Office.js : insertion au curseur, conversion de la sélection (localisation par paragraphe + `search`, remplacement expression par expression, erreurs isolées) |
| `src/taskpane/` | Interface React (panneau) + enregistrement des commandes (runtime partagé) |
| `manifest.xml` | Onglet ruban, bouton, menu contextuel `ContextMenuText`, runtime partagé, raccourci (`shortcuts.json`) |

**Pourquoi ce choix ?** `mathml2omml` (npm) a été évalué : il ne produit ni objets délimiteurs pour `\left( \right)`, ni objets fonction (`ln` en italique). La feuille XSLT `MML2OMML.XSL` de Word n'est pas redistribuable. Le convertisseur maison reproduit exactement l'OMML que Word génère quand on saisit l'équation à la main (vérifié en comparant avec la sortie de l'éditeur d'équations de Word).

### Ajouter une commande LaTeX
1. **Macro** (le plus simple) : ajouter une entrée dans `src/engine/macros.ts`, ex. `"\\vect": "\\mathbf{#1}"`.
2. **Symbole / accent / opérateur** : compléter les tables de `src/engine/charmaps.ts`.
3. **Nouvel élément MathML** : ajouter un gestionnaire dans `ELEMENT_HANDLERS` (`src/engine/omml.ts`).
Puis ajouter un cas dans `tests/engine.test.ts`.

### Commandes Office : ce qui est supporté
- **Ruban** : onglet personnalisé `MathPaste` (manifeste XML, `PrimaryCommandSurface`).
- **Menu contextuel** : point d'extension `ContextMenu` / `OfficeMenu id="ContextMenuText"`, pris en charge par Word via le manifeste XML. Office.js ne permet pas d'ajouter dynamiquement des éléments au menu contextuel : l'entrée est statique et toujours visible lors d'un clic droit sur du texte.
- **Raccourci clavier** : `Ctrl+Alt+Maj+M` via `ExtendedOverrides` + runtime partagé. Disponible uniquement sur les versions récentes de Word ; si votre build ne le prend pas en charge, le ruban et le clic droit restent disponibles. En cas de conflit avec un autre raccourci, modifiez `src/shortcuts.json`.

---

## 6. Tests

```bash
npm test            # 69 tests unitaires du moteur (conversion + détection)
npm run typecheck   # vérification TypeScript
npm run validate    # validation du manifeste (envoie le manifeste au service de validation Microsoft)
npm run test:word   # validation dans le VRAI Word via COM
```

`test:word` insère l'OMML généré dans Word (`Range.InsertXML`, même importeur que `insertOoxml`) et vérifie :
1. les 3 équations de référence insérées en **une seule opération** → 3 objets équation natifs ;
2. indices, exposants, fraction verticale, racine, `ln` en fonction native, parenthèses extensibles ;
3. sauvegarde `.docx`, réouverture, édition individuelle (linéarisation / reconstruction) ;
4. scénario : 3 expressions dans le document, **seule la 2e** remplacée — les autres et le texte autour restent identiques ;
5. texte mixte : seule la formule est remplacée, l'équation est en ligne.

Les PDF produits par Word (`tests/word-out/*.pdf`) permettent un contrôle visuel.

Prévisualiser l'interface du panneau dans un navigateur (sans Word) : `npm run preview:ui` puis http://localhost:3000/taskpane.html.

---

## 7. Dépannage

| Problème | Solution |
|---|---|
| L'onglet MathPaste n'apparaît pas | `npm run stop`, puis `npm start`. Vérifier que `https://localhost:3000/taskpane.html` s'ouvre dans Edge sans alerte de certificat (sinon `npm run certs`). |
| « Document Word inaccessible » | Le panneau a été ouvert hors de Word ou le document est protégé / en lecture seule. |
| Une expression n'est pas convertie | Le détail s'affiche dans le panneau ; le texte original est conservé. |
| Cache du complément | Supprimer `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` puis relancer Word. |
