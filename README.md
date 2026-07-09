# Étude / correction INSi depuis StudioVision (v9)

## Robustesse (v9) — après le plantage « objet '' » de Microsoft Access
Le plantage rencontré (erreur Jet « n'a pas pu trouver l'objet '' » sur l'ouverture du
sous-formulaire, après des milliers d'opérations) est une **saturation transitoire d'Access**.
La v9 la traite :

- **Reprise fiable** : seuls les résultats **définitifs** (00 / 01 / 02 / pas_sexe_*) sont
  conservés. Les **erreurs** (`sous_formulaire_absent`, `exception`, …) sont **automatiquement
  retraitées** au prochain lancement — **plus besoin de supprimer des lignes** dans le JSON.
- **Réessais** : l'ouverture du sous-formulaire est retentée (3 fois, avec récupération) ; un
  patient en échec transitoire est **rejoué une fois** après une pause. Ton patient qui
  marchait « à la main » passera donc tout seul.
- **Anti-saturation** : la connexion à la base est **mise en cache** (fini les milliers
  d'objets `CurrentDb`), avec **libération mémoire + pause tous les 50 patients**, et un
  rafraîchissement du lien COM tous les 200 patients.
- **Anti-cascade** : si **StudioVision est fermé** (ou ne répond plus), l'étude tente une
  reconnexion puis **s'arrête proprement** (au lieu d'enchaîner des centaines d'erreurs comme
  la dernière fois). L'avancement est sauvegardé ; relancer reprend tout seul.

Concrètement, après ton incident : **relance simplement `5 - Etude INSi (complete)`**. Les
patients déjà OK sont gardés, les lignes en erreur (dont celles créées après la fermeture de
StudioVision) sont refaites, et l'étude continue.

---

# Étude / correction INSi depuis StudioVision (v9)

Teste l'appel INSi sur les patients ayant consulté une année, **corrige les fiches dont le
champ sexe est vide**, et compte les résultats par type (00 trouvé / 01 non trouvé / 02
plusieurs / cas sexe / autres). Le programme s'arrête à la lecture du code réponse (il
n'ouvre pas WebDMP).

## Nouveauté v8 : le sexe absent est corrigé (plus jamais d'appel sans sexe)
Quand le champ **SEXE** d'une fiche est vide, l'appel INSi échoue (`insi_24 – Le format du
champ sexe est incorrect`). La v8 ne fait **plus jamais** d'appel avec un sexe vide : elle
**remplit d'abord le sexe**, puis appelle.

Méthode (quand le sexe est absent) :
1. Si le **n°SS** (`SS`) est présent **et** que la fiche n'est pas un enfant (`Titre` ≠
   « Enfant ») → le sexe est déduit du 1er chiffre du NIR (**1 = M, 2 = F**) et essayé en
   premier ; l'autre sexe est essayé ensuite en cas d'échec.
2. Sinon (pas de n°SS, ou enfant dont le n°SS est celui d'un parent) → essais **F puis M**.
3. Jusqu'à **2 appels INSi**. Dès qu'un appel **trouve** le patient (code 00, ou 02), le
   sexe qui a fonctionné est **conservé** dans la fiche (`patients.SEXE`, corrigé
   durablement). Si aucun essai n'aboutit, le champ sexe est **remis à vide**.

Classifications quand la résolution échoue et que le sexe est remis vide :
`pas_sexe_pas_numSS` (aucun n°SS), `pas_sexe_enfant` (enfant), `sexe_non_resolu` (n°SS
présent mais aucun sexe ne donne 00/02).

⚠️ **Le champ SEXE est écrit dans la base `patients`** (valeurs 1=M / 2=F, réglables via
`SEXE_M`/`SEXE_F`). Vérifiez ces valeurs sur votre base — d'où l'étape 6 ci-dessous, à faire
en premier sur un petit lot.

## Reprise de l'étude
`5 - Etude INSi (complete)` **reprend automatiquement** là où elle s'était arrêtée : les
patients déjà présents dans `etude_insi_resultats.json` sont conservés et ignorés. Vous
pouvez donc relancer sans tout recommencer (Ctrl+C sauvegarde l'avancement à tout moment).

## Ordre conseillé (fichiers .bat)
StudioVision **ouvert, une fiche patient affichée**.

0. `0 - Inspecter les formulaires` — diagnostic (contrôles).
1. `1 - Lister patients 2026` — vérifie la lecture de PUBLIC.mdb (aucun appel).
2. `2 - Tester ouverture de N fiches` — ouverture par RecordSource, **sans INSi**.
3. `3 - Tester UN patient` — un code : ouverture + (résolution du sexe si besoin) + INSi.
4. `6 - Corriger le sexe (patients sans reponse)` — **applique la résolution du sexe** aux
   patients déjà notés « sans_reponse » dans un résultat existant, et **met le JSON à jour**
   (sauvegarde `.bak`). Idéal pour **vérifier la méthode sur un petit lot** (`--limit 20`)
   avant de tout relancer.
5. `4 - Etude (test 10 patients)`, puis `5 - Etude (complète)`.

## ⚠️ Point réglementaire important (INSi en masse)
D'après l'ANS / le Ségur : **il n'est pas autorisé par la CNIL d'appeler la récupération du
téléservice INSi en masse (ou en lot) pour l'ensemble d'une base d'identités** ; cela doit se
faire au fil des venues des patients. Les appels sont tracés (nombre total par période,
nombre moyen par identité…). Appeler INSi pour tous les patients 2026, même étalé sur
plusieurs heures, relève de cet usage « en masse ». Deux options plus sûres pour corriger les
sexes : **écrire le sexe déduit du n°SS sans appeler INSi**, ou **ne lancer les appels qu'au
fil des venues**. L'outil est fourni ; l'usage relève de votre responsabilité.


## 7 — Analyse statistique (hors ligne)

`7 - Analyse statistique (depuis etude_insi_resultats.json).bat` lit le fichier de résultats
et produit un **rapport complet**. **Aucune connexion à StudioVision n'est nécessaire** : simple
lecture du JSON, exécutable à tout moment, même si l'étude est inachevée. Le lanceur demande
(optionnellement) le nombre total de patients de l'année, pour afficher le **taux de couverture**.

Contenu du rapport :
1. **Résultat principal** — appels aboutis, répartition 00 / 01 / 02, **taux de réussite INSi**.
2. **Répartition complète** de toutes les classifications (avec histogrammes).
3. **Fiches sans sexe** — combien corrigées, par méthode (n°SS / enfant / essais F-M), avec le
   taux de résolution de chaque méthode.
4. **Croisements** sur les appels aboutis : par **sexe**, **tranche d'âge**, **présence du n°SS**,
   **titre « Enfant »** — utile pour comprendre *qui* produit les erreurs 01.
5. **INS** obtenus.
6. **Appels au téléservice** émis (total, et patients ayant nécessité 2 appels).
7. **Qualité d'exécution** : incidents à reprendre, incohérences nom/fiche.

Fichiers produits :
- `etude_insi_statistiques.json` — toutes les mesures, exploitables ailleurs ;
- `etude_insi_rapport.md` — le rapport, partageable tel quel ;
- `etude_insi_01_non_trouve.csv` — **la liste des patients dont les traits divergent** du
  référentiel (à corriger dans StudioVision) ;
- `etude_insi_a_reprendre.csv` — les incidents techniques ;
- `etude_insi_sexe_corrige.csv` — les fiches dont le sexe a été corrigé.

Les CSV sont en `;` + UTF-8 BOM : **double-clic → Excel** s'ouvre correctement (accents inclus).
Aucune dépendance à installer (bibliothèque standard uniquement).

En ligne de commande : `python analyse_statistique.py [--fichier x.json] [--total 3500] [--sans-csv]`

---

## Sorties
- `etude_insi_resultats.json` — un enregistrement par patient (réponse INSi + `nom_fiche`,
  `sexe_fiche`, `num_ss_present`, `titre`, `sexe_source`, `sexe_corrige`, `nb_appels_insi`).
- `etude_insi_statistiques.json` — proportions par type.
- `etude_ouverture_resultats.json` — résultats du mode « test ouverture ».

Le **numéro INS complet** et le **n°SS** ne sont **pas** enregistrés (INS seulement marqué
présent + version masquée).

## Réglages (haut de `etude_insi.py`)
`TABLE_PATIENTS`, `CHAMP_NUMSS="SS"`, `CHAMP_TITRE="Titre"`, `CHAMP_SEXE_TABLE="SEXE"`,
`VAL_TITRE_ENFANT="Enfant"`, `SEXE_M=1`, `SEXE_F=2`, `RS_GABARIT`, `INTER_APPEL`, `NAV_ATTENTE`.
