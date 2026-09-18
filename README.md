# Automatisation Excel AYA

Logiciel de bureau (Windows / macOS / Linux) qui met a jour la BDD des commissions
a partir du fichier Reference des promoteurs et du fichier d'affectation des KAM.

- **Fichier 1 - Reference** : source de verite. **STRICTEMENT LECTURE SEULE.**
- **Fichier 2 - BDD PROMOTERS MONTH** : le seul fichier corrige, dans une **copie** (jamais ecrase).
- **Fichier 3 - Affectation KAM** : sert uniquement a retrouver le KAM. **Jamais modifie.**

Traitement **deterministe** : aucune supposition, aucune donnee inventee, aucun
ecrasement arbitraire, aucun duplicate base uniquement sur le STORE.

---

## 1. Installation

1. Installer **Python 3.10+** : https://www.python.org/downloads/ (Windows : cocher *Add Python to PATH*).
2. Dans le dossier du logiciel :

```bash
pip install -r requirements.txt
python main.py
```

Windows : double-cliquer sur `Lancer_Windows.bat`. Linux : `sudo apt install python3-tk` si Tkinter manque.

Mode ligne de commande (automatisation / tests) :

```bash
python main.py --cli -r Reference.xlsx -t BDD.xlsx -k KAM.xlsx -o Resultat.xlsx
python main.py --cli ... --preview          # analyse seule, aucun fichier ecrit
python main.py --cli ... --auto-idaya       # numerote IDAYA sur les lignes creees
```

---

## 2. Structure attendue des fichiers

Les colonnes sont reconnues **par leur nom** (majuscules, accents et espaces ignores).
Si les en-tetes ne sont pas reconnus, le logiciel applique automatiquement les
**positions exactes** ci-dessous et le signale dans l'interface.

**Fichier 1 - Reference** - feuille `Feuil2`, en-tetes ligne 1
| Colonne | C | F | G | H | I | J | L |
|---|---|---|---|---|---|---|---|
| Contenu | Divisions | ID Promoter | Store Code | PROMOTER | STORE | CITY | Working Days |

**Fichier 2 - BDD a traiter** - feuille `BDD PROMOTERS MONTH`, en-tetes ligne 1
| A | B | C | D | E | F | G | H | I | J | K | L | M | N |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Annee | Mois | Division | Division1 | KAM | ID Promoter | Code Store | Promoter | Store | City | STATUT | DAY | IDAYA | Column1 |

**Fichier 3 - Affectation KAM** - une feuille `VD` et une feuille `DA` (colonnes Store et KAM).

---

## 3. Regles appliquees (aucune autre)

**Ordre de traitement, pour chaque ligne du fichier 2 :**

| Etape | Regle | Comportement |
|---|---|---|
| 1 | **Store (I) vide** | Toute la ligne est coloree en **ROUGE**. Aucune recherche Store / Promoter / ID / Code Store / KAM, aucune creation, aucune modification, aucun remplissage automatique. |
| 2 | **Store absent de la Reference** | Signale "Store non trouve". Aucune donnee inventee, aucune recherche approchante, ligne laissee intacte. |
| 3 | **Divisions de la Reference** | `VD`, `DA` et `VD+DA` sont traitees. `VD+DA` = **deux** affectations (VD et DA) traitees separement. Toute autre valeur (ex. `RAC`) va en **A VERIFIER** et n'est jamais transformee en VD ou DA. |
| 4 | **Meme Store + meme Division avec 2 promoteurs** | **A VERIFIER**. Aucun choix automatique, aucune ligne ecrasee, donnees laissees intactes. |
| 5 | **Cle de traitement** | `Code Store + Division`. `C003470765VD` et `C003470765DA` sont **deux cles differentes**. Le Store seul ne suffit jamais. |
| 6 | **Cle presente dans le fichier 2** | La ligne existante est **mise a jour** (jamais de 2e ligne). |
| 7 | **Cle absente** | Une nouvelle ligne est **creee**. |
| 8 | **KAM** | Recherche par **Store + Division** (feuille VD ou feuille DA). KAM introuvable = colonne ecrite avec le marqueur **`#`** (ligne existante ou creee) + signale. Jamais de KAM copie d'une division a l'autre. |
| 9 | **Column1** | Toujours recalculee : `Code Store + Division1`, uniquement quand le Code Store et la Division viennent de la Reference. Une ancienne valeur fausse est corrigee. |
| 10 | **IDAYA** | N'est jamais une cle de duplicate. Les IDAYA existants ne sont jamais modifies. L'option de numerotation ne touche que les lignes **creees**. |

**Colonnes ecrites depuis la Reference** (regles 4 et 5 du cahier des charges) :
`Division` (C), `Division1` (D), `KAM` (E), `ID Promoter` (F), `Code Store` (G),
`Promoter` (H), `Store` (I), `City` (J), `DAY` (L), `Column1` (N).

**Colonnes jamais modifiees sur une ligne existante** : `Annee` (A), `Mois` (B),
`STATUT` (K), `IDAYA` (M).

### Comment une ligne existante est reconnue
Dans cet ordre de confiance :
1. **`Store (I)` + `Division1 (D)`** — le nom du magasin est ce qui l'identifie vraiment
   (c'est ainsi qu'on le cherche dans la Reference). Indispensable quand deux magasins
   partagent le meme `Store Code` dans la Reference : chacun retrouve **sa** ligne ;
2. `Code Store (G) + Division1 (D)` — la cle metier reconstruite ;
3. la valeur ecrite dans `Column1 (N)` — qui peut etre ancienne ou fausse.

C'est ce qui permet de corriger une `Column1` erronee (regle 12) sans creer de doublon.
Si deux magasins differents finissent avec la meme cle, c'est signale en `A VERIFIER`.

### Cellules calculees
Une cellule qui contient une **formule** n'est jamais ecrasee : elle se recalcule
toute seule. Sur une ligne creee, les formules de la ligne du meme `Code Store`
sont recopiees en ajustant leurs references (comme un copier-coller Excel).

---

## 4. Statistiques (definitions exactes)

| Statistique | Definition |
|---|---|
| Lignes analysees | Nombre de lignes de donnees du fichier 2. |
| Lignes modifiees | Lignes existantes dont au moins une cellule a change. |
| Lignes deja conformes | Cle trouvee, valeurs deja identiques a la Reference. |
| Lignes ajoutees | Nouvelles lignes creees (cle absente du fichier 2). |
| Lignes Store vide | Lignes colorees en rouge. |
| Stores non trouves | Nombre de **magasins** du fichier 2 absents de la Reference. |
| KAM trouves / non trouves | Par affectation traitee (Store + Division). |
| **Duplicates evites** | **Uniquement** les creations annulees parce que la combinaison `Code Store + Division` avait deja ete traitee. Ce n'est jamais le nombre de lignes traitees. |
| A verifier | Cas ambigus que le logiciel refuse de trancher. |

---

## 5. Rapport : pourquoi chaque ligne a ete traitee ainsi

Le rapport (`..._RAPPORT.txt` et `..._RAPPORT.xlsx`) contient **une decision par ligne**
du fichier 2, avec la raison et le detail des cellules changees :

| Decision | Signification |
|---|---|
| `ROUGE` | Store vide -> coloration uniquement. |
| `MODIFIEE` | Cle presente -> ligne mise a jour (liste des cellules et anciennes valeurs). |
| `CONFORME` | Cle presente, deja identique a la Reference. |
| `CREEE` | Cle absente -> nouvelle ligne (source Reference indiquee). |
| `STORE NON TROUVE` | Store absent de la Reference. |
| `A VERIFIER` | RAC, 2 promoteurs pour la meme cle, cle en double dans le fichier 2... |
| `IGNOREE` | Ligne du fichier 2 sans affectation correspondante dans la Reference. |

---

## 6. Les 3 points ou les regles ne disent rien (choix documentes)

1. **Colonne `Division` (C) quand la Reference vaut `VD+DA`** : la valeur brute
   `VD+DA` est ecrite en C et la division reelle (`VD` ou `DA`) en D (`Division1`).
   Pour une Reference `VD` ou `DA`, C et D contiennent la meme valeur, comme exige.
2. **Colonnes d'une ligne CREEE que la Reference ne fournit pas** (`Annee`, `Mois`,
   `STATUT`, colonnes supplementaires) : la nouvelle ligne est une **copie** de la
   ligne du meme `Code Store` deja presente dans le fichier 2 (valeurs, formules et
   mise en forme), puis les colonnes de la Reference sont ecrites par-dessus.
   S'il n'y a aucune ligne a copier, elles restent vides et la ligne passe en
   `A VERIFIER`. Le rapport indique toujours la ligne source.
3. **`KAM` introuvable** : rien n'est invente. La colonne recoit le marqueur
   **`#`** (decision explicite de l'utilisateur, meme convention deja presente
   dans son fichier), que la ligne soit existante ou creee. C'est toujours
   signale dans le rapport.

---

## 7. Ce qui est conserve dans le fichier resultat

Conserve : toutes les feuilles (y compris masquees), valeurs, **formules**, couleurs,
polices, bordures, largeurs, volets figes, tableaux Excel (agrandis lors d'un ajout de
ligne), filtres automatiques et mises en forme conditionnelles.

Non conserve (limite d'openpyxl, signale au chargement) : connexions de donnees
externes / Power Query, tableaux croises dynamiques, graphiques, images, segments,
parametres d'impression, vues de feuille nommees. Les resultats des formules sont
recalcules par Excel a l'ouverture.

---

## 8. Architecture

```
main.py            point d'entree (interface ou ligne de commande)
gui.py             interface CustomTkinter (repli Tkinter)
pipeline.py        orchestration : lecture -> analyse -> ecriture
excel_reader.py    lecture, detection des feuilles/colonnes (nom puis position)
matcher.py         normalisation des textes, cles, divisions
kam_service.py     recherche du KAM par Store + Division
processor.py       REGLES METIER (produit un plan, ne touche pas a Excel)
excel_writer.py    application du plan, ecriture d'un nouveau fichier
logger.py          decision + raison pour chaque ligne, rapports .txt et .xlsx
tests/             fixtures Excel + 28 tests automatiques
```

---

## 9. Tests

```bash
python -m unittest discover -s tests -v
```

28 tests, dont les **10 cas obligatoires** :
Store vide - Store non trouve - Store+VD existant - Store+DA existant -
VD existant mais DA absent - VD+DA - meme Store + meme Division avec 2 promoteurs -
KAM different entre VD et DA - Column1 incorrect - deuxieme execution sans duplicate.
Plus : RAC non traite, KAM vide si introuvable, colonnes jamais touchees, statistiques
exactes, rapport explicatif, sources non modifiees, mise en forme conservee,
detection par position quand les en-tetes sont illisibles, reconnaissance de la
ligne par le nom du Store, formules jamais ecrasees, copie complete de la ligne soeur.

---

## 10. Problemes frequents

| Message | Solution |
|---|---|
| `No module named openpyxl` | `pip install -r requirements.txt` |
| `No module named tkinter` | Linux : `sudo apt install python3-tk` |
| `aucune feuille ne contient les colonnes obligatoires` | Choisir la feuille dans la liste deroulante, ou verifier les en-tetes. |
| `le fichier est peut-etre ouvert dans Excel` | Fermer le fichier resultat puis relancer. |
