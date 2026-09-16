# Automatisation Excel AYA

Logiciel de bureau (Windows / macOS / Linux) qui met a jour automatiquement la BDD
des commissions a partir du fichier Reference des promoteurs et du fichier
d'affectation des KAM.

- **Fichier 1 - Reference** : source de verite (promoteurs, divisions, jours). **Jamais modifie.**
- **Fichier 2 - BDD a traiter** : le seul fichier corrige. **Jamais ecrase** : un nouveau fichier est cree.
- **Fichier 3 - Affectation KAM** : sert uniquement a retrouver le KAM. **Jamais modifie.**

---

## 1. Installation (etape par etape)

1. Installer **Python 3.10 ou plus** : https://www.python.org/downloads/
   Sous Windows, cocher **"Add Python to PATH"** pendant l'installation.
2. Telecharger ce dossier sur l'ordinateur.
3. Ouvrir un terminal (Windows : touche Windows -> taper `cmd`) dans le dossier du logiciel, puis :

```bash
pip install -r requirements.txt
```

4. Demarrer le logiciel :

```bash
python main.py
```

Sous Windows, on peut aussi **double-cliquer sur `Lancer_Windows.bat`**.

> Linux uniquement : si le message "No module named tkinter" apparait, faire
> `sudo apt install python3-tk`.

---

## 2. Utilisation (interface)

```
1. Fichiers      -> choisir les 3 fichiers Excel (la feuille est detectee automatiquement,
                    mais on peut la forcer avec la liste deroulante a droite)
2. Options       -> a laisser decochees pour appliquer les regles metier strictes
3. Lancer l'analyse (apercu)  -> AUCUN fichier n'est ecrit, on voit seulement ce qui va se passer
4. Statistiques + Rapport     -> verifier surtout la ligne "A verifier"
5. Confirmer et generer le fichier -> choisir le nom du fichier resultat
6. Ouvrir le fichier resultat / Ouvrir le rapport
```

Le bouton *Confirmer et generer* n'ecrit jamais par-dessus un des 3 fichiers sources :
le logiciel refuse et affiche une erreur.

### Mode ligne de commande (optionnel, pour automatiser)

```bash
python main.py --cli ^
  -r "Liste_promoteurs.xlsx" ^
  -t "BDD_COMMISSIONS_AOUT.xlsx" ^
  -k "Affectation_des_KAMs.xlsx" ^
  -o "BDD_COMMISSIONS_AOUT_TRAITE.xlsx"
```

`--preview` = analyse seule, sans ecriture. `python main.py --help` liste toutes les options.

---

## 3. Colonnes obligatoires

Les colonnes sont reconnues **par leur nom** (pas par leur position), sans tenir compte
des majuscules, des accents ni des espaces en trop. Si une colonne obligatoire manque,
le traitement s'arrete avec un message clair.

| Fichier | Colonnes obligatoires |
|---|---|
| 1 - Reference | ID Promoter, Store Code, PROMOTER, STORE, Divisions, Working Days |
| 2 - BDD a traiter | ID Promoter, Code Store, Promoter, Store, Division1, KAM, DAY, Column1 |
| 3 - Affectation KAM | Store, KAM |

Colonnes facultatives utilisees si elles existent : `CITY`, `Annee`, `Mois`, `Division`,
`STATUT`, `IDAYA`, `Code store`.

---

## 4. Regles appliquees

| # | Regle | Comportement du logiciel |
|---|---|---|
| 1 | **Store vide** | Toute la ligne est coloree en **rouge**. Aucune autre modification, aucune recherche, aucun duplicate. |
| 2 | Recherche du Store | Comparaison robuste : espaces multiples, majuscules/minuscules, accents, tirets. Une cellule vide n'est jamais un match. |
| 3 | Donnees reprises de la Reference | `ID Promoter`, `Store Code` -> `Code Store`, `PROMOTER` -> `Promoter`, `Divisions` -> `Division`, division unitaire -> `Division1`, `Working Days` -> `DAY`, cle -> `Column1`. |
| 4 | `Divisions = VD+DA` | Genere **deux** affectations : une VD et une DA. |
| 5 | Meme Store sur 2 lignes Reference | Chaque ligne garde **ses propres** informations : jamais de melange VD/DA. |
| 6 | Cle `Code Store + Division1` | Si la cle existe dans `Column1` -> la ligne est **modifiee**. Sinon -> une ligne est **creee**. |
| 7 | Duplicates | Une cle n'est jamais creee deux fois. Le compteur "Duplicates evites" indique le nombre de creations evitees. |
| 8 | KAM | Recherche dans le fichier 3 a partir du Store **et** de la division (voir point 5 ci-dessous). Jamais de recherche si Store est vide. |
| 9 | Fichier resultat | Toujours un **nouveau** fichier. Les 3 sources restent identiques (verifie par test automatique). |
| 10 | Cas ambigus | Le logiciel **ne choisit jamais au hasard** : il inscrit la ligne dans le rapport "A VERIFIER". |

### Colonnes jamais modifiees sur une ligne existante
`Annee`, `Mois`, `City`, `STATUT`, `IDAYA`, `Store` restent tels quels
(`City` peut etre mise a jour via une option).

---

## 5. Points detectes dans vos fichiers reels (a valider)

L'analyse des 3 fichiers fournis a revele 4 situations que le cahier des charges ne
tranchait pas. Aucune regle n'a ete inventee : voici ce qui a ete fait.

1. **Le KAM depend du Store ET de la division.**
   `Affectation_des_KAMs_VDDA.xlsx` contient 2 feuilles (`VD` et `DA`). Sur les
   122 magasins presents dans les deux feuilles, **122 ont un KAM different en VD et en DA**
   (ex. Aswak Assalam Hay Riad : VD = HAZZAZ MEHDI, DA = ABOURACHID HOUSSINE).
   -> Le logiciel cherche donc le KAM dans la feuille de la division traitee.
   Si le fichier 3 n'a qu'une seule feuille sans nom de division, la recherche se fait
   uniquement par Store (comportement decrit dans le cahier des charges).
   Si le magasin n'existe que dans l'autre division, **le KAM n'est pas recopie** :
   la valeur existante est conservee et la ligne est signalee.

2. **Division `RAC`.** La Reference contient 6 lignes avec `Divisions = RAC`
   (Electroplanet Fes, Fes Saiss, Meknes, Nador, Oujda, Tanger), non prevue par les regles.
   -> Par defaut ces lignes ne sont **pas** traitees et sont listees dans "A VERIFIER".
   L'option *"Traiter les divisions non prevues"* permet de les traiter comme VD/DA.

3. **Deux promoteurs pour le meme Store + meme division.**
   Ex. Electroplanet Derb Sultan : 2 lignes `DA` avec 2 promoteurs differents ;
   Marjane Hay Riad : 2 lignes `VD` et 2 lignes `DA`. La cle `Code Store + Division1`
   ne peut contenir qu'une seule ligne.
   -> Le logiciel **ne modifie rien** et inscrit le detail dans "A VERIFIER".

4. **Colonne `Column1` parfois incoherente.** Dans le fichier fourni, 3 lignes ont un
   `Column1` qui ne correspond pas a `Code Store + Division1` (ex. ligne 68).
   -> La recherche se fait sur `Column1` **et** sur `Code Store + Division1`, puis en
   dernier recours sur `Store + Division1`. La cle est ensuite reecrite correctement.
   Cela evite de creer un doublon a cause d'une cle fausse.

---

## 6. Options (toutes decochees par defaut)

| Option | Effet |
|---|---|
| Mettre a jour la colonne City | Recopie `CITY` de la Reference dans `City`. |
| Traiter les divisions non prevues (RAC...) | Traite toute valeur de `Divisions` comme une division normale. |
| Numeroter automatiquement IDAYA | Les lignes creees recoivent le numero suivant. Sinon `IDAYA` reste vide et est signale. |
| Ajouter les magasins de la Reference absents du fichier 2 | Ajoute les affectations des magasins qui n'apparaissent pas du tout dans le fichier 2 (93 magasins dans vos fichiers). **Hors regles du cahier des charges.** |
| Surligner en orange les magasins absents de la Reference | Aide visuelle, aucune donnee modifiee. |

---

## 7. Ce qui est conserve dans le fichier resultat

Conserve : toutes les feuilles (y compris masquees), les valeurs, les **formules**,
les couleurs, les polices, les bordures, les largeurs de colonnes, les volets figes,
les tableaux Excel (agrandis automatiquement quand des lignes sont ajoutees),
les filtres automatiques et les mises en forme conditionnelles.

Non conserve (limite de la bibliotheque openpyxl, signale par le logiciel au chargement) :
connexions de donnees externes / Power Query, tableaux croises dynamiques, graphiques,
images, segments, parametres d'impression, vues de feuille nommees.
Les valeurs **calculees** des formules sont recalculees par Excel a l'ouverture.

> Si le fichier 2 contient une de ces fonctionnalites, faites une copie de securite
> avant de remplacer votre fichier de travail par le fichier resultat.

---

## 8. Rapport de traitement

A cote du fichier resultat, deux rapports sont crees :
`..._RAPPORT.txt` (lecture rapide) et `..._RAPPORT.xlsx` (filtrable dans Excel).

Categories : `MODIFICATION`, `CREATION`, `LIGNE ROUGE`, `STORE INTROUVABLE`,
`KAM INTROUVABLE`, `DUPLICATE EVITE`, `A VERIFIER`, `ERREUR`.

**Regardez toujours la categorie `A VERIFIER` en premier** : ce sont les situations que
le logiciel a refuse de trancher tout seul.

---

## 9. Architecture du code

```
main.py            point d'entree (interface ou ligne de commande)
gui.py             interface graphique (CustomTkinter, repli Tkinter)
pipeline.py        chef d'orchestre : lecture -> analyse -> ecriture
excel_reader.py    lecture des 3 fichiers, detection des feuilles et des colonnes
matcher.py         nettoyage et comparaison des textes, cles, divisions
kam_service.py     recherche du KAM (par Store + division)
processor.py       REGLES METIER (ne touche jamais a Excel : produit un plan)
excel_writer.py    application du plan et enregistrement du nouveau fichier
logger.py          rapport de traitement (.txt et .xlsx)
tests/             fichiers Excel de test + 22 tests automatiques
```

Le decoupage garantit qu'une regle metier se corrige dans `processor.py` uniquement,
sans toucher a l'interface ni a l'ecriture Excel.

---

## 10. Tests

```bash
python -m unittest discover -s tests -v
```

22 tests couvrent : detection des colonnes, ligne rouge, VD+DA, duplicates, conflits,
division inconnue, KAM par division, non-modification des sources, conservation de la
mise en forme, idempotence (relancer le traitement ne change plus rien) et les options.

---

## 11. Problemes frequents

| Message | Solution |
|---|---|
| `No module named openpyxl` | `pip install -r requirements.txt` |
| `No module named tkinter` | Windows/macOS : reinstaller Python ; Linux : `sudo apt install python3-tk` |
| `aucune feuille ne contient les colonnes obligatoires` | Verifier l'orthographe des en-tetes, ou choisir la feuille manuellement dans la liste deroulante. |
| `le fichier est peut-etre ouvert dans Excel` | Fermer le fichier resultat dans Excel puis relancer la generation. |
| Le traitement semble long | Normal : un classeur de plusieurs Mo met quelques secondes a s'ouvrir. La barre de progression avance ensuite. |
