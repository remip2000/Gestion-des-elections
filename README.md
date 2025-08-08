# 🗳️ Élections Cogolin – Tableur de Suivi des Résultats

Ce classeur Excel (avec macros) est conçu pour **organiser, suivre et centraliser les résultats d'une élection municipale à Cogolin**, bureau par bureau. Il facilite le **dépouillement, la saisie, l’analyse des résultats** et le **suivi de la participation horaire**.

## 📁 Fichier
**Nom :** [`Elections Cogolin v1.12.xlsm`](./Elections%20Cogolin%20v1.12.xlsm)

**Version :** 1.12  
**Compatibilité :** Microsoft Excel avec macros activées (format `.xlsm`)

---

## 🔧 Fonctionnalités principales

- **Feuilles par bureau** (`Bureau 1`, `Bureau 2`, ..., `Bureau 12`) pour saisir :
  - Le suivi de la participation heure par heure
  - Le dépouillement des bulletins
  ![Aperçu du tableur](./Capture%20d’écran%20Bureau%20exemple.png)
- **Feuille "Synthèse"** :
  - Agrégation automatique des résultats de tous les bureaux
  - Calculs de participation, taux d’abstention, votes exprimés
  - Vérifications des totaux
    ![Aperçu du tableur](./Capture%20d’écran%20Synthèse.png)
- **Feuille "Participation horaire"** :
  - Suivi de la participation par tranche horaire dans chaque bureau
  - Calcul des pourcentages de participation cumulée
  - Comparaison possible avec le tour précdent
    ![Aperçu du tableur](./Capture%20d’écran%20Participation%20horaire.png)
- **Macros intégrées** (nécessite l’activation) :
  - Automatisations pour l'instanciation initiale du fichier
  - Générer les bureaux sur le modèle de `Bureau 0`
  - Répliquer la couleur des candidats dans tous les graphiques
  - Importer les données de participation du premier tour

---

## Protection des feuilles

L'utilisateur __ne peut et ne doit__ modifier que les cases 🟦 **bleu clair** 🟦. Toutes les cases ⬜ **gris clair** ⬜ sont calculées automatiquement. Si vous voulez adapter le fichier, le mot de passe de protection de chaque feuille est `toto`.

## 🧑‍💻 Utilisation

0. **Activer les macros** à l’ouverture du fichier Excel. *Utile mais pas nécessaire pour l'utilisation*.
1. Pour le bureau centralisateur :
   - Ouvrir l'onglet "Synthèse"
   - Renseigner les données de l'élection (noms des bureaux, nombre d'inscrits, date, élection, ville)
   - Renseigner le nom des candidats et leur couleur associée. Le bouton
3. Pour chaque bureau :
   - Ouvrir l’onglet correspondant (ex: *Bureau 5*)
   - Pendant la journée, renseigner les données de participation horaire à droite de la feuille
   - Au dépouillement, remplir les cases avec les résultats au fur et à mesure des enveloppes de centaines de bulletins dépouillés
4. Visualiser la progression de la participation dans *Participation horaire*.
5. **Suivre en temps réel la synthèse** dans l’onglet *Synthèse*. Toutes les données des bureaux sont centralisées dans *Synthèse*

### Dépouillement : deux méthodes possibles

On propose une [`feuille de dépouillement`](./Feuille_depouillement.xlsx) à imprimer en A3 portrait pour aider les scrutateurs à compter les bulletins.

1. **Décompte absolu** : On remplit le tableau du bas, table par table. Les scores des candidats sont croissants d'enveloppe (de centaine) en enveloppe et le calcul des scores par centaines se fait automatiquement dans le tableau du haut. Ce mode évite aux scrutateurs de faire des calculs à chaque centaine. La somme est cumulative de centaine en centaine.
[`Exemple en mode absolu`](./Feuille_depouillement%20exemple%20absolu.pdf)
2. **Décompte relatif** : On remplit le tableau du haut directement. On entre les scores relatif à la centaine dépouillée. La somme de chaque centaine doit être égale à 100.
[`Exemple en mode relatif`](./Feuille_depouillement%20exemple%20relatif.pdf)

---

## 📊 Exemple

L’onglet `Exemple` montre comment renseigner les données pour un bureau fictif avec des votes pour 10 candidats, ainsi que la progression horaire. Les feuilles de dépouillement d'exemple sont ausi fournies.

---

## ✅ Avantages

- Centralisation rapide et fiable des données
- Vérifications automatiques des totaux
- Aide à la transparence lors du dépouillement
- Lecture facile des résultats par tous les scrutateurs

---

## 📌 À savoir

- Ce classeur a été développé pour une élection municipale mais peut être adapté à d'autres scrutins.
- Les noms des candidats peuvent être modifiés directement dans les feuilles.
- Ce projet n’inclut pas de protection par mot de passe mais certaines feuilles peuvent être protégées par macro pour éviter les erreurs de saisie.

---

## 📄 Licence

Ce projet est sous licence **GNU General Public License v3.0** – voir le fichier [`LICENSE`](./LICENSE) pour plus d’informations.
