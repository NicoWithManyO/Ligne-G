# Conventions de code


## 1. Langue & Nommage
- **Langue du code** : uniquement **anglais** (fonctions, variables, classes, fichiers…).
- **Commentaires** : uniquement **en français**, bien rédigés, clairs et utiles.
- **Conventions de nommage** :
  - `camelCase` pour fonctions, variables, fichiers, objets
  - MAJUSCULES + underscores pour les constantes (`MAX_VALUE`, `DEFAULT_PATH`)
  - Pas d’accents, pas d’abréviations obscures

---
## 2. Commentaires
- Toujours commenter **le pourquoi**, jamais le comment.
- En-tête de chaque fonction :
  - But de la fonction
  - Paramètres attendus
  - Valeur(s) retournée(s)
  - Préconditions éventuelles
> Exemple :
> ```python
> # Calcule la moyenne d'une liste de valeurs numériques
> # @param values: liste de floats ou d'entiers
> # @return: moyenne (float)
> ```

---
## 3. Débogage
- Logs **techniques, concis, sans phrases rédigées**
- Toujours tracer les étapes critiques du code, sans surcharge inutile
- Doit pouvoir être activé/désactivé facilement (flag global, condition, etc.)
- Format recommandé :
  ```text
  [nomFonction] variable = valeur
  ```

---
## Bonnes pratiques générales
Le code doit rester **clair, modulaire et pérenne** : une fonction = une responsabilité, aucune valeur codée en dur sans justification, zéro dépendance implicite, et toute logique doit être compréhensible sans contexte. La **lisibilité prime toujours sur l’optimisation micro**.

- Toute logique "métier" (ex: vérification de conformité) doit être isolée et testable indépendamment
- Toujours laisser **deux lignes vides** entre chaque fonction dans les modules VBA
