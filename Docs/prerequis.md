# Prérequis

Ce document décrit les étapes nécessaires pour que le système fonctionne correctement dans Excel, notamment en ce qui concerne l'accès au code VBA et l'import des modules.

---

## ? 1. Activer la référence "Microsoft VBA Extensibility 5.3"

Cette référence est indispensable pour que le code puisse modifier dynamiquement les modules VBA (injection, import automatique…).

### Étapes :
1. Ouvre l'éditeur VBA (`Alt` + `F11`)
2. Menu **Outils > Références…**
3. Coche la case **Microsoft Visual Basic for Applications Extensibility 5.3**
4. Clique sur **OK**

---

## ? 2. Autoriser l'accès au modèle d'objet VBA

Cette option permet à Excel d'accéder à son propre modèle d'objet (indispensable pour manipuler le code depuis une macro).

### Étapes :
1. Menu **Fichier > Options**
2. Clique sur **Centre de gestion de la confidentialité**
3. Dans le menu à gauche, clique sur **Centre de gestion de la confidentialité**
4. Clique sur le bouton **Paramètres du Centre de gestion de la confidentialité**
5. Va dans la section **Paramètres des macros**
6. Coche la case :
   > ?? **Accès approuvé au modèle d'objet du projet VBA**
7. Clique sur **OK**, puis à nouveau **OK** pour sortir

---

## ? 3. Charger manuellement le module `modLoader.bas`

Le module `modLoader.bas` est le point d'entrée du système. Il contient les fonctions nécessaires pour charger automatiquement tous les autres modules dans le projet VBA.

### Étapes :
1. Dans l'éditeur VBA (`Alt` + `F11`), sélectionne **Fichier > Importer un fichier…**
2. Navigue jusqu'au dossier `vbaCodebase` et sélectionne `modLoader.bas`
3. Vérifie qu'il apparaît bien dans la liste des modules à gauche
4. Sauvegarde ton fichier Excel

Ensuite, il sera possible d'utiliser les fonctions de chargement automatique (`loadModulesFromFolder`, etc.)

---

## ? 4. Injecter le démarrage automatique (`Workbook_Open`)

Après avoir importé le module `modLoader.bas`, il est nécessaire d'exécuter manuellement la procédure suivante :

```vba
Call injectWorkbookOpenInit
```

Cette macro insère automatiquement l'appel à `initWorkbook` dans l'événement `Workbook_Open`, ce qui permet d'initialiser correctement le projet à chaque ouverture du fichier.

> ?? Cette étape doit être réalisée une fois après l'import initial de `modLoader.bas`. Elle peut être relancée en cas de réinitialisation du projet.

> ?? Remarque : Le module `ThisWorkbook` n'est pas supprimé lors de l'exécution de `loadModulesFromFolder`,  
> car il s'agit d'un objet système intégré.  
> Pour injecter automatiquement le démarrage (`Workbook_Open`), il faut exécuter manuellement `injectWorkbookOpenInit`.
