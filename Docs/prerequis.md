# Pr�requis

Ce document d�crit les �tapes n�cessaires pour que le syst�me fonctionne correctement dans Excel, notamment en ce qui concerne l'acc�s au code VBA et l'import des modules.

---

## ? 1. Activer la r�f�rence "Microsoft VBA Extensibility 5.3"

Cette r�f�rence est indispensable pour que le code puisse modifier dynamiquement les modules VBA (injection, import automatique�).

### �tapes :
1. Ouvre l'�diteur VBA (`Alt` + `F11`)
2. Menu **Outils > R�f�rences�**
3. Coche la case **Microsoft Visual Basic for Applications Extensibility 5.3**
4. Clique sur **OK**

---

## ? 2. Autoriser l'acc�s au mod�le d'objet VBA

Cette option permet � Excel d'acc�der � son propre mod�le d'objet (indispensable pour manipuler le code depuis une macro).

### �tapes :
1. Menu **Fichier > Options**
2. Clique sur **Centre de gestion de la confidentialit�**
3. Dans le menu � gauche, clique sur **Centre de gestion de la confidentialit�**
4. Clique sur le bouton **Param�tres du Centre de gestion de la confidentialit�**
5. Va dans la section **Param�tres des macros**
6. Coche la case :
   > ?? **Acc�s approuv� au mod�le d'objet du projet VBA**
7. Clique sur **OK**, puis � nouveau **OK** pour sortir

---

## ? 3. Charger manuellement le module `modLoader.bas`

Le module `modLoader.bas` est le point d'entr�e du syst�me. Il contient les fonctions n�cessaires pour charger automatiquement tous les autres modules dans le projet VBA.

### �tapes :
1. Dans l'�diteur VBA (`Alt` + `F11`), s�lectionne **Fichier > Importer un fichier�**
2. Navigue jusqu'au dossier `vbaCodebase` et s�lectionne `modLoader.bas`
3. V�rifie qu'il appara�t bien dans la liste des modules � gauche
4. Sauvegarde ton fichier Excel

Ensuite, il sera possible d'utiliser les fonctions de chargement automatique (`loadModulesFromFolder`, etc.)

---

## ? 4. Injecter le d�marrage automatique (`Workbook_Open`)

Apr�s avoir import� le module `modLoader.bas`, il est n�cessaire d'ex�cuter manuellement la proc�dure suivante :

```vba
Call injectWorkbookOpenInit
```

Cette macro ins�re automatiquement l'appel � `initWorkbook` dans l'�v�nement `Workbook_Open`, ce qui permet d'initialiser correctement le projet � chaque ouverture du fichier.

> ?? Cette �tape doit �tre r�alis�e une fois apr�s l'import initial de `modLoader.bas`. Elle peut �tre relanc�e en cas de r�initialisation du projet.

> ?? Remarque : Le module `ThisWorkbook` n'est pas supprim� lors de l'ex�cution de `loadModulesFromFolder`,  
> car il s'agit d'un objet syst�me int�gr�.  
> Pour injecter automatiquement le d�marrage (`Workbook_Open`), il faut ex�cuter manuellement `injectWorkbookOpenInit`.
