![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/c1bc2456-fb91-46d5-a984-dd913be9a590)

# Series-Plotter : Introduction générale

Ce fichier ***Series plotter(version).xlsm*** est simplement conçu pour accélérer considérablement la production de graphiques Excel et ajouter d'autres types de graphiques non gérés nativement par Excel (par exemple, diagramme ternaire, graphique "temporel" empilé...).
Le fichier comprend également des outils pour les géosciences comme :
  - permettre d'ajouter une grille pour les diagrammes binaires et ternaires, avec une liste de grilles géochimiques déjà implémentées (cf. fichier de grille).
  - permettre de normaliser automatiquement les données géochimiques, avec une liste de normalisations géochimiques déjà implémentées (cf. fichier de normalisation).

Ce n'est jamais parfait, donc vous pourriez rencontrer des problèmes et des erreurs, n'hésitez donc pas à me signaler tout problème.
Il n'y a aucune restriction à l'utilisation et à la distribution des fichiers "Series Plotter", tant que vous mentionnez ce GitHub.

*Version Excel :*
  - devrait fonctionner sur n'importe quelle version antérieure à Excel 2016
  - jamais testé sur une version antérieure à Excel 2016
  - certains utilitaires/outils peuvent ne pas fonctionner sur la version restreinte d'Excel (comme la version étudiante ou l'application)

*Systèmes :*
  - Windows 10 (32 bits) : non testé mais devrait fonctionner
  - Windows 10 (64 bits) : testé et fonctionne
  - Windows 11 (64 bits) : non testé mais devrait fonctionner
  - MacOS (Monterey) : testé et fonctionne pour la plupart des fonctionnalités principales
  
Pour les utilisateurs de MacOS, plusieurs fonctionnalités de développement ne sont pas incluses dans Excel pour Mac, donc je travaille à corriger les bugs et à faire fonctionner toutes les fonctionnalités... mais cela prend du temps.
Liste des problèmes connus : enregistrement en .pdf (mais enregistrement en .svg...), exportation en .ppt (actuellement pas possible nativement)....

**Concepteur : Alexis Derycke** 
  - [Reseach Gate](https://www.researchgate.net/profile/Alexis-Derycke)

# Sommaire :
**- [Comment utiliser le fichier](#part_1)**

**- [Comment configurer la mise en page (couleur, forme,...)](#part_2)**

**- [Type de graphique](#part_3)** : *XY , XYZ (Z comme couleur), diagramme ternaire, XY empilé verticalement, diagramme en toile d'araignée, normalisation, grilles,...*

**- [Comment sélectionner les données pour le "Tracé à partir de la sélection"](#part_4)**

**- [Comment sélectionner les axes X-Y-Z dans la feuille "Liste des graphiques"](#part_5)**


# Comment utiliser le fichier <a id="part_1"></a> 
Il vous suffit de télécharger le fichier, de l'ouvrir et d'autoriser l'exécution des macros. Pour vous aider à comprendre le fichier, il y a des astuces (comme dans Excel réel) qui apparaissent lorsque vous laissez votre souris sur la plupart des boutons :

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/1529f0e0-fc43-4baa-90cd-df656b282ed9)

Vous pouvez également trouver plusieurs vidéos (~2 min) dans le dossier d'aide qui vous présenteront le fichier et vous montreront rapidement comment l'utiliser. Des exemples de données associées sont également disponibles dans le dossier pour tester facilement le traceur de séries et reproduire ce qui est montré dans les vidéos.

Si vous voulez simplement jouer avec le Series Plotter, vous pouvez cliquer sur **"Modèle de données"** dans le groupe ***AIDE*** et cela ajoutera automatiquement un ensemble de données avec lesquelles jouer.

**Tracé à partir de sélections**
  - étape 1 : ouvrir le fichier et **autoriser l'exécution des macros**
  - étape 2 : ajouter des données. Vous pouvez le faire comme pour n'importe quelle autre feuille Excel, ou cliquer sur **"Modèle de données"**
  - étape 3 : sélectionner une plage comme pour tout graphique Excel classique
  - étape 4 : choisir le type de tracé. Pour ce faire, cliquez sur la liste déroulante **"Sélectionner le(s) type(s) de graphique(s)"** dans le ruban et sélectionnez le tracé souhaité (par exemple, "XY - classique")
  - étape 5 : générer le(s) graphique(s). Pour ce faire, cliquez sur le bouton **"Tracé à partir de la sélection"**

**Tracé à partir des en-têtes**
  - étape 1 : ouvrir le fichier et **autoriser l'exécution des macros**
  - étape 2 : ajouter des données. Vous pouvez le faire comme pour n'importe quelle autre feuille Excel, ou cliquer sur **"Modèle de données"**. La structure des données est assez simple : première ligne = en-tête, deuxième ligne = unité, troisième ligne = vide.
  - étape 3 : récupérer vos en-têtes de données. Pour ce faire, cliquez sur le bouton **"Obtenir colonne"**
  - étape 4 : choisir les données à tracer. Pour ce faire, rendez-vous sur la feuille **"Liste graphique"** et sélectionnez la position correspondante dans la grille
  - étape 5 : choisir le type de tracé. Pour ce faire, cliquez sur la liste déroulante **"Sélectionner le(s) type(s) de graphique(s)"** dans le ruban et sélectionnez le tracé souhaité (par exemple, "XY - classique")
  - étape 6 : choisir l'en-tête avec l'ID de l'échantillon. Pour ce faire, cliquez sur le **gauche** déroulant de **"détermination de l'échantillon"** dans le ruban et sélectionnez l'en-tête approprié (par exemple, "Nom") 
  - étape 7 : générer le(s) graphique(s). Pour ce faire, cliquez sur le bouton **"Tracé à partir des en-têtes"**
  
Pour plus d'options, laissez votre souris sur n'importe quel bouton et une info

-bulle apparaîtra, ou lisez les documentations suivantes

https://github.com/ADerycke/Series-Plotter/assets/130437433/d620555c-989c-427b-a564-139f6c200737

# Comment configurer la mise en page (couleur, forme,...) : <a id="part_2"></a> 

Tous les graphiques/schémas sont des graphiques/schémas Excel complets (encore plus complexes comme les graphiques empilés ou le diagramme ternaire), vous pouvez donc les éditer/copier/utiliser comme n'importe quel autre graphique/schéma Excel.

En plus de cette édition manuelle, vous pouvez modifier plusieurs paramètres pour tous les graphiques/schémas automatiquement à l'aide du ruban ou des feuilles "Mise en page".

## Options du ruban :

**OPTION (Série)**

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/7e28bd85-5ad5-4c12-b5ea-2dc64b13573b)

Vous trouverez toutes les options pour changer et récupérer la mise en page des séries. "Définir la mise en page de la série" ouvrira une petite fenêtre pour éditer les mises en page des séries de tous les graphiques en même temps. "Récupérer la couleur de la série" récupérera les couleurs des séries d'un graphique pour le prochain tracé. "Obtenir la mise en page du graphique" vous permet de personnaliser complètement votre graphique et d'utiliser cette mise en page pour un tracé ultérieur. 

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/c3644b2f-bf9b-43fe-bc4d-e5fd3ad3be02)

**OPTION (Graphique)**

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/e1a645b5-b9b3-40aa-b588-8d93d606c719)

Vous trouverez toutes les options liées au graphique

## Options des feuilles "Mise en page" :

Il n'est pas obligatoire d'utiliser ces feuilles, c'est seulement si vous voulez configurer et réutiliser la mise en page de série personnalisée.

Vous pouvez faire un clic droit sur les cellules de la colonne "Couleur" et une fenêtre de sélection de couleurs va apparaître, après la sélection la cellule prendra la couleur que vous avez sélectionnée.

Vous pouvez copier-coller les couleurs obtenues dans toutes les cellules, ça fonctionnera. Si vous supprimez la valeur à l'intérieur de la cellule, alors la couleur revient à "nul", ce qui signifie automatique pour Excel.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/277658b7-6458-4106-8b56-4ff51611f52a)


Si vous sélectionnez plusieurs cellules, alors cela générera automatiquement un dégradé de couleur :

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/a5df3ba4-0fd5-4a5e-a9ee-46f76744accd)


# Type de graphique : <a id="part_3"></a> 

## graphique XY :

![graph XY](https://user-images.githubusercontent.com/130437433/233654680-7dec2505-8e34-4ba6-90ac-d4c969450cd1.png)

**note :** pour le diagramme binaire, plusieurs grilles automatiques sont déjà prêtes à être ajoutées au graphique
![image](https://user-images.githubusercontent.com/130437433/233656030-1236a5e2-c85f-40e5-a94c-f834610e12a8.png)

![graph XY - grille](https://github.com/ADerycke/Series-Plotter/assets/130437433/9b0c2ee8-5b1b-4162-a395-3155466ece9f)

## graphique XYZ (Z comme couleur) :

![Couleur Z - 1](https://github.com/ADerycke/Series-Plotter/assets/130437433/e610d7d9-df17-44d7-8f63-23b37830a7f5)

**note :** les valeurs min et max sont déterminées automatiquement à partir de l'ensemble de données. Vous pouvez sélectionner une troisième couleur et indiquer sa valeur. Vous pouvez également sélectionner un arrondi automatique des valeurs min et max.
![Couleur Z](https://github.com/ADerycke/Series-Plotter/assets/130437433/f04832a0-ecdb-4d51-8944-dd2c03fe736f)

## graphique XYZ (diagramme ternaire) :

![graphique XYZ](https://user-images.githubusercontent.com/130437433/233654854-ad3199e7-f194-4f03-9329-8e0e1341f9b2.png)
![graphique XYZ (bis)](https://user-images.githubusercontent.com/130437433/233799016-01b8f8f8-f5e9-4137-879a-6b409bcc31da.png)

**note :** pour le diagramme ternaire, plusieurs grilles automatiques sont déjà prêtes à être ajoutées au graphique

![image](https://user-images.githubusercontent.com/130437433/233655226-8d13ca9e-ea7e-4495-881e-ff361bf8c55b.png)


## graphique XY - empilé :

![graphique XY - empilé ](https://user-images.githubusercontent.com/130437433/233655423-5dc175fc-9fe8-4177-9faa-4711021abeee.png)

## graphique Y (diagramme de spectres/diagramme en toile d'araignée) :

![graphique Y](https://user-images.githubusercontent.com/130437433/233655533-995ee4f8-3b7d-4268-9f38-8be0b044d1d1.png)

**note :** pour le diagramme en toile d'araignée, une normalisation automatique est possible (message contextuel lors de la sélection).
Cette normalisation sera appliquée/supprimée à vos données en fonction de ce que vous voulez.
Les normalisations disponibles peuvent être trouvées et sélectionnées sur la feuille "Normalisation" qui apparaît après sa sélection.
![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/f1bc2a27-678e-45de-a50b

-678474d599d0)

## Loi normale et calcul KDE :

![graphique Loi normale](https://github.com/ADerycke/Series-Plotter/assets/130437433/97ad0c89-154d-4c71-9103-68ade5a1e215)

## Moyenne pondérée et calcul :

![graphique-Ages  Ma](https://github.com/ADerycke/Series-Plotter/assets/130437433/35dc744b-c9ff-48b1-bfc7-cf359496fc8a)

## X - Histogramme (génération automatique) :

![X - Histogramme](https://github.com/ADerycke/Series-Plotter/assets/130437433/e660c782-9aeb-4ce7-b2be-4eb6e6be83f6)

**note :** lorsque vous sélectionnez ceci, vous entrez la plage souhaitée (ici de 6 à 18) et l'intervalle souhaité (ici 1) et Excel va automatiquement calculer la distribution dans votre ensemble de données

## Ajustement de plusieurs gaussiennes à une distribution KDE :

![graphique détermination de pics](https://github.com/ADerycke/Series-Plotter/assets/130437433/60c02689-7a26-4d8c-9543-2034f2f3c29d)

# Comment sélectionner les données pour le "Tracé à partir de la sélection"  : <a id="part_4"></a> 

  - **graphique XY** : colonne 1 = axe X, colonne 2 = axe Y
  - **graphique XY - empilé** : colonne 1 = axe X et colonne 2 = axe Y *(graphique 1)*; colonne 3 = axe X et colonne 4 = axe Y *(graphique 2)* ; etc...
  - **graphique XYZ (Z comme couleur)** : colonne 1 = axe X, colonne 2 = axe Y, colonne 3 = couleur
  - **graphique XYZ (diagramme ternaire)** : colonne 1 = Gauche, colonne 2 = Haut, colonne 3 = Droit, (colonne 4 = Bas.)
  - **Histogramme X (génération automatique)** : colonne 1 = données

Exemple de sélection pour un diagramme ternaire :
![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/f7bb98f6-7d94-4476-a9d2-0fe2fedf5fcf)


# Comment sélectionner les axes dans la "Liste des graphiques" : <a id="part_5"></a> 

## graphique XY :
Il suffit d'entrer les axes X et Y...

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/886f5d18-8c44-4e45-ae9d-afa8800e0d94)

## graphique XY - empilé :
vous pouvez utiliser des axes X identiques ou différents pour le temps

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/e87bc40f-29a3-4be0-b348-8c5842243cb1)

## graphique XYZ (Z comme couleur) : 
l'axe Z est lu dans la zone comme l'axe X

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/38a5664f-f57c-4a0f-b6aa-891cf1dc143b)

## graphique XYZ (diagramme ternaire) :
le pôle 3 du diagramme ternaire est lu uniquement sur l'axe X (ajoutez un quatrième pour obtenir un double diagramme ternaire)

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/5934a7a8-f547-46f1-97d9-9c83da936181)

## graphique Y (diagramme de spectres/diagramme en toile d'araignée) :
utilisez uniquement l'axe X pour entrer les données à tracer, vous pouvez générer un seul graphique à la fois.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/45f80dea-5df6-49cc-954f-fbc499c11661)

## Histogramme X (génération automatique) :
utilisez uniquement l'axe X pour déterminer les données à tracer (une donnée = un graphique)

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/7730c79d-0107-4d0b-a12a-ff79a5a44c15)

## Graphiques statistiques :
pas besoin de la "Liste des graphiques", il utilise simplement la plage sélectionnée comme pour **"Tracé à partir de la sélection"**. Sélectionnez simplement les données que vous voulez analyser et cliquez sur le bouton approprié.

  - **loi normale et KDE (distribution de base)** : colonne 1 = données, colonne optionnelle deux = erreur
  - **moyenne pondérée et calcul** : colonne 1 = données, colonne deux = erreur
  - **ajustement de plusieurs gaussiennes à une distribution KDE** : colonne 1 = données, optionnel : colonne deux = erreur
