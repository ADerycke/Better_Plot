![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/2aa0d052-8c53-41d8-b906-e5547c0a2f9e)

[![fr](https://img.shields.io/badge/lang-fr-red.svg)](https://github.com/ADerycke/Series-Plotter/blob/main/README.fr.md)
# Series-Plotter : Introduction générale

Ce fichier ***Series plotter(version).xlsm*** est simplement conçu pour accélérer significativement la production de graphiques Excel et ajouter d'autres types de graphiques non gérés nativement par Excel (par exemple, diagramme ternaire, graphique "temporel" empilé...).
Le fichier propose également des outils pour les géosciences comme :
  - permettant d'ajouter une grille pour les diagrammes binaires et ternaires, avec une liste de grilles géochimiques déjà implémentées (cf. fichier de grille).
  - permettant de normaliser automatiquement les données géochimiques, avec une liste de normalisations géochimiques déjà implémentées (cf. fichier de normalisation).

Ce n'est jamais parfait donc vous pourriez rencontrer des problèmes et des erreurs, n'hésitez donc pas à me signaler tout problème.
Il n'y a aucune restriction à l'utilisation et à la distribution du fichier "Series Plotter", tant que vous faites référence à ce GitHub.

*Version Excel :*
  - devrait fonctionner sur n'importe quelle version posterieur à Excel 2016
  - jamais testé sur une version antérieure à Excel 2016
  - certains utilitaires/outils peuvent ne pas fonctionner sur la version restreinte d'Excel (comme la version étudiante ou l'application)

*Systèmes :*
  - Windows 10 (32 bits) : non testé mais devrait fonctionner
  - Windows 10 (64 bits) : testé et fonctionne
  - Windows 11 (64 bits) : non testé mais devrait fonctionner
  - MacOS (Monterey) : testé et fonctionne pour la plupart des fonctionnalités principales

Pour les utilisateurs de MacOS, plusieurs fonctionnalités de développement ne sont pas incluses dans Excel pour Mac, donc je travaille sur la correction des bugs et pour que toutes les fonctionnalités fonctionnent... mais cela prend du temps.
Liste des problèmes connus : enregistrement en .pdf (mais enregistrement en .svg...), exportation en .ppt (pas nativement possible pour le moment)....

**Concepteur : Alexis Derycke** 
  - [Reseach Gate](https://www.researchgate.net/profile/Alexis-Derycke)

# Résumé :
**- [Comment utiliser le fichier](#part_1)**

**- [Comment configurer la mise en page (couleur, forme,...)](#part_2)**

**- [Type de graphique](#part_3)** : *XY , XYZ (Z comme couleur), diagramme ternaire, graphique empilé vertical XY, échelle de temps géologique, diagramme en radar, normalisation, grilles,...*

**- [Comment sélectionner les données pour le "Graphique à partir de la sélection"](#part_4)**

**- [Comment sélectionner les axes X-Y-Z dans la feuille "Liste des graphiques"](#part_5)**


# Comment utiliser le fichier <a id="part_1"></a> 
Vous devez simplement télécharger le fichier, l'ouvrir et autoriser l'exécution des macros. Pour vous aider à comprendre le fichier, des astuces s'affichent lorsque vous laissez votre souris sur la plupart des boutons, comme dans Excel :

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/af88da3c-7cc5-4235-8cd5-ddbc0d933168)

Vous pouvez également trouver plusieurs vidéos (~2min) dans le dossier d'aide qui vont vous présenter le fichier et vous montrer rapidement comment l'utiliser. Des exemples de données associées sont également disponibles dans le dossier pour tester facilement le series plotter et reproduire ce qui est montré dans les vidéos.

Si vous voulez simplement jouer avec Series Plotter, vous pouvez cliquer sur **"Modèle de données"** dans le groupe ***AIDE*** et cela ajoutera automatiquement un jeu de données pour jouer.

## Tracé à partir de sélections ![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/f8864374-40ed-4934-970f-d188de367301)
  - étape 1 : ouvrez le fichier et **autorisez l'exécution des macros**
  - étape 2 : ajoutez des données. Vous pouvez le faire comme pour n'importe quelle autre feuille Excel, ou cliquez sur **"Modèle de données"**
  - étape 3 : sélectionnez une plage comme pour n'importe quel graphique Excel classique
  - étape 4 : choisissez le type de graphique. Pour ce faire, cliquez sur la liste déroulante **"Sélectionnez le(s) type(s) de graphique"** dans le ruban et sélectionnez le graphique souhaité (par exemple, "XY - classique")
  - étape 5 : générez le(s) graphique(s). Pour ce faire, cliquez sur le bouton **"Graphique à partir des sélections"**

## Tracé à partir des en-têtes ![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/b03fb12b-7305-470a-8800-2e6a827c6045)
  - étape 1 : ouvrez le fichier et **autorisez l'exécution des macros**
  - étape 2 : ajoutez des données. Vous pouvez le faire comme pour n'importe quelle autre feuille Excel, ou cliquez sur **"Modèle de données"**. La structuration des données est assez simple : première ligne = en-tête, deuxième ligne = unité, troisième ligne = vide.
  - étape 3 : récupérez vos en-têtes de données. Pour ce faire, cliquez sur le bouton **"Obtenir la colonne"**
  - étape 4 : choisissez les données à tracer. Pour ce faire, allez dans la feuille **"Liste graphique"** et sélectionnez la position correspondante dans la grille
  - étape 5 : choisissez le type de graphique. Pour ce faire, cliquez sur la liste déroulante **"Sélectionnez le(s) type(s) de graphique"** dans le ruban et sélectionnez le graphique souhaité (par exemple, "XY -

 classique")
  - étape 6 : choisissez l'en-tête avec l'ID de l'échantillon. Pour ce faire, cliquez sur la liste déroulante **gauche** de **"détermination de l'échantillon"** dans le ruban et sélectionnez l'en-tête approprié (par exemple, "Nom") 
  - étape 7 : générez le(s) graphique(s). Pour ce faire, cliquez sur le bouton **"Graphique à partir des en-têtes"**
  
Pour plus d'options, laissez votre souris sur n'importe quel bouton et une info-bulle va apparaître, ou lisez les documentations suivantes

https://github.com/ADerycke/Series-Plotter/assets/130437433/d620555c-989c-427b-a564-139f6c200737

## Tracé de l'incertitude et/ou de l'erreur des données 
Si vous utilisez "**Tracé à partir des en-têtes**", le fichier Excel peut gérer automatiquement l'ajout de barres d'erreur. Pour ce faire, il vous suffit d'ajouter l'incertitude/l'erreur des données dans la colonne suivante avec les en-têtes incluant "±".
L'erreur peut être saisie en absolu ou en relatif, si la colonne d'incertitude/erreur n'a pas d'unité, elle sera considérée comme une erreur absolue, sinon vous devez ajouter "%" à l'unité de l'en-tête.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/5c0453ab-b71b-4dec-a8ee-40c4cf308d8f)

Si vous le souhaitez, vous pouvez ajouter automatiquement les en-têtes d'incertitude/erreur en utilisant l'option **"Outils de données"**.

# Comment configurer la mise en page (couleur, forme,...) : <a id="part_2"></a> 

Tous les graphiques sont des graphiques Excel complets (même les plus complexes comme les graphiques empilés ou les graphiques ternaires), donc vous pouvez les éditer/copier/utiliser comme n'importe quel autre graphique Excel.

En plus de cette édition manuelle, vous pouvez modifier plusieurs paramètres pour tous les graphiques automatiquement à l'aide du ruban ou des feuilles "Mise en page".

## Options du ruban :

***Outils de données* dans DONNÉES**

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/ed994745-b7fd-4f3e-b549-d3f134942bf5)

Cela comprend différentes options pour vous aider à manipuler vos données, comme :
  - effacer toutes les feuilles *Données & Graphiques*
  - convertir les cellules considérées comme des "chaînes" par Excel en "nombres" correctement reconnus
  - ajouter une "ligne de départ" pour laisser quelques lignes au-dessus des en-têtes (pour saisir des informations de données, des constantes, etc.)
  - faire une mise en page automatique pour les en-têtes
  - etc

**OPTIONS**

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/b23a7806-a909-4ea6-a073-469734cdba13)

Vous pouvez trouver toutes les options pour modifier et récupérer la mise en page des séries :
- "mise en page des séries" ouvrira une petite fenêtre pour éditer les mises en page des séries de tous les graphiques en même temps
- "récupérer les couleurs des séries" récupérera les couleurs des séries à partir d'un graphique pour le prochain tracé
- "récupérer la mise en page du graphique" vous permet de personnaliser entièrement votre graphique et d'utiliser cette mise en page pour un tracé ultérieur

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/c3644b2f-bf9b-43fe-bc4d-e5fd3ad3be02)


## Options de la feuille "Mise en page" :

Il n'est pas obligatoire d'utiliser cette feuille, c'est seulement si vous voulez configurer et réutiliser des mises en page de séries personnalisées.

Vous pouvez cliquer avec le bouton droit sur les cellules de la colonne "Couleur" et une fenêtre de sélection des couleurs va apparaître, après la sélection la cellule prendra la couleur que vous avez choisie.

Vous pouvez copier-coller les couleurs obtenues dans toutes les cellules, cela fonctionnera. Si vous supprimez la valeur à l'intérieur de la cellule, alors la couleur redeviendra "nulle", c'est-à-dire automatique pour Excel.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/277658b7-6458-4106-8b56-4ff51611f52a)


Si vous sélectionnez plusieurs cellules, cela générera automatiquement un dégradé de couleur : 

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/a5df3ba4-0fd5-4a5e-a9ee-46f76744accd)


# Type de graphique : <a id="part_3"></a> 

## graphique XY (grille) :

![graphique XY](https://user-images.githubusercontent.com/130437433/233654680-7dec2505-8e34-4ba6-90ac-d4c969450cd1.png)

**option 1 :** vous pouvez tracer l'erreur sous forme de barre classique ou sous forme d'ellipse en le sélectionnant dans "options de tracé"

![graphique XY - ellipse](https://github.com/ADerycke/Series-Plotter/assets/130437433/cd408c8c-d0b4-4ba8-bb3a-8585cd32ee6d)

**option 2 :** pour le diagramme binaire, plusieurs grilles automatiques sont déjà prêtes à être ajoutées au graphique
![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/50454a58-6944-4741-8c21-04db95056ee5)

![graphique XY - grille](https://github.com/ADerycke/Series-Plotter/assets/130437433/9b0c2ee8-5b1b-4162-a395-3155466ece9f)

## graphique XYZ (Z comme couleur) :

![Couleur Z - 1](https://github.com

/ADerycke/Series-Plotter/assets/130437433/e610d7d9-df17-44d7-8f63-23b37830a7f5)

**note :** la valeur minimale et maximale est déterminée automatiquement à partir de l'ensemble de données. Vous pouvez sélectionner une troisième couleur et indiquer sa valeur. Vous pouvez également sélectionner un arrondi automatique de la valeur minimale et maximale.
![Couleur Z](https://github.com/ADerycke/Series-Plotter/assets/130437433/f04832a0-ecdb-4d51-8944-dd2c03fe736f)

## graphique XYZ (diagramme ternaire) :

![graphique XYZ](https://user-images.githubusercontent.com/130437433/233654854-ad3199e7-f194-4f03-9329-8e0e1341f9b2.png)
![graphique XYZ (bis)](https://user-images.githubusercontent.com/130437433/233799016-01b8f8f8-f5e9-4137-879a-6b409bcc31da.png)

**note :** pour le diagramme ternaire, plusieurs grilles automatiques sont déjà prêtes à être ajoutées au graphique

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/c98abc04-d290-48e3-b097-2f80aeef22c4)


## graphique XY - empilé :

![graphique XY - empilé ](https://user-images.githubusercontent.com/130437433/233655423-5dc175fc-9fe8-4177-9faa-4711021abeee.png)


## Échelle temporelle géologique :

![échelle temporelle](https://github.com/ADerycke/Series-Plotter/assets/130437433/e2a2a480-f3c0-470a-bd72-b2738c528246)

**note :** il s'agit d'un graphique Excel classique, vous pouvez donc le modifier (min, max, taille...) comme n'importe quel autre graphique Excel. Pour l'instant, l'échelle temporelle géologique ajoutée correspond à 2022. Elle a été ajoutée en utilisant le travail de [M. J. Williams](https://github.com/morganjwilliams/pyrolite.git). Après génération, le graphique inclura le nom de la période (caché sur l'image ci-dessus)

Vous pouvez sélectionner entre différentes résolutions et/ou ajouter n'importe quelle échelle temporelle en utilisant le bouton **"Ajouter une échelle temporelle"** sur les "Options de la fenêtre". Si vous sélectionnez un graphique avant de cliquer sur *Tracer*, l'échelle temporelle remplacera les axes correspondants.
![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/ab73f19b-0708-47fc-8c94-eaa1810ad092) 

![échelle temporelle 2](https://github.com/ADerycke/Series-Plotter/assets/130437433/8b4bc925-85f9-41dc-80b2-af5adbeca7fa)

## graphique Y (spectres/diagramme en radar) :

![graphique Y](https://user-images.githubusercontent.com/130437433/233655533-995ee4f8-3b7d-4268-9f38-8be0b044d1d1.png)

**note :** pour le diagramme en radar, une normalisation automatique est possible.
Cette normalisation sera appliquée/supprimée à vos données en fonction de ce que vous voulez.
La normalisation disponible peut être trouvée et sélectionnée sur la feuille "Normalisation" qui apparaît après sa sélection.
![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/044f0f23-3541-406e-b56b-fdb89edb459d)


## Loi normale et calcul KDE :

![loi normale](https://github.com/ADerycke/Series-Plotter/assets/130437433/97ad0c89-154d-4c71-9103-68ade5a1e215)

## calcul de la moyenne pondérée et tracé :

![graphique-Ages  Ma](https://github.com/ADerycke/Series-Plotter/assets/130437433/35dc744b-c9ff-48b1-bfc7-cf359496fc8a)

## Histogramme X (génération automatique) :

![Histogramme X](https://github.com/ADerycke/Series-Plotter/assets/130437433/e660c782-9aeb-4ce7-b2be-4eb6e6be83f6)

**note :** lorsque vous sélectionnez ceci, vous entrez la plage souhaitée (ici 6 à 18) et l'intervalle souhaité (ici 1) et Excel calculera automatiquement la distribution dans votre ensemble de données

## Ajustement de plusieurs gaussiennes à une distribution KDE :

![graphique détermination des pics](https://github.com/ADerycke/Series-Plotter/assets/130437433/60c02689-7a26-4d8c-9543-2034f2f3c29d)

## Conception du graphique radial par [Galbraith (1988)](https://www.tandfonline.com/doi/abs/10.1080/00401706.1988.10488400) pour la visualisation des erreurs

![graphique - graphique radial](https://github.com/ADerycke/Series-Plotter/assets/130437433/f61b6967-85a9-4390-95e1-0bc8ec16b74f)

**note :** pour l'instant, ce graphique n'est qu'une représentation, j'attends actuellement une implémentation en Python dans Excel pour permettre des calculs statistiques complexes à l'intérieur.

# Comment sélectionner les données pour le "Tracé à partir de la sélection"  : <a id="part_4"></a> 

  - **graphique XY** : colonne 1 = axe X, colonne 2 = axe Y
  - **graphique XY - empilé** : colonne 1 = axe X et colonne 2 = axe Y *(graphique 1)* ; colonne 3 = axe X et colonne 4 = axe Y *(graphique 2)* ; etc...
  - **graphique XYZ (Z comme couleur)** : colonne 1 = axe X, colonne 2 = axe Y, colonne 3 = couleur
  - **graphique XYZ (diagramme

 ternaire)** : colonne 1 = axe A, colonne 2 = axe B, colonne 3 = axe C
  - **graphique Y (spectres/diagramme en radar)** : colonne 1 = axe Y, et toutes les autres colonnes sont les valeurs
  - **Loi normale et calcul KDE** : colonne 1 = axe X
  - **calcul de la moyenne pondérée et tracé** : colonne 1 = axe X, colonne 2 = axe Y, colonne 3 = axe Z
  - **Histogramme X (génération automatique)** : colonne 1 = axe X
  - **Ajustement de plusieurs gaussiennes à une distribution KDE** : colonne 1 = axe X

Pour certains graphiques, vous pouvez également sélectionner un échantillon et un en-tête avec l'ID de l'échantillon.

# Comment sélectionner les axes X-Y-Z dans la feuille "Liste des graphiques"  : <a id="part_5"></a> 

Vous pouvez sélectionner vos données sur la feuille "Liste des graphiques", pour ce faire :
  - allez à la feuille "Liste des graphiques"
  - trouvez la colonne "Categorie du graphique" et choisissez le graphique que vous voulez
  - allez à la colonne "X, Y, Z Axis"
  - sélectionnez l'axe souhaité pour votre graphique dans la cellule correspondante
  - générez le graphique à partir de la sélection

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/799b9a10-2689-4da8-a4e7-77c4870b7b49)

N'oubliez pas qu'il existe également un autre onglet "Graphique" qui contient les boutons permettant de générer les graphiques et qui vous guidera dans la sélection des données.
