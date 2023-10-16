![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/b424b0f9-6329-48d9-bb85-59c1af6edd63)

# General introduction

This file is simply made for accelerate significantly the Excel chart production, and add other kind of chart not natively handle by Excel (e.g., ternary diagram, stacked "temporal" chart...). The file allows to add grid for binary and ternary diagram, with a list of already implemented geochemical grids (cf. grid files).

It's a "work in progress project" so maybe you gonna encounter some trubble and error, so don't hesitated to reported me any problem.
There is no restriction to use and distribute the "Series Plotter" files, as long as you refer this GitHub. 

*Excel version :*
  - should work on any version greeter than Excel 2016
  - never tested on version older than Excel 2016
  - some very specifiques utility may not work on the restrained version of Excel (as students version or application)

*Systems :*
  - Windows 10 (32bit) : not tested but should work
  - Windows 10 (64bi) : tested and work
  - Windows 11 (64bit) : not tested but should work
  - MacOS (Monterey) : tested and don't work
  
  For the MacOS user, Microsoft removed several developpement facilities from Mac Excel, so i'm not working on a proper version by now, sorry...

**Conceptor : Alexis Derycke** 
  - [Reseach Gate](https://www.researchgate.net/profile/Alexis-Derycke)

# Summary :
**- [How to use the file](#part_1)**

**- [How to set-up the layout (color, shape,...)](#part_2)**

**- [Graphic type](#part_3)**

*XY , XYZ (Z as color), ternary diagram, vertical stacked XY, spider diagram, normalisation, grids,...*

**- [How to selec the X-Y-Z axes in the "Graphics list"](#part_4)**


# How to use the file <a id="part_1"></a> 
You just have to download the file, openning it and allow the macro execution.

You can find several videos (~2min) in the helping folder that gonna introduce you to the file and show you quickly how to use it. You can find associate data example in the folder to easly test the files and reproduce what its show in videos.

**Step by Step simple use :**
  - open the file
  - add some data (you can click on the "template" to see how to input data)
  - click on the "Get column" button
  - select the wanted element to plot in the "Graphic list" sheet
  - select the wanted kind of graph
  - click on "Plot Graphic" button
  
*for more option, let your mouse on any button and a screentips gonna appear*

https://github.com/ADerycke/Series-Plotter/assets/130437433/d620555c-989c-427b-a564-139f6c200737

# How to set-up the layout (color, shape,...) : <a id="part_2"></a> 

All graphs/plots are full Excel graphs/plots (event more complexe one as stacked or ternary plot), so you can edit/copy/use them like any other Excel graphs/plots.

Additionnally, to this manual editing, you can modify several parameters for all graphs/plots automatically using the ribbon or the "Layout" sheets.

## Ribbon options :

This part of the ribbon allow you to change one serie layout at the same time for all the graphs ("Set series color"), or retriver the layout from one graph ("Get series layout").

The second one will store the layout on the sheets "Layout" for the next time you gonna generate plot.

The three next button change to one situation to another (plot point or line, smooth line...). "Axes auto" is for anchore the axes at 0.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/31d3911c-8068-4176-bbef-51996f77ba3a)

The second part of the options is for ON/OFF layout like add/removing the legend, the vertical principal grid...
Each time for all graphs/plots.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/2420d5f6-f9e9-4c6d-8f88-aabf10b29f34)

## Sheets "Layout" options :

It's not mandatory to use this sheets, it's only if you want to setup custom layout.

You can right clik on cells of the "Color" column and a colors selection window gonna appear, after the selection the cell gonna take the color you select.

You can copy and past the color obtain to all cells, it gonna work. If you remove the value inside the cell, then the color go back to "nul", meaning automatic for excel.

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/277658b7-6458-4106-8b56-4ff51611f52a)


If you select multiple cells, then it gonna automatically generate a color gradiant : 

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/a5df3ba4-0fd5-4a5e-a9ee-46f76744accd)


# Graphic type : <a id="part_3"></a> 

## graphic XY :

![graph XY](https://user-images.githubusercontent.com/130437433/233654680-7dec2505-8e34-4ba6-90ac-d4c969450cd1.png)

**note :** for binary diagram, multiple automatic grid are already ready to add to graph
![image](https://user-images.githubusercontent.com/130437433/233656030-1236a5e2-c85f-40e5-a94c-f834610e12a8.png)

![graph XY - grid](https://github.com/ADerycke/Series-Plotter/assets/130437433/9b0c2ee8-5b1b-4162-a395-3155466ece9f)

## graphic XYZ (Z as color) :

![Zcolor - 1](https://github.com/ADerycke/Series-Plotter/assets/130437433/e610d7d9-df17-44d7-8f63-23b37830a7f5)

**note :** the min and max value are determine automatically from the dataset. You can select a third color and indicate it's value. You can also select an automatic round of the min an max value.
![Zcolor](https://github.com/ADerycke/Series-Plotter/assets/130437433/f04832a0-ecdb-4d51-8944-dd2c03fe736f)

## graphic XYZ (ternary diagram) :

![graph XYZ](https://user-images.githubusercontent.com/130437433/233654854-ad3199e7-f194-4f03-9329-8e0e1341f9b2.png)
![graph XYZ (bis)](https://user-images.githubusercontent.com/130437433/233799016-01b8f8f8-f5e9-4137-879a-6b409bcc31da.png)

**note :** for ternary diagram, multiple automatic grid are already ready to add to graph

![image](https://user-images.githubusercontent.com/130437433/233655226-8d13ca9e-ea7e-4495-881e-ff361bf8c55b.png)


## graphic XY - stacked :

![graph XY - stacked ](https://user-images.githubusercontent.com/130437433/233655423-5dc175fc-9fe8-4177-9faa-4711021abeee.png)

## graphic Y (spectra/spider diagram) :

![graph Y](https://user-images.githubusercontent.com/130437433/233655533-995ee4f8-3b7d-4268-9f38-8be0b044d1d1.png)

**note :** for spider diagram, an automatic normalisation is possible (msg pop-up at the selection).
This normalisation gonna be applied/remove to your data in fonction of what you want.
The available normalisation can be found and selected on the "Normalisation" sheet that apprear after it selection.
![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/f1bc2a27-678e-45de-a50b-678474d599d0)

## Normal law and KDE (basic distribution) :

![graph Normal law](https://github.com/ADerycke/Series-Plotter/assets/130437433/97ad0c89-154d-4c71-9103-68ade5a1e215)

## X - Histogram (automatic generation) :

![X - Histogram](https://github.com/ADerycke/Series-Plotter/assets/130437433/e660c782-9aeb-4ce7-b2be-4eb6e6be83f6)

**note :** when you select this, you enter the wanted range (here 6 to 18) and the wanted intervalle (here 1) and excel gonna automatically calculate the distribution in your dataset


# How to selec the X-Y-Z axes in the "Graphics list" : <a id="part_4"></a> 

## graphic XY :
Just enter X and Y axes classicaly...

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/886f5d18-8c44-4e45-ae9d-afa8800e0d94)

## graphic XYZ (Z as color) :
the Z axe is read at the same position a X axe

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/b77def31-8547-4183-8472-4f9240847829)

## graphic XYZ (ternary diagram) :
the 3 pole of the ternary diagramme is read only on the Y axe (add a third one to get a double ternary diagram)

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/8aea9da8-be10-4dc4-aa75-7b892bd0f987)

## graphic XY - stacked :
you can use same or differente X axes for the time

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/e87bc40f-29a3-4be0-b348-8c5842243cb1)

## graphic Y (spectra/spider diagram) :
use only the Y axe to enter the data to plot (data in the same row will be plotted together)

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/a319a9e3-47b7-438c-997b-8b320f0488f7)

## Normal law and KDE (basic distribution) :
no need of the "Graphics list"

## X - Histogram (automatic generation) :
use only the X axe to determine the data to plot

![image](https://github.com/ADerycke/Series-Plotter/assets/130437433/f21c0779-f18b-4f8e-89cf-2969ae9236b9)

