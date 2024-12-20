![header](https://github.com/ADerycke/Better_Plot/assets/130437433/adff3a48-21ea-427e-aa40-11848ce527e7)
[![fr](https://img.shields.io/badge/lang-fr-red.svg)](https://github.com/ADerycke/Better_Plot/blob/main/README.fr.md)

**IMPORTANT : AFTER SEPTEMBER 2024, THE OFFICIAL HELPS / SUPPORT and DOWNLOAD HAVE BEEN MOVE TO THE WEBSITE :**
[DeryckeHub](https://deryckehub.ovh/en/better-plot-excel/)

# Better Plot : General introduction

This file ***Better Plot(version).xlsm*** is simply made to accelerate significantly Excel chart production, and add other kind of chart not natively handle by Excel (e.g., ternary diagram, stacked "temporal" chart...).
The file also include tools for geosciences like :
  - allowing to add grid for binary and ternary diagram, with a list of already implemented geochemical grids (cf. grid file).
  - allowing to automatically normalise geochemical data, with a list of already implemented geochemical normalisation (cf. normalisation file).

It's never perfect so maybe you gonna encounter some troubles and errors, so don't hesitate to report me any problem.
There is no restriction to use and distribute the "Better Plot" files, as long as you refer this GitHub. 

*Excel version :*
  - should work on any version earlier than Excel 2016
  - never tested on version older than Excel 2016
  - some utilities/tools may not work on the restrained version of Excel (as students version or application)

*Systems :*
  - Windows 10 (32bit) : not tested but should work
  - Windows 10 (64bit) : tested and work
  - Windows 11 (64bit) : not tested but should work
  - MacOS (Monterey) : tested and work for most of the main functionnalities
  
For the MacOS user, several development facilities are not included in Mac Excel, so i'm working on correcting bug and make work all the functionnalities... but it take time.
List of know problems : save as .pdf (but save as .svg...), export in a .ppt (not natively possible by now)....

**Conceptor : Alexis Derycke** 
  - [Reseach Gate](https://www.researchgate.net/profile/Alexis-Derycke)

# Summary :
**- [How to use the file](#part_1)**

**- [How to set-up the layout (color, shape,...)](#part_2)**

**- [Graphic type](#part_3)** : *XY , XYZ (Z as color), ternary diagram, vertical stacked XY, geological time scale, spider diagram, normalisation, grids, better regression...*

**- [How to select the data for the "Plot from selection"](#part_4)**

**- [How to select the X-Y-Z axes in the "Graphics list" sheet](#part_5)**


# How to use the file <a id="part_1"></a> 
You just have to download the file, opening it and allow the macro execution. To help you to understand the file, there is some tips (like the real excel) that pop if you let your mouse above most of the button :

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/eb6d4f7c-8c7e-4607-9fbf-df08c61243ee)

You also can find several videos (~2min) in the helping folder that gonna introduce you to the file and show you quickly how to use it. Associate data examples are in the folder to easily test the Better Plot and reproduce what its show in videos.

If you want to simply play with Better Plot, you can click on **"Data template"** in the ***HELP*** group and it gonna automatically add a dataset to play with.

## ![image](https://github.com/ADerycke/Better_Plot/assets/130437433/f8864374-40ed-4934-970f-d188de367301) Plot from selections 
  - step 1 : open the file and **allow macro execution**
  - step 2 : add some data to the "Data & Graphics" sheet. You can do it as any other excel sheets, or click on **"Data template"** in the *HELP*
  - step 3 : select a range like for any classical Excel plot (like two columns for example)
  - step 4 : chose the type of plot. To do so, clik on dropdown **"Select the graphic(s) type"** in the ribbon and select the wanted plot (e.g. "XY - classic")
  - step 5 : generate the graphic(s). To do so, click on the **"Plot from selections"** button

## ![image](https://github.com/ADerycke/Better_Plot/assets/130437433/b03fb12b-7305-470a-8800-2e6a827c6045) Plot from headers 
  - step 1 : open the file and **allow macro execution**
  - step 2 : add some data to the "Data & Graphics" sheet. You can do it as any other excel sheets, or click on **"Data template"** in the *HELP*. The data structuration is quite simple : first row = hearder, second row = unit, thrid row = empty.
  - step 3 : retriver your data header. To do so, click on the **"Get column"** button
  - step 4 : chose the data to plot. To do so, go on the **"Graphic list"** sheet and select the corresponding position in the grid
  - step 5 : chose the type of plot. To do so, clik on dropdown **"Select the graphic(s) type"** in the ribbon and select the wanted plot (e.g. "XY - classic")
  - step 6 : chose the header with sample ID. To do so, clik on the **left** dropdown of **"sample determination"** in the ribbon and select the appropriate header (e.g. "Name") 
  - step 7 : generate the graphic(s). To do so, click on the **"Plot from headers"** button
  
For more options, let your mouse on any button and a screen tips gonna appear, or read the following documentations

https://github.com/ADerycke/Better_Plot/assets/130437433/d620555c-989c-427b-a564-139f6c200737

## Plot data uncertainty and / or error 
If you use **"Plot from headers"** the excel file can automatically handle the error-bar add. To do so, you just have to add the data uncertainty/error in the following column with the headers including "±".
The error can be enter as absolute ("*[unit]*") or relative ("*[%]*"), if uncertainty/error column don't have any unit, it will be considere at absolut error, otherwise you have to add "%" to the header unit.

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/5c0453ab-b71b-4dec-a8ee-40c4cf308d8f)

If you want, you can automatically had uncertainty/error header by using the **"Data tools"** option.

# How to set-up the layout (color, shape,...) : <a id="part_2"></a> 

All graphs/plots are full Excel graphs/plots (event more complex one as stacked or ternary plot), so you can edit/copy/use them like any other Excel graphs/plots.

In additiona to this manual editing, you can modify several parameters for all graphs/plots automatically using the ribbon or the "Layout" sheets.

## Ribbon options :

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/f5726fca-0615-4111-880f-b64a5a9e9882)

This include different options to help you to handle your data, like :
  - clean all the *Data & Graphics* sheets
  - converting cells see as "string" by Excel to "number" properly recognise
  - add a "sarting row" to lets some row above the headers (to enter data information, constante, etc)
  - do an automatic layout for the headers
  - etc

You can find all the options to change and retriver the series layout :
- "series layout" will open a small windows to edit series layouts of all graph at the same time
- "retriver series color" will retriver the series colors form a graph for the next plotting
- "retriver graph layout" allow you to completely custom your graph and use this layout for latter plot 

## Right click options :

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/8d35ef13-2967-419a-9b1e-3cf0d257bf27)

You can right click on series to :
  - add regression to data with associated predicted envelop
  - ad 95% enveloppe of an XY dataset
  - add min/max/mean of an Y dataset
  - add first, last quartile and median of an Y dataset
  - add a moving average on a XY temporal dataset

Or right click on a axe to :
  - add a constant
  - replace the axe by a time scale

![RC - regression](https://github.com/ADerycke/Better_Plot/assets/130437433/e3e0f424-9648-4df2-a0f0-28d94fcafa93)

![RC - axe et autre](https://github.com/ADerycke/Better_Plot/assets/130437433/98a5db70-c627-44f5-9b3f-39e671d5a1b2)

## Sheets "Layout" options :

It's not mandatory to use this sheets, it's only if you want to setup and re-use custom series layout.

You can right click on cells of the "Color" column and a colors selection window gonna appear, after the selection the cell gonna take the color you select.

You can copy and past the color obtains to all cells, it gonna work. If you remove the value inside the cell, then the color goes back to "nul", meaning automatic to excel.

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/5905e8ab-845a-492a-b15b-766923497400)


If you select multiple cells, then it gonna automatically generate a color gradient : 

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/a5df3ba4-0fd5-4a5e-a9ee-46f76744accd)


# Graphic type : <a id="part_3"></a> 

## graphic XY (grid) :

![graph XY](https://user-images.githubusercontent.com/130437433/233654680-7dec2505-8e34-4ba6-90ac-d4c969450cd1.png)

**option 1 :** you can plot error as classical bar or as elipse by selecting it in "plot options"

![graph XY - elispe](https://github.com/ADerycke/Better_Plot/assets/130437433/cd408c8c-d0b4-4ba8-bb3a-8585cd32ee6d)

**option 2 :** for binary diagram, multiple automatic grid are already ready to add to graph

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/b91ff562-7397-4a52-9709-012e1763ad97)

![graph XY - grid](https://github.com/ADerycke/Better_Plot/assets/130437433/9b0c2ee8-5b1b-4162-a395-3155466ece9f)

## graphic XYZ (Z as color) :

![Zcolor - 1](https://github.com/ADerycke/Better_Plot/assets/130437433/e610d7d9-df17-44d7-8f63-23b37830a7f5)

**note :** min and max value are determined automatically from the dataset. You can select a third color and indicate its value. You can also select an automatic round of the min and max value.
![Zcolor](https://github.com/ADerycke/Better_Plot/assets/130437433/f04832a0-ecdb-4d51-8944-dd2c03fe736f)

## graphic XYZ (ternary diagram) :

![graph XYZ](https://user-images.githubusercontent.com/130437433/233654854-ad3199e7-f194-4f03-9329-8e0e1341f9b2.png)
![graph XYZ (bis)](https://user-images.githubusercontent.com/130437433/233799016-01b8f8f8-f5e9-4137-879a-6b409bcc31da.png)

**note :** for ternary diagram, multiple automatic grid are already ready to add to graph

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/205e4690-c9d1-42fb-b799-24617903c992)


## graphic XY - stacked :

![graph stacked X](https://github.com/ADerycke/Better_Plot/assets/130437433/1cdc605f-5fa7-43cc-b923-146d6e6907d0)
![graph stacked Y](https://github.com/ADerycke/Better_Plot/assets/130437433/424ffc90-501b-4565-8f31-de4389a7a6bd)

**note :** the scale change is done thanks to optional parameters

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/906d70f2-8ea0-4fa5-8598-abe21cf55598)

## Geological time scale :

![timescale](https://github.com/ADerycke/Better_Plot/assets/130437433/e2a2a480-f3c0-470a-bd72-b2738c528246)

**note :** this is a classical excel plot, so you can edit it (min, max, size...) as any other excel chart. By now the added Geological Time Scale correspond to the 2022. It was added using the work of [M. J. Williams](https://github.com/morganjwilliams/pyrolite.git). After generation, the graph will include the Period name (hidded in the picture above)

You can select between different resolution and/or add any time scale using the button **"Add a time scale"** on the **"Other options"**. If you select a graph/chart before clicking on *Plot*, the time scale will replace the corresponding axes.

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/ab73f19b-0708-47fc-8c94-eaa1810ad092) 

![timescale 2](https://github.com/ADerycke/Better_Plot/assets/130437433/8b4bc925-85f9-41dc-80b2-af5adbeca7fa)

## graphic Y (spectra/spider diagram) :

![graph Y](https://user-images.githubusercontent.com/130437433/233655533-995ee4f8-3b7d-4268-9f38-8be0b044d1d1.png)

![graph Y - mean](https://github.com/ADerycke/Better_Plot/assets/130437433/7590a88f-121a-43df-baae-4e22718afd64)

**note :** for spider diagram, an automatic normalisation is possible.
This normalisation gonna be applied/remove to your data in fonction of what you want.
The available normalisation can be found and selected on the "Normalisation" sheet that apprear after it selection.

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/1d9a28cf-283c-4c99-aca3-8910d5272610)


## regression and mobil average

![lin  reg](https://github.com/ADerycke/Better_Plot/assets/130437433/a5ec4279-b471-49be-8cd3-a8ce6e3b966e)

![moving average](https://github.com/ADerycke/Better_Plot/assets/130437433/5090d6c0-4f29-4ecd-aee9-4c46ca08838e)

**note :** you can add this to any series by rigth clicking on a serie :

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/43b6ec6f-2a41-4186-bd73-0d9077dbb7f7)

## normal law and KDE calculation :

![graph Normal law](https://github.com/ADerycke/Better_Plot/assets/130437433/97ad0c89-154d-4c71-9103-68ade5a1e215)

## weigth mean plot and calculation :

![graph-Ages  Ma](https://github.com/ADerycke/Better_Plot/assets/130437433/35dc744b-c9ff-48b1-bfc7-cf359496fc8a)

## X - Histogram (automatic generation) :

![X - Histogram](https://github.com/ADerycke/Better_Plot/assets/130437433/e660c782-9aeb-4ce7-b2be-4eb6e6be83f6)

**note :** when you select this, you enter the wanted range (here 6 to 18) and the wanted interval (here 1) and excel gonna automatically calculate the distribution in your dataset

## Fit of multiple gaussian to a KDE distribution :

![graph preak determination](https://github.com/ADerycke/Better_Plot/assets/130437433/60c02689-7a26-4d8c-9543-2034f2f3c29d)

## Radial plot design by [Galbraith (1988)](https://www.tandfonline.com/doi/abs/10.1080/00401706.1988.10488400) for error visualisation

![graph - radial plot](https://github.com/ADerycke/Better_Plot/assets/130437433/f61b6967-85a9-4390-95e1-0bc8ec16b74f)

**note :** by now this plot is only a representation, I'm currently waiting python implementation in excel to allow complex statistic calculation in it.

# How to select the data for the "Plot from selection"  : <a id="part_4"></a> 

  - **graphic XY** : column 1 = axe X, column 2 = axe Y
  - **graphic XY - stacked** : column 1 = axe X and column 2 = axe Y *(chart 1)*; column 3 = axe X and column 4 = axe Y *(chart 2)* ; etc...
  - **graphic XYZ (Z as color)** : column 1 = axe X, column 2 = axe Y, column 3 = color
  - **graphic XYZ (ternary diagram)** : column 1 = Left, column 2 = Top, column 3 = Right, (column 4 = Bot.)
  - **X - Histogram (automatic generation)** : column 1 = data

Example of selection for a ternary diagram :
![image](https://github.com/ADerycke/Better_Plot/assets/130437433/f7bb98f6-7d94-4476-a9d2-0fe2fedf5fcf)


# How to select the axes in the "Graphics list" : <a id="part_5"></a> 

## graphic XY :
Just enter X and Y axes ...

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/886f5d18-8c44-4e45-ae9d-afa8800e0d94)

## graphic XY - stacked :
you can use same or different X axes for the time

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/e87bc40f-29a3-4be0-b348-8c5842243cb1)

## graphic XYZ (Z as color) : 
the Z axe is read in the area as the X axe

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/38a5664f-f57c-4a0f-b6aa-891cf1dc143b)

## graphic XYZ (ternary diagram) :
the 3 pole of the ternary diagram is read only on the X axe (add a fourth one to get a double ternary diagram)

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/5934a7a8-f547-46f1-97d9-9c83da936181)

## graphic Y (spectra/spider diagram) :
use only the X axe to enter the data to plot, you can generate only one chart at time.

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/45f80dea-5df6-49cc-954f-fbc499c11661)

## X - Histogram (automatic generation) :
use only the X axe to determine the data to plot (one data = one chart)

![image](https://github.com/ADerycke/Better_Plot/assets/130437433/7730c79d-0107-4d0b-a12a-ff79a5a44c15)

## Statistical plots :
no need of the "Graphics list", it just use the selected range as for **"Plot from selection"**. Just select the data you want to analyse and clik on the proper button.

  - **normal law and KDE (basic distribution)** : column 1 = data, optinal column two = error
  - **weigth mean plot and calculation** : column 1 = data, column two = error
  - **fit of multiple gaussian to a KDE distribution** : column 1 = data, optinal : column two = error
