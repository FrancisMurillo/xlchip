xlchip
====

A VBA module manager that installs other related modules for you.

When creating a VBA project, I find myself looking for two things: a testing framework and a code import/export tool.
The latter being more important as I can setup a code management process for my VBA modules to a VCS such as Git or Mercurial. VB is forthcoming with these so I created <a href="https://github.com/FrancisMurillo/xlvase">vase</a> and <a href="https://github.com/FrancisMurillo/xlwheat">wheat</a> to respectively address the issue in some way. But installing the projects by importing the modules manually is such a chore, that's why this was born.

This module manager downloads the project files from the net and installs the modules specified by that project. However, this isn't a module namespace resolution tool and all that. This is just an import tool for the pragmatic VBA developer who just wants to install the modules to start working. 

This is called *chip* because you can attach more *chips* or modules to your circuit. To "*chip on*" is the phrase here. 

Currently, I'm using this for VBA Excel development in my company since there is no development process given. A process is better than no process at all, a process smell if there was one. 

**To God I plant this seed, may it turn into a forest**

quick start
====

There's only one thing you need to import for your project: <a href="https://raw.githubusercontent.com/FrancisMurillo/xlchip/master/Modules/ChipInit.bas">ChipInit.bas</a>. You can create a new module and copy paste the code but importing the module is the most efficient way to go.

Also you need to include these references to your project before running any commands. These are:
* Microsoft Scripting Runtime
* Microsoft Visual Basic for Applications Extensibility 5.*(The version hasn't been tested but works with 5.3)

If you don't have them, I suggest installing them as my favorite repos have that as their dependency as well.

Executing the commands for this module happens at the *Intermediate Window*, make sure you know how to run Subs on that window. Run the following script, let the intellisense guide your typing.(Avoid running the procedure/script by keying F5 while the cursor is on the function as this is not the canonical way it was made.)

```
ChipInit.InstallChipFromRepo
```

That will download the whole chip modules in your project. The core module among them is, obviously, *Chip.bas*. You can install or "*chip on*" my favorite tools by running the following script or procedure,

```
Chip.ChipOnFromRepo "Vase"
```

And

```
Chip.ChipOnFromRepo "Wheat"
```

That's pretty much it. After wading through the window outputs, you have those modules installed. This is pretty much the workflow I had in mind. Want a *chip*? Just type in the name and it will download it for you.

chips
====
These are the *chips* available at this time:
* Vase(xlvase) - A unittesting framework 
* Wheat(xlwheat) - A code export/import tool module
* Butil(xlbutil) - A bunch of utility modules

To update the list, run the install procedure to refresh the repo list or you can just import the module <a href="https://raw.githubusercontent.com/FrancisMurillo/xlchip/master/Modules/ChipList.bas">ChipList.bas</a> if you want to be precise. 

```
ChipInit.InstallChipFromRepo
```

Although the list is small, it is a list. If there are more projects, I'll add to that. I do encourage you to turn your collection of VBA modules to a *chip* project so that people can install it easily. 

chip project
====

The only requirement for this manager to read the project as a *chip* project is that it has the module *ChipInfo.bas*. This is where what the manager looks for and runs to get the configurations needed to import the project. This module must have a *Sub WriteInfo()*, as this will be executed by *Application.Run*. The implementation looks like this.

```
Public Sub WriteInfo()
    ChipReadInfo.References = Array( _
        "Microsoft Visual Basic for Applications Extensibility *", _
        "Microsoft Scripting Runtime")
    ChipReadInfo.Modules = Array( _
        "Wheat", "WheatLib", "WheatConfig")
End Sub
```

This is a sample implementation of Wheat for *ChipInfo.bas*. The main idea is that it transfers or sets the values into *ChipReadInfo.bas* so that *Chip* can read it. The limitations of *Application.Run* forced me to this method and it is somewhat akward. If you have a better idea to transfer variables between two projects, drop me a message. Just in case, I did think about an external configuration file such as a JSON or CSV or what-not but I prefer the idea of keeping the files internal to the project itself. Less external file management issues and somewhat clinging to a Python configuration file.

So anyway, there are two things you have to set, both expect an zero-based array of strings:

1. **ChipReadInfo.References** - the boring of the two, you type the references the *chip* needs to work. For *Wheat*, *Vase* and *Chip* they need these modules which is also mentioned in the quick start section. Notice the asteric wildcard character in the first reference, this is so you can match the reference to a similar version of the reference. You can type in for example 5.3 instead of that but if the user has 5.4 installed, it won't detect that as valid; on the flip side it can also match to 5.2, which might nnot be intended. The matching mechanics is dicatated by the **Like** operator, Regular Expressions for VBA is not well supported yet but someone could make a *chip* for that. If these details bore you, you can leave it as a blank array or *Array()* but it's not a good idea to sweep dependencies under the carpet.
2. **ChipReadInfo.Modules** - the interesting of the two, you type the modules you want the project to export. For *Wheat*, they are just the main command module, the library module, and the user configuration module. This is what the user gets when they install your *chip*. 

Be careful when typing, there is no validation or test to make sure the references you entered is valid or the modules you entered is not missing an extra s or i. The template is simple enough, so use that as a guide when filling in the blanks.

If you filled it in correctly, you're project is now *chippable* for others. You can test it out by using the local version of *chip on* using the procedure.

```
Chip.ChipOnLocally 
```

This will open the Browse Dialog. Open your *chip* from another workbook(assuming you set it up to with *Chip* as well), it will read the configuration and hopefully install the modules you want installed. Drop me a message where your is *chip* so I can add it to the list.


chip project
====

The list of important commands for this *Chip*. Majority of the workflow is already discussed above, anything else will require optional parameters or modifying the source code itself; you can read the method headers to get an idea how the routines are executed.

1. **ChipInit.InstallChipFromRepo** - This installs/reinstalls/updates the remaining *Chip* modules by downloading it over the wire, this includes the repo list.
2. **ChipInit.InstallChipLocally** - Same as above except if you have a copy of the RELEASE file, you can install it locally. This is helpful for testing purposes as well as for limited environments, which I work in.
3. **Chip.ChipOnFromRepo** - This installs a *chip* project as dictated in the *ChipList.bas*. You can check out what the names in the list are as well as what URL they point to.
4. **Chip.ChipOnLocally** - Same as the two above, you can install *chip* modules locally by browsing them in your filesystem.
 

what's next
====

There's a lot of limitations of VBA I discover while I add more *chips*. I would avoid using native methods and black magic methods if possible. I would like to keep the interface and logic as plain and expressive as possible as this thing grows. If you have features or problems, drop me a message and let's see if we can make it work.
