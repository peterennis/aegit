# Quickstart

This is assuming you have already setup a GitHub account.

Minimum requirement is Microsoft Access 2010.

1. Fork this repo https://github.com/peterennis/aegit_Template
2. Clone it
3. In the project folder rename the file **aegit_Template.accd** to **aegit_Template.accdb** and run the file **aegit_Template.accdb**
	1. The file **.gitignore** is set to ignore **accdb** files so this work around allows the template file to be available.
4. Open the module **aebasTEST_aegit_expClass** in the VBA editor
5. Run the command **aegit_Template_EXPORT**
7. Run Compact and Repair (Make a backup copy of the dbs: Ctr-C, Ctrl-V) 
8. Review the output in **.\src** and **.\src\xml**
9. Commit the changes to GitHub
10. Edit, Export, Compact (Backup), Commit, Rinse & Repeat

To use in your own database import module **aebasTEST_aegit_expClass** and
the class module **aegit_expClass**. Adjust the public constants that define the path to your project as appropriate. The default values will export data in the project folder that was selected when cloning.

**NOTE:** The class is configured to EXCLUDE the aegit files from export.
  
**NOTE:** Your initial commit of changes will look similar to this
[Second Commit](https://github.com/peterennis/aegit_Template/commit/dd71ff55a71372abcbeadec7ce250f09ff4ad6bd)

If you look at **src/frm_Dummy.frm.txt** there is binary encoding for ImageData
and that is a result of the "Name AutoCorrect Option" setting of Access. So get
rid of it, *then may you live long and prosper!*

*File* -> *Options* -> *Current Database* -> *Name AutoCorrect Options*

>Uncheck: Track name AutoCorrect info

>Uncheck: Perform name AutoCorrect

