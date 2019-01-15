# How to compile the FreeBasic compiler

Download the FreeBasic compiler source from GitHub https://github.com/freebasic/fbc and extract it to your hard drive.

Start WinFBE and create a new WinFBE Project (top menu Project / New Project)

Project Path - Navigate to the folder where you extracted the FB source code. Select the folder where the compiler source code is located (eg. \fbc-master\src\compiler) and name the project file "FBCompiler.wfbe".

Other Options 32-bit & Other Options 64-bit - for each of these additional options we want to add the directive "-d ENABLE_STANDALONE" (without the double quotes)

You can leave the checkbox "Create resource file and manifest" checked.

WinFBE will create a blank, unsaved, source file called "Untitled1" by default. You can simply close that file because we will not be creating any new new files. We will be loading existing source files. Click on the red "X" in the tab next to the "Untitled1" name to close it.

Load the FB compile source files into the project. It is important to use the "Project / Add Files to Project" menu option rather than the "Files / Open Files" option because using the project option will load the files much faster because they are not opened in the top tab control. They will simply be parsed and added  to the project. Select all of the files in the "\fbc-master\src\compiler" folder (except for the .wfbe subfolder).

Once all of the files are loaded, open the Explorer treeview branch called "Module" and double click on the file "fbc.bas". This will load the file into the editor for editing.

We need to designate the "fbc.bas" file as the "Main" file for the project. This can be done in several ways. All of the following achieve the same result:

- Click on the statusbar panel that should now read "MODULE" and select "Main file" from the popup menu. 
- Right-click menu in the editor and select "Main file".
- Right-click menu from the tab control tab and select "Main file".
- Right-click menu from the Explorer treeview node for "fbc.bas" and select "Main file".
- In the Explorer treeview, drag and drop the "fbc.bas" node from the Module branch to the Main branch.

Compile the EXE. Select "Compile / Compile" (Ctrl+F5) from the top menu. It will take a while for all of the individual object modules to be compiled and linked together to create the FBC.EXE compiler.

You should be able to find the newly created "fbc.exe" executable in the "\fbc-master\src\compiler" folder.

Subsequent compiles of the source code should be lightning fast because WinFBE will only recompile object files whose main source code has changed. If you only change a few source files then only those files will have to be recompiled. If needed, you can always initiate a full recompile by selecting "Compile / Rebuild All" (Ctrl+Alt+F5).

Note: The process above only compiles the FB compile source code. It does not compile the FB runtime library or the gfxlib2 library.

