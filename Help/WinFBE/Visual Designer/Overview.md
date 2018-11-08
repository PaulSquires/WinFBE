# Overview
The visual designer is still a major work in progress so please do not assume that all functionality is available and/or working correctly. The purpose of this release to show the progress of the visual designer to date and to allow input from users regarding its direction and whether changes in basic design philosophy should occur at this point.

NOTE: Code may fail, GPF’s may occur in the editor. Do not rely on this version of the designer for any production quality code. Treat this mainly as a demo.

The following areas of the visual designer are not working:

- Menu Editor
- Toolbar Editor
- Statusbar Editor
- Images/Pictures/Icons. Any control property that relies on these types has not been implemented yet. Need to create an interface between the editor and resource file.
- Guidelines and snap to lines on the design form

There are a few controls that have been implemented. More to come as the core structure of the designer itself matures.

- Forms
- Label
- Button (Jose Roca’s CXpButton class)
- MaskedEdit (Jose Roca’s CMaskedEdit class) 
- TextBox
- CheckBox
- Frame
- ListBox
- ComboBox 
- OptionButton
- Frame
- StatusBar (only through code at this time)

The forms and controls are being implemented via the WinFormsX library that I am writing and maintaining on GitHub.

The rest of this document will give you a bare bones introduction to the use of the visual designer and code to perform various tasks.

The visual designer will automatically generate code whenever something changes related to a form or control. The generated code is based on the syntax of the **WinFormsX library**.

When the code is first generated, the visual designer will output an additional line:

```
Application.Run(Form1)
```

That line essentially bootstraps the form so that it will display. Obviously you don’t need a line like this for every form that is generated. You can safely move that line to a main source file so that it only runs once at program startup. Any subsequent similar generated lines for other forms can be safely deleted.

WinFBE generates a **TYPE** structure for every form in the application. That TYPE extends a bass class called **wfxForm** which contains the core functionality of a form. You can find the wfxForm source code and all other source code for other controls such as **wfxLabel, wfxTextBox** in the WinFormsX include folder. A global shared variable is generated for each form TYPE. For example, if your form is named Form1 then the shared variable is also named Form1.

```
Dim Shared Form1 as Form1Type
```

If you have controls on your form then you access them via dot syntax. For example, if you have a button called Button1 located on the form called Form1 then you access the button via **Form1.Button1**

```
Form1.Button1.Text = "Cancel"
```

