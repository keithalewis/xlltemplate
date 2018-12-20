# xlltemplate

Template for Excel add-ins

Fork or clone the repository. Set the `Configuration` (Debug or Release) and `Platform` (x86 or x64) and build the `template` project.  The platform must be the same as Excel: x86 for 32-bit or x64 for 64-bit.

## Debugging

Set the configuration to `Debug|x86` if using 32-bit Excel or `Debug|x64` for 64-bit Excel.
Set breakpoints by clicking on the left boarder at the line you want to stop at, the hit `F5` to start debugging.

## Documentation

Instanciate a `AddIn xai_`_anyname_`(Documentation(LR("Documentation goes here")));` object to create a 
[Sandcastle Helpfile Builder](https://github.com/EWSoftware/SHFB) project file in the
Debug folder of your project. Put the `.chm` file it creates in the same location as the `.xll` add-in to
integrate with Excel's function wizard <ins>Help on this function</ins>.