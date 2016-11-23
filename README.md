# AddExistingItemInVirtualFolder

<!-- Replace this badge with your own-->
[![Build status](https://ci.appveyor.com/api/projects/status/u3u89ivwx4urari3?svg=true)](https://ci.appveyor.com/project/sandercox/AddExistingItemInVirtualFolder)

<!-- Update the VS Gallery link after you upload the VSIX-->
Get the [CI build](http://vsixgallery.com/extension/AddExistingItemInVirtualFolder.SanderCox.6eb67602-1ce0-4a6f-82d4-9d6ccc2b6d72/).

---------------------------------------

Hijack the Add Existing Item dialog in Visual Studio solutions for "Filter folders" (commonly used in Visual C++ projects).

If we can determine a working directory for the files based on files around this virtual folder, present an open file dialog at 
this location instead of add the default project root.

See the [change log](CHANGELOG.md) for changes and road map.

## Features

- Overwrite Add Existing File dialog for filter folders
- Leave actions alone when not applied to filter folders.

## Contribute
For cloning and building this project yourself, make sure
to install the
[Extensibility Tools 2015](https://visualstudiogallery.msdn.microsoft.com/ab39a092-1343-46e2-b0f1-6a3f91155aa6)
extension for Visual Studio which enables some features
used by this project.

## License
[Apache 2.0](LICENSE)