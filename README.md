# VBA-Add-in-Project-Export-Import
Export VBA code from projects (and reimport)

A simple VBA IDE add-in to allow the export of a complete project, either as a whole, excluding forms or an individual component. It also exports the list of references. Forms can be excluded as the .frx file is changed event if no changes to the forms have been made, so if version controlling and you know no forms have been changed they can be left out of the export.

It also allows a reimport of a project, deleting the existing project. References are not imported.

