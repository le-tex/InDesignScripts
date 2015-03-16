# image-export.jsx

A script for Adobe InDesign to export images to web-friendly formats.

##  Installation

1. open InDesign Scripts panel
2. right-click on folder entitled "User" and select "Reveal in Explorer/Finder"
3. copy image-export.jsx to this folder

## Requirements

* image-export.jsx needs at least Adobe InDesign Version 8.0 (CS6)
* image-export_pre-cs6.jsx works with Adobe InDesign Version 7.0 (CS5) but supports only JPG output. I've tested it on CS5 but it may work on CS4 as well.

## Features

* adjusted size, bleed, transparency and orientation settings are applied and automatically exported
* for multiple linked images, every instance is stored as separate image with a unique name and its settings
* select output directory, compression ratio, density and choose between JPG and PNG format

## Limitations

* Embedded images are not exported. Consider to convert embedded images to linked images instead.


## IDML

* for multiple instances of one image, each instance is exported with a unique filename
* the new filename is stored in the IDML as ``Label`` of the ``Rectangle`` element

```
<Rectangle>
  <Properties>
    <Label>
      <KeyValuePair Key="letex:fileName" Value="new-filename-a3.jpg" />
    </Label>
  </Properties>
  <!-- (...) -->
</Rectangle>
```
