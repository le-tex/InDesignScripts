# insert-alt-text.jsx

A script for Adobe InDesign to insert alt texts from an XML source.

##  Requirements

You need an XML file with this structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<links>
  <link name="bicycle.tiff" alt="A grey racing cycle with yellow decals leaning against a wall."/>
  <link name="rainbow.tiff" alt="Rainbow shows up over Leipzig."/>
</links>
```

## Description

For each XML link element whose file reference matches an InDesign link reference, the alt text is
inserted into the Object Export Options of the rectangle.

```xml
<ObjectExportOption AltTextSourceType="SourceCustom" CustomAltText="Rainbow shows up over Leipzig.">
  <Properties>
    <AltMetadataProperty NamespacePrefix="$ID/" PropertyPath="$ID/" />
  </Properties>
</ObjectExportOption>
```

Additionally, a `letex:alt-text` label is inserted:

```xml
<Rectangle>
  <Properties>
    <Label>
      <KeyValuePair Key="letex:altText" Value="hier steht der zweite Alternativtext" />
    </Label>
    <!-- (...) -->
  </Properties>
</Rectangle>
```
