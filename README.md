




## Features

- **PPTX extract**: The service breaks down PPTX files into individual slides and those into individual sections.
- **PPTX rebuild**: The second service then puts the sections and attachments back together into the PPTX file.
- **PPTX export-html**: The third service extracts PPTX files into PDF and HTML slides.
- **PPTX htmlpage**: The fourth service converts PPTX files to HTML pages.

## Example of using the /pptx/extract:

```csharp
// pptx/extract
{
  	"screenshotJpegQuality": 90,
  	"directories": [
    "C:\\Repository\\WebPptx\\samples\\"
    ]
}
```

## Example of using the /pptx/rebuild :

```csharp
// pptx/rebuild
{
  	"framesJsonPath": "C:\\Repository\\WebPptx\\samples\\Prezentation01\\screenshots\\frames.json",
    "outputPath": "C:\\Repository\\WebPptx\\samples\\Prezentation01.pptx",
    "overwrite": true,
    "useSlideFallback": true
}
```

## Example of using the /pptx/export-html :

```csharp
// pptx/export-html
{
  	"path": "C:\\Repository\\WebPptx\\samples\\"
}
```

## Example of using the /pptx/htmlpage :

```csharp
// pptx/htmlpage
{
  	"path": "C:\\Repository\\WebPptx\\samples\\"
}
```
