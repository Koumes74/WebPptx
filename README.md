




## Features

- **PPTX extract**: The service breaks down PPTX files into individual slides and those into individual sections.
- **PPTX rebuild**: The second service then puts the sections and attachments back together into the PPTX file.

## Example of using the PPTX file layout:

```csharp
// pptx/extract
{
  	"screenshotJpegQuality": 90,
  	"directories": [
    "C:\\Repository\\WebPptx\\samples\\"
    ]
}
```

## Example of using PPTX file composition:

```csharp
// pptx/rebuild
{
  	"framesJsonPath": "C:\\Repository\\WebPptx\\samples\\Prezentation01\\screenshots\\frames.json",
    "outputPath": "C:\\Repository\\WebPptx\\samples\\Prezentation01.pptx",
    "overwrite": true,
    "useSlideFallback": true
}
```

