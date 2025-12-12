# VBA Progress Bar Class

A comprehensive VBA class for displaying a customizable progress bar during long-running operations in Excel or other Office applications.

![Demo project](User_Forms.gif)

## Overview

The `clsProgresBar` class provides an easy-to-use interface for displaying progress during long-running operations. It features a clean, customizable interface with support for dual messages, visual indicators, and user cancellation.

## Features

- **Dual Message Display**: Top and bottom message areas for detailed progress information
- **Visual Progress Indicator**: Animated progress line with customizable colors
- **User Cancellation**: Support for cancelling operations with the Esc key
- **Customizable Appearance**: Configurable symbols, colors, and dimensions
- **Performance Optimized**: Efficient updates to minimize UI lag
- **Error Handling**: Comprehensive error management and reporting
- **Time Tracking**: Built-in timer to track operation duration
- **Logging Capability**: Automatic logging of progress events

## Installation

1. Import the class module `clsProgresBar.cls` into your VBA project
2. Import the form module `frmProgresBar_2.frm` into your VBA project
3. Use the class in your code as shown in the usage examples

## Usage

### Basic Usage

```vba
Sub ExampleUsage()
    Dim oProg As clsProgresBar
    Set oProg = New clsProgresBar
    
    ' Initialize the progress bar
    oProg.Initialize "Processing data...", "Data Processing", "Initializing...", _
                     enumTypeCaptionLabel.enProcent, 100
    
    ' Simulate a long-running operation
    Dim i As Long
    For i = 1 To 100
        ' Update progress bar
        oProg.Update i / 10, "Processing item " & i
    Next i
    
    Set oProg = Nothing
End Sub
```

### Advanced Usage

```vba
Sub AdvancedExample()
    Dim oProg As clsProgresBar
    Set oProg = New clsProgresBar
    
    ' Initialize with custom settings
    oProg.Initialize "Loading files...", "File Loader", "Starting up...", _
                     enumTypeCaptionLabel.enAll, 500, RGB(0, 128, 255), RGB(20, 200, 200), "|", True, "*"
    
    ' Resize the progress bar
    oProg.Resize 600, 40, 30, 20
    
    ' Process with progress updates
    Dim i As Long
    For i = 1 To 500
        If oProg.Update(i / 500, "File " & i & " of 500") Then
            ' User requested cancellation
            Exit For
        End If
        
        ' Add delay to simulate work
        Application.Wait Now + TimeValue("0:00:01")
    Next i
    
    ' Access log data if needed
    Dim logData As Variant
    logData = oProg.LogData
    
    Set oProg = Nothing
End Sub
```

## API Reference

### Enums

#### `enumTypeCaptionLabel`
Defines the type of information displayed on the progress bar:
- `enNone`: No information displayed
- `enProcent`: Percentage only
- `enValue`: Current/Total values
- `enTime`: Elapsed time
- `enProcentValue`: Percentage and value
- `enProcentTime`: Percentage and elapsed time
- `enValuTime`: Value and elapsed time
- `enAll`: All information types

#### `enumParametrVersion`
Provides access to version information:
- `enName`: Class name
- `enAuthor`: Author name
- `enVersion`: Version string
- `enLicense`: License information
- `enDateOfCreation`: Creation date
- `enDateOfUpdate`: Last update date
- `enDescription`: Description text
- `enAll`: All version information

### Methods

#### `Initialize(sHeaderCaption As String, sMessageTop As String, sMessageBottom As String, TypeCaptionLabel As enumTypeCaptionLabel, Optional CountItems As Long = 0, Optional LineFrontColor As XlRgbColor = -1, Optional LineBackColor As XlRgbColor = -1, Optional sLineFrontSimvol As String = "|", Optional bPictureShow As Boolean = False, Optional sPictureSimvol As String = vbNullString)`
Initializes the progress bar with specified parameters.

- `sHeaderCaption`: Text for the form caption
- `sMessageTop`: Message to display in the top message area
- `sMessageBottom`: Message to display in the bottom message area
- `TypeCaptionLabel`: Type of information to display
- `CountItems`: Total number of items for progress calculation
- `LineFrontColor`: Color of the front progress line
- `LineBackColor`: Color of the back progress line
- `sLineFrontSimvol`: Character to use for the progress line
- `bPictureShow`: Whether to show the progress indicator symbol
- `sPictureSimvol`: Symbol to display at the progress position

#### `Update(procent As Single, sMessageBottom As String, Optional CountItems As Long = 0, Optional LineFrontColor As XlRgbColor = -1, Optional LineBackColor As XlRgbColor = -1) As Boolean`
Updates the progress bar with new values and returns True if user requested cancellation.

- `procent`: Progress percentage (0.0 to 1.0)
- `sMessageBottom`: Message to display in the bottom message area
- `CountItems`: Total number of items (optional)
- `LineFrontColor`: New color for front progress line (optional)
- `LineBackColor`: New color for back progress line (optional)

#### `Resize(WidthForm As Single, HeightMessageTop As Single, HeightMessageBottom As Single, HeightLineProgres As Single)`
Changes the size of the progress bar.

- `WidthForm`: Form width
- `HeightMessageTop`: Top message height
- `HeightMessageBottom`: Bottom message height
- `HeightLineProgres`: Progress line height

### Properties

#### `HeaderCaption As String`
Gets or sets the form header text.

#### `MessageTop As String`
Gets or sets the top message text.

#### `MessageBottom As String`
Gets or sets the bottom message text.

#### `PictureShow As Boolean`
Gets or sets the visibility of the progress indicator symbol.

#### `PictureSimvol As String`
Gets or sets the progress indicator symbol text.

#### `PictureColor As XlRgbColor`
Gets or sets the color of the progress indicator symbol.

#### `LineFrontColor As XlRgbColor`
Gets or sets the color of the progress line.

#### `LineBackColor As XlRgbColor`
Gets or sets the color of the progress line background.

#### `TimeWork As Date`
Gets the elapsed time since initialization.

#### `TypeCaptionLabel As enumTypeCaptionLabel`
Gets or sets the type of caption label.

#### `LogData As Variant`
Gets the log data array containing progress events.

## Error Handling

The class handles potential errors gracefully:
- Invalid percentage values are normalized to the 0-1 range
- Division by zero is prevented in calculations
- Proper cleanup occurs in the Class_Terminate event

## Customization

The progress bar can be customized in several ways:

- Colors for progress line and indicator symbol
- Size and dimensions
- Visibility of the indicator symbol
- Header and message text
- Progress indicator symbol
- Type of information displayed (percentage, values, time, etc.)

## Performance

The implementation is optimized for performance by:

- Efficient UI updates
- Proper memory management
- Minimizing expensive operations during updates
- Using DoEvents to maintain UI responsiveness

## License

Apache License 2.0