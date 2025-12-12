# VBA Progress Bar Class

A comprehensive VBA class for displaying a customizable progress bar during long-running operations in Excel or other Office applications.

## Overview

The `clsProgresBar` class provides an easy-to-use interface for displaying progress during long-running operations. It features a clean, customizable interface with support for dual messages, visual indicators, and user cancellation.

## Features

- **Dual Message Display**: Top and bottom message areas for detailed progress information
- **Visual Progress Indicator**: Animated progress line with customizable colors
- **User Cancellation**: Support for cancelling operations with the Esc key
- **Customizable Appearance**: Configurable symbols, colors, and dimensions
- **Performance Optimized**: Efficient updates to minimize UI lag
- **Error Handling**: Comprehensive error management and reporting

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
    oProg.Initialize "Processing data...", "Data Processing", True, ">"
    
    ' Simulate a long-running operation
    Dim i As Long
    For i = 1 To 1000
        ' Update progress bar
        oProg.Update i, 1000, "Processing item " & i
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
    oProg.Initialize "Loading files...", "File Loader", True, "*"
    
    ' Customize colors
    oProg.lProgressColor = RGB(0, 128, 255) ' Blue progress line
    oProg.lPictColor = RGB(255, 165, 0)     ' Orange symbol
    
    ' Resize the progress bar
    oProg.Resize 600, 40, 30
    
    ' Process with progress updates
    Dim i As Long
    For i = 1 To 500
        oProg.Update i, 500, "File " & i & " of 500"
        
        ' Add delay to simulate work
        Application.Wait Now + TimeValue("0:00:01")
    Next i
    
    Set oProg = Nothing
End Sub
```

## API Reference

### Methods

#### `Initialize(sMessage As String, sHeader As String, Optional bShowPict As Boolean = False, Optional sPict As String = vbNullString)`
Initializes the progress bar with specified parameters.

- `sMessage`: Message to display in the top message area
- `sHeader`: Header text for the form
- `bShowPict`: Whether to show the symbol/picture indicator
- `sPict`: Symbol/picture to display

#### `Update(i As Long, iCount As Long, sMsg As String)`
Updates the progress bar with new values.

- `i`: Current step number
- `iCount`: Total number of steps
- `sMsg`: Message to display in the bottom message area

#### `Resize(Width As Double, HeightMessage As Double, HeightMessageTwo As Double)`
Changes the size of the progress bar.

- `Width`: Form width
- `HeightMessage`: Top message height
- `HeightMessageTwo`: Bottom message height

### Properties

#### `Header As String`
Gets or sets the form header text.

#### `MessageTop As String`
Gets or sets the top message text.

#### `MessageBottom As String`
Gets or sets the bottom message text.

#### `ShowPict As Boolean`
Gets or sets the visibility of the symbol/picture indicator.

#### `sPict As String`
Gets or sets the symbol/picture text.

#### `lPictColor As Long`
Gets or sets the color of the symbol/picture indicator.

#### `lProgressColor As Long`
Gets or sets the color of the progress line.

## Error Handling

The class includes comprehensive error handling:

- Division by zero in the Update method raises error 1001
- Invalid parameters in the Resize method raise error 1002

## Customization

The progress bar can be customized in several ways:

- Colors for progress line and indicator symbol
- Size and dimensions
- Visibility of the indicator symbol
- Header and message text
- Progress indicator symbol

## Performance

The implementation is optimized for performance by:

- Only updating UI elements when values change
- Efficient string operations
- Minimal use of expensive operations during updates

## License

Apache License 2.0