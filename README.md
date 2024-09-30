# VBA Image Converter

This project contains a VBA script and UserForm to convert PowerPoint slides or PDF pages into images and insert them into a Word document at cursor location. The UserForm allows you to choose the size reduction for the inserted images.

## Files

- `ConvertAndInsertSlides.bas`: Contains the main VBA script for image conversion and insertion.
- `frmImageSizeReduction.frm`: The UserForm that allows users to choose the size reduction of the images.
- `frmImageSizeReduction.frx`: Additional resource file for the UserForm.

## How to Use

1. Import the `.bas` file (VBA script) and `.frm` file (UserForm) into your Word VBA project.
2. Run the `ConvertAndInsertSlides` macro to convert and insert images into your Word document.
2. Ensure UserForm options buttons '(NAME)' column following this naming scheme corresponding to the percentile change in image size: 'opt1', 'opt5', 'opt10', 'opt20' etc.

## Requirements

- Microsoft PowerPoint (for converting PowerPoint slides)
- Ghostscript (for converting PDF pages)
