# PDFMerge

PowerScript with Gui elements that prompts user to select RTFs and PDFs.

Selected RTFs files are converted to PDFs using Word (must be installed).

PDFs are then all merged into files broken up as they hit 15MB (requires PSWritePDF PowerShell Module).

The PDF file name are based in timestamp as they are created:

	merged_pdf-yyyy-MM-dd-hh-mm-ss-ms.pdf

