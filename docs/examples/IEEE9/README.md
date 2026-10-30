# IEEE9 image annotation example

`IEEE9.bmp` is the approved input. `IEEE9.png` is a lossless display copy with
the same RGB pixels. Only this source and results derived from it are included.

Run from the repository root:

```cmd
annotate-image.cmd "docs\examples\IEEE9\IEEE9.bmp" "%TEMP%\easy-uiauto-ieee9" --profile block-diagram
```

The recorded run uses local YOLOv9-E + RapidOCR PP-OCRv6, confidence 0.1,
the block-diagram profile, square CPU input and local geometry refinement.
It produced 42 graphic observations and 102 text observations with no vision
API requests. The observations are not verified UIA controls or executable XPath
locators, and their count is not a measured accuracy rate.

Open [results/index.html](results/index.html) locally for all three layers,
coordinates and raw diagnostics. Blue `G` labels identify graphics; orange `T`
labels identify text. Dense previews redraw labels at 2x while JSON/CSV coordinates
remain in the recognition input's pixels. `crop.offset` maps back to the original.

Result paths have been rewritten relative to `results/`, including
`../IEEE9.bmp` for the source. Coordinates, predictions and timings are unchanged;
machine-specific paths and unrelated experiments are excluded.
