# easy-uiauto GUI icons

50 SVG files represent 49 semantic icons. `toggle-on.svg` and `toggle-off.svg`
are the two states of the `toggle` concept. The editable source is each SVG;
`manifest.json` records its Chinese label, group, palette color, and design number.

All assets use a 24 × 24 viewBox, a 1.75-unit rounded stroke, at least two units of
padding, and a transparent background. They contain vector geometry only, with
no fonts, external files, raster images, filters, or scripts.

## Color and sizes

The root `color` attribute provides the default semantic color. Shapes inherit
it using `currentColor`, including solid dots and the play/stop shapes. Change
only this attribute to recolor the icon; keep `fill="none"` on outlined shapes.

- Normal: the color specified in the manifest.
- Disabled: `#9CA3AF`.
- Selected on a dark background: `#FFFFFF`.
- Navigation/menu: 16 px. Buttons: 18 px. Main entry points: 20 px.
- SVG can be rendered directly at higher DPI; no separate PNG is required.

The set has been checked with Qt's `QSvgRenderer` at 16, 20, 24, 32, and 48 px
in all three color states. `status-busy.svg` is a static arc; animation belongs
to the eventual GUI integration.

## Preview and verification

From the repository root, with the GUI dependencies installed:

```console
python scripts/preview_gui_icons.py
```

This validates SVG geometry, palette inheritance, clipping, and transparent
padding, then regenerates:

- `docs/gui-icons-preview.html`: self-contained preview with search, categories,
  size/state controls, and individual SVG downloads.
- `docs/assets/gui-icons/contact-sheet.png`: complete visual overview.
- `docs/assets/gui-icons/validation.json`: rendering check summary.

The detailed usage mapping and design brief are in
`docs/gui-icons-design.zh-CN.md`. GUI integration lives in `qt_ui/icons.py`:
the shared icon engine renders SVG at the requested device pixel ratio and
provides disabled/selected variants. Buttons, navigation and status labels
use the same assets.
