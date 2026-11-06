## 3.2.1.0
### Fixes
* Gave this build its own visual identity - GUID, name ("Power KPI Plus") and an icon with a "+" badge. It shared the GUID of the Power KPI published on AppSource, so Power BI loaded the AppSource version in its place and none of the features of this build showed up

## 3.2.0.0
### Features
* Added a Tertiary Values field drawn against a third Y axis, placed on the right of the secondary one. It comes with its own Tertiary Y Axis and Tertiary Y Axis Reference Lines cards, offering the same settings as the secondary ones, and both cards only show up once the field is bound
* Added a Display setting to the Y axis, secondary Y axis and tertiary Y axis cards: Standard draws every series from its own value as before, Stacked piles the series of the axis on top of each other and Stacked 100% turns each pile into the shares of its total, with the axis labels read as percentages. Stacking applies to the Line, Area, Scatter and Line and markers types; tooltips, data labels and the KPI keep showing the raw values
* Added a Line and markers option to the Line Type setting, drawing the line with a marker on each data point
### Fixes
* Fixed the type errors the ComboComponent specs raised since the Scatter line type made the radius factor a required render option

## 3.1.5.0
### Features
* Added a Zoom Slider card, off by default, showing a zoom slider on the X axis, the Y axis or both. The slider narrows the axis range the same way the built in visuals do; the chart is rebuilt when the drag ends and the zoom resets when the data is refreshed. The X axis slider is offered only on a date or numeric axis, a categorical one having no range to narrow
### Fixes
* Fixed the chart being painted outside its drawing area, over the Y axis labels and past the edge of the visual, whenever the axis range is narrower than the data - a pinned X axis Min/Max, a zoom range, or the Y axis Min/Max that the visual already offered

## 3.1.4.0
### Features
* Added Min and Max settings to the X axis, pinning the axis range instead of fitting it to the data - numeric fields on a numeric axis, date fields on a date axis

## 3.1.3.0
### Features
* Added a Scatter option to the Line Type setting, rendering each data point as a standalone dot without a connecting line

## 3.1.2.0
### Features
* Added a Label Area Sizing option (Auto / Fixed) to the Y axis and the secondary Y axis to reserve a fixed pixel width for axis labels, so the plot area keeps the same position when filters change the magnitude of the displayed values

## 3.1.1.0
### Fixes
* Fixed other lines being recolored when changing a single line's color (default palette colors are now reserved per line regardless of overrides)
* Fixed X-axis tick labels overlapping when the visual is narrowed and the axis font size is increased - the axis now redistributes to fewer, wider-spaced ticks that fit their available space, falling back to per-label truncation only when a single tick still can't fit
* Fixed visual not activating when clicking on chart lines or empty SVG canvas area

## 3.1.0.0
### Features
* Added a line color mode (Joint / Granular) to configure colors either per group or per individual line
### Fixes
* Fixed series colors being shared between measures of the same group

## 3.0.4.0
### Fixes
* Fixed format value

## 3.0.3.0
* Fixed data point starts kpi color segment

## 3.0.2.0
* Fixed toggle auto values
* Added locale to y-axis labels formatting
* API 5.11.0

## 3.0.1.0
* Added high contrast mode
* Added context menu
* Migrated to formatting model

## 3.0.0
* Updated packages
* API 5.3.0
* Migrated to format pane
* Removed `jquerry`

## 2.0.1
* Conditional loading of `core-js/stable` only for sandbox mode
* API 2.5.0
* `@babel/polyfill` replacement by `core-js/stable`

## 2.0.0
* API 2.2.0
* Webpack based PBIVIZ

## 1.8.0
* Bidirectional color segmentation
