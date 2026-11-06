/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

import powerbi from "powerbi-visuals-api";
import { select as d3Select, Selection } from "d3-selection";
import { createColorPalette, testDom } from "powerbi-visuals-utils-testutils";
import { IMargin } from "powerbi-visuals-utils-svgutils";

import { IDataRepresentation } from "../src/dataRepresentation/dataRepresentation";
import { IDataRepresentationAxis } from "../src/dataRepresentation/dataRepresentationAxis";
import {
    getPlottedBaseline,
    getPlottedY,
    IDataRepresentationPoint,
} from "../src/dataRepresentation/dataRepresentationPoint";
import { DataRepresentationScale } from "../src/dataRepresentation/dataRepresentationScale";
import { IDataRepresentationSeries } from "../src/dataRepresentation/dataRepresentationSeries";
import { DataRepresentationTypeEnum } from "../src/dataRepresentation/dataRepresentationType";
import { YAxisDisplayMode } from "../src/settings/descriptors/axis/yAxisDescriptor";
import { DisplayUnitsType } from "../src/settings/descriptors/numberDescriptorBase";
import { LineInterpolation, LineStyle, LineType } from "../src/settings/descriptors/line/lineTypes";
import { Settings } from "../src/settings/settings";
import { ComboComponent } from "../src/visualComponent/combo/comboComponent";
import { IYAxisComponentRenderOptions, YAxisComponent } from "../src/visualComponent/axes/yAxisComponent";

import {
    buildValueRolesDataView,
    convert,
    IValueColumn,
    setDisplayMode,
} from "./valueRolesDataBuilder";

const twoMeasures: IValueColumn[] = [
    { role: "Values", name: "Sales", values: [10, 20, 30] },
    { role: "Values", name: "Cost", values: [1, 2, 3] },
];

function getBands(series: IDataRepresentationSeries): Array<[number, number]> {
    return series.points.map((point: IDataRepresentationPoint) => [point.stackedY0, point.stackedY1]);
}

function getRawValues(series: IDataRepresentationSeries): number[] {
    return series.points.map((point: IDataRepresentationPoint) => point.y);
}

function convertWithMode(
    columns: IValueColumn[],
    mode: YAxisDisplayMode,
    configureSettings?: (settings: Settings) => void,
): IDataRepresentation {
    return convert(buildValueRolesDataView(columns), (settings: Settings) => {
        setDisplayMode(settings.yAxis, mode);

        if (configureSettings) {
            configureSettings(settings);
        }
    });
}

describe("Y axis display mode", () => {
    describe("Standard", () => {
        it("should be the default and leave every point unstacked", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(twoMeasures));

            expect(new Settings().yAxis.getDisplayMode()).toBe(YAxisDisplayMode.standard);

            dataRepresentation.groups[0].series.forEach((series: IDataRepresentationSeries) => {
                series.points.forEach((point: IDataRepresentationPoint) => {
                    expect(point.stackedY0).toBeUndefined();
                    expect(point.stackedY1).toBeUndefined();
                    expect(getPlottedY(point)).toBe(point.y);
                    expect(getPlottedBaseline(point)).toBeUndefined();
                });
            });
        });

        it("should keep the axis fitted to the raw values", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(twoMeasures));

            expect(dataRepresentation.groups[0].y.min).toBe(1);
            expect(dataRepresentation.groups[0].y.max).toBe(30);
        });
    });

    describe("Stacked", () => {
        it("should pile every series on top of the ones bound before it", () => {
            const [sales, cost] = convertWithMode(twoMeasures, YAxisDisplayMode.stacked).groups[0].series;

            expect(getBands(sales)).toEqual([[0, 10], [0, 20], [0, 30]]);
            expect(getBands(cost)).toEqual([[10, 11], [20, 22], [30, 33]]);
        });

        it("should keep the raw values for the tooltips, the labels and the KPI", () => {
            const dataRepresentation: IDataRepresentation = convertWithMode(twoMeasures, YAxisDisplayMode.stacked);
            const [sales, cost] = dataRepresentation.groups[0].series;

            expect(getRawValues(sales)).toEqual([10, 20, 30]);
            expect(getRawValues(cost)).toEqual([1, 2, 3]);
            expect(cost.current.y).toBe(3);
            expect(dataRepresentation.series[0].current.y).toBe(30);
        });

        it("should draw each point at the top of its band", () => {
            const cost: IDataRepresentationSeries = convertWithMode(twoMeasures, YAxisDisplayMode.stacked)
                .groups[0].series[1];

            expect(cost.points.map(getPlottedY)).toEqual([11, 22, 33]);
            expect(cost.points.map(getPlottedBaseline)).toEqual([10, 20, 30]);
        });

        it("should fit the axis to the totals and include the zero baseline", () => {
            const dataRepresentation: IDataRepresentation = convertWithMode(twoMeasures, YAxisDisplayMode.stacked);

            expect(dataRepresentation.groups[0].y.min).toBe(0);
            expect(dataRepresentation.groups[0].y.max).toBe(33);
            expect(dataRepresentation.groups[0].y.scale.getDomain()).toEqual([0, 33]);
        });

        it("should pile negative values down from zero, apart from the positive ones", () => {
            const [first, second] = convertWithMode([
                { role: "Values", name: "First", values: [10, -5] },
                { role: "Values", name: "Second", values: [-4, -6] },
            ], YAxisDisplayMode.stacked).groups[0].series;

            expect(getBands(first)).toEqual([[0, 10], [0, -5]]);
            expect(getBands(second)).toEqual([[0, -4], [-5, -11]]);
        });

        it("should give no room in the pile to a point without a value", () => {
            const [first, second] = convertWithMode([
                { role: "Values", name: "First", values: [10, null, 30] },
                { role: "Values", name: "Second", values: [1, 2, 3] },
            ], YAxisDisplayMode.stacked).groups[0].series;

            expect(first.points[1].y).toBeNull();
            expect(getBands(second)).toEqual([[10, 11], [0, 2], [30, 33]]);
        });

        it("should still let the author pin the axis range", () => {
            const dataRepresentation: IDataRepresentation = convertWithMode(
                twoMeasures,
                YAxisDisplayMode.stacked,
                (settings: Settings) => {
                    settings.yAxis.max.value = 50;
                },
            );

            expect(dataRepresentation.groups[0].y.max).toBe(50);
        });

        it("should only stack the axis it is set on", () => {
            const dataRepresentation: IDataRepresentation = convertWithMode([
                ...twoMeasures,
                { role: "SecondaryValues", name: "Margin", values: [5, 6, 7] },
                { role: "SecondaryValues", name: "Rate", values: [1, 1, 1] },
            ], YAxisDisplayMode.stacked);

            expect(dataRepresentation.groups[0].series[1].points[0].stackedY1).toBe(11);

            dataRepresentation.groups[1].series.forEach((series: IDataRepresentationSeries) => {
                expect(series.points[0].stackedY1).toBeUndefined();
            });

            expect(dataRepresentation.groups[1].y.max).toBe(7);
        });

        it("should stack the secondary and tertiary axes on their own", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView([
                ...twoMeasures,
                { role: "SecondaryValues", name: "Margin", values: [5, 6, 7] },
                { role: "SecondaryValues", name: "Rate", values: [1, 1, 1] },
                { role: "TertiaryValues", name: "Units", values: [100, 200, 300] },
                { role: "TertiaryValues", name: "Returns", values: [10, 20, 30] },
            ]), (settings: Settings) => {
                setDisplayMode(settings.secondaryYAxis, YAxisDisplayMode.stacked);
                setDisplayMode(settings.tertiaryYAxis, YAxisDisplayMode.stacked);
            });

            expect(dataRepresentation.groups[0].series[1].points[0].stackedY1).toBeUndefined();
            expect(getBands(dataRepresentation.groups[1].series[1])).toEqual([[5, 6], [6, 7], [7, 8]]);
            expect(getBands(dataRepresentation.groups[2].series[1])).toEqual([[100, 110], [200, 220], [300, 330]]);
            expect(dataRepresentation.groups[2].y.max).toBe(330);
        });
    });

    describe("Stacked 100%", () => {
        it("should turn every pile into the shares of its total", () => {
            const [first, second] = convertWithMode([
                { role: "Values", name: "First", values: [1, 3, 0] },
                { role: "Values", name: "Second", values: [3, 1, 5] },
            ], YAxisDisplayMode.stacked100).groups[0].series;

            expect(getBands(first)).toEqual([[0, 0.25], [0, 0.75], [0, 0]]);
            expect(getBands(second)).toEqual([[0.25, 1], [0.75, 1], [0, 1]]);
        });

        it("should keep the raw values", () => {
            const [first] = convertWithMode(twoMeasures, YAxisDisplayMode.stacked100).groups[0].series;

            expect(getRawValues(first)).toEqual([10, 20, 30]);
        });

        it("should run the axis from 0 to 1 and format it as a percentage", () => {
            const axis: IDataRepresentationAxis = convertWithMode(twoMeasures, YAxisDisplayMode.stacked100).groups[0].y;

            expect(axis.min).toBe(0);
            expect(axis.max).toBeCloseTo(1, 10);
            expect(axis.format).toBe("0%");
            expect(axis.isPercentage).toBe(true);
        });

        it("should not divide by zero on a category whose values are all empty or zero", () => {
            const [first, second] = convertWithMode([
                { role: "Values", name: "First", values: [0, null] },
                { role: "Values", name: "Second", values: [0, null] },
            ], YAxisDisplayMode.stacked100).groups[0].series;

            [...getBands(first), ...getBands(second)].forEach(([baseline, top]) => {
                expect(baseline).toBe(0);
                expect(top).toBe(0);
            });
        });
    });

    describe("rendering", () => {
        const viewport: powerbi.IViewport = { height: 300, width: 400 };

        function renderSeries(series: IDataRepresentationSeries, lineType: LineType, axis: IDataRepresentationAxis): Element {
            const rootElement: HTMLElement = testDom(viewport.height.toString(), viewport.width.toString());
            const element: Selection<HTMLElement, any, any, any> = d3Select(rootElement);

            new ComboComponent({ element }).render({
                areaOpacity: 0.5,
                colorPalette: createColorPalette(),
                gradientPoints: series.gradientPoints,
                interpolation: LineInterpolation.linear,
                lineStyle: LineStyle.solidLine,
                lineType,
                opacity: 1,
                radiusFactor: 1.4,
                series,
                thickness: 2,
                viewport,
                x: DataRepresentationScale
                    .create()
                    .domain([0, series.points.length - 1], DataRepresentationTypeEnum.NumberType),
                y: axis.scale,
            });

            return rootElement;
        }

        /**
         * Reads the Y coordinates of a path built with linear segments only (M/L commands).
         */
        function getPathYCoordinates(path: Element): number[] {
            return (path.getAttribute("d").match(/-?\d+(\.\d+)?,-?\d+(\.\d+)?/g) || [])
                .map((pair: string) => parseFloat(pair.split(",")[1]));
        }

        function getStackedSeries(): { series: IDataRepresentationSeries; axis: IDataRepresentationAxis } {
            const dataRepresentation: IDataRepresentation = convertWithMode([
                { role: "Values", name: "Bottom", values: [10, 10, 10] },
                { role: "Values", name: "Top", values: [10, 10, 10] },
            ], YAxisDisplayMode.stacked);

            // Points are re-indexed on the X axis so the spec can use a numeric scale
            const series: IDataRepresentationSeries = dataRepresentation.groups[0].series[1];
            series.points.forEach((point: IDataRepresentationPoint, index: number) => {
                point.x = index;
            });

            return { series, axis: dataRepresentation.groups[0].y };
        }

        it("should draw the line of a stacked series along the top of its band", () => {
            const { series, axis } = getStackedSeries();
            const rootElement: Element = renderSeries(series, LineType.line, axis);

            // The domain is [0, 20] and the band of the top series is [10, 20]: its line
            // runs along the top edge of the chart
            getPathYCoordinates(rootElement.querySelector("path")).forEach((y: number) => {
                expect(y).toBeCloseTo(0, 5);
            });
        });

        it("should fill a stacked area down to the series below it instead of the bottom", () => {
            const { series, axis } = getStackedSeries();
            const rootElement: Element = renderSeries(series, LineType.area, axis);
            const areaPath: Element = rootElement.querySelector(".powerKpi_areaComponent_area");

            const yCoordinates: number[] = getPathYCoordinates(areaPath);

            expect(yCoordinates.length).toBeGreaterThan(0);
            // From the top edge down to the middle of the chart, never to its bottom
            expect(Math.min(...yCoordinates)).toBeCloseTo(0, 5);
            expect(Math.max(...yCoordinates)).toBeCloseTo(viewport.height / 2, 5);
        });

        it("should place the markers on the top of the band", () => {
            const { series, axis } = getStackedSeries();
            const rootElement: Element = renderSeries(series, LineType.scatter, axis);

            Array.from(rootElement.querySelectorAll("circle")).forEach((circle: Element) => {
                expect(parseFloat(circle.getAttribute("cy"))).toBeCloseTo(0, 5);
            });
        });
    });

    describe("Stacked 100% axis labels", () => {
        function renderAxisLabels(configureSettings?: (settings: Settings) => void): string[] {
            const axis: IDataRepresentationAxis = convertWithMode(twoMeasures, YAxisDisplayMode.stacked100).groups[0].y;
            const settings: Settings = new Settings();

            if (configureSettings) {
                configureSettings(settings);
            }

            const viewport: powerbi.IViewport = { height: 300, width: 400 };
            const rootElement: HTMLElement = testDom(viewport.height.toString(), viewport.width.toString());
            const margin: IMargin = { bottom: 0, left: 0, right: 0, top: 0 };
            const component: YAxisComponent = new YAxisComponent({ element: d3Select(rootElement) });

            const options: IYAxisComponentRenderOptions = {
                axis,
                colorPalette: createColorPalette(),
                locale: "en-US",
                margin,
                settings: settings.yAxis,
                viewport,
            };

            component.preRender(options);
            component.render(options);

            return Array.from(rootElement.querySelectorAll(".tick text"))
                .map((element: Element) => element.textContent);
        }

        // At the default density the axis leaves out the tick of its maximum, as every
        // Y axis of the visual does, so the labels are checked for their format only
        function expectPercentages(labels: string[]): void {
            expect(labels.length).toBeGreaterThan(1);
            expect(labels).toContain("0%");
            expect(labels).toContain("50%");

            labels.forEach((label: string) => {
                expect(label).toMatch(/^\d+%$/);
            });
        }

        it("should read as percentages", () => {
            expectPercentages(renderAxisLabels());
        });

        it("should ignore display units left over from a standard display", () => {
            expectPercentages(renderAxisLabels((settings: Settings) => {
                settings.yAxis.displayUnits.value = DisplayUnitsType.Thousands;
            }));
        });
    });

    describe("formatting pane", () => {
        it("should hide the display units of a Stacked 100% axis only", () => {
            const dataRepresentation: IDataRepresentation = convertWithMode(twoMeasures, YAxisDisplayMode.stacked100);
            const settings: Settings = dataRepresentation.settings;

            settings.filterFormattingProperties(
                dataRepresentation,
                DataRepresentationTypeEnum.DateType,
                { getDisplayName: (key: string) => key } as powerbi.extensibility.ILocalizationManager,
                false,
            );

            expect(settings.yAxis.displayUnits.visible).toBe(false);
            expect(settings.secondaryYAxis.displayUnits.visible).toBe(true);
            expect(settings.tertiaryYAxis.displayUnits.visible).toBe(true);
        });

        it("should offer the display mode on every Y axis card", () => {
            const settings: Settings = new Settings();

            settings.getYAxes().forEach((axis) => {
                expect(axis.slices).toContain(axis.displayMode);
                expect(axis.displayMode.items.map((item) => item.value)).toEqual([
                    YAxisDisplayMode.standard,
                    YAxisDisplayMode.stacked,
                    YAxisDisplayMode.stacked100,
                ]);
            });
        });
    });
});
