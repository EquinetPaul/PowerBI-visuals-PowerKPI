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

import { IDataRepresentation } from "../src/dataRepresentation/dataRepresentation";
import { IDataRepresentationSeries } from "../src/dataRepresentation/dataRepresentationSeries";
import { YAxisDescriptor, YAxisDrawOrder } from "../src/settings/descriptors/axis/yAxisDescriptor";
import { Settings } from "../src/settings/settings";

import {
    buildValueRolesDataView,
    convert,
    IValueColumn,
} from "./valueRolesDataBuilder";
import { VisualBuilder } from "./visualBuilder";

/**
 * On the Values axis, whose domain is [10, 100], Big reaches the top of the chart while
 * Small stays low, so the automatic rule draws Big first. Margin and Units are alone on
 * their axis and each spans it entirely, which ties them with Big.
 */
const columns: IValueColumn[] = [
    { role: "Values", name: "Big", values: [10, 100] },
    { role: "Values", name: "Small", values: [10, 20] },
    { role: "SecondaryValues", name: "Margin", values: [0.1, 0.3] },
    { role: "TertiaryValues", name: "Units", values: [1000, 5000] },
];

function setDrawOrder(axisSettings: YAxisDescriptor, drawOrder: YAxisDrawOrder): void {
    axisSettings.drawOrder.value = axisSettings.getNewComplexValue(
        drawOrder,
        axisSettings.drawOrder.items,
    );
}

function getDrawnNames(configureSettings?: (settings: Settings) => void): string[] {
    const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(columns), configureSettings);

    return dataRepresentation.sortedSeries.map((series: IDataRepresentationSeries) => series.name);
}

describe("Y axis draw order", () => {
    describe("series order", () => {
        it("should default to Auto and keep the automatic order", () => {
            const settings: Settings = new Settings();

            settings.getYAxes().forEach((axis: YAxisDescriptor) => {
                expect(axis.drawOrder.value.value).toBe(YAxisDrawOrder.auto);
            });

            // The series reaching highest is drawn first; ties keep the binding order
            expect(getDrawnNames()).toEqual(["Big", "Margin", "Units", "Small"]);
        });

        it("should draw the series of a Back axis first, behind every other series", () => {
            expect(getDrawnNames((settings: Settings) => {
                setDrawOrder(settings.tertiaryYAxis, YAxisDrawOrder.back);
            })).toEqual(["Units", "Big", "Margin", "Small"]);
        });

        it("should draw the series of a Front axis last, over every other series", () => {
            expect(getDrawnNames((settings: Settings) => {
                setDrawOrder(settings.yAxis, YAxisDrawOrder.front);
            })).toEqual(["Margin", "Units", "Big", "Small"]);
        });

        it("should allow any order of the three axes", () => {
            expect(getDrawnNames((settings: Settings) => {
                setDrawOrder(settings.yAxis, YAxisDrawOrder.back);
                setDrawOrder(settings.secondaryYAxis, YAxisDrawOrder.front);
            })).toEqual(["Big", "Small", "Units", "Margin"]);

            expect(getDrawnNames((settings: Settings) => {
                setDrawOrder(settings.yAxis, YAxisDrawOrder.front);
                setDrawOrder(settings.secondaryYAxis, YAxisDrawOrder.back);
            })).toEqual(["Margin", "Units", "Big", "Small"]);
        });

        it("should keep the automatic order between axes sharing a choice", () => {
            expect(getDrawnNames((settings: Settings) => {
                settings.getYAxes().forEach((axis: YAxisDescriptor) => setDrawOrder(axis, YAxisDrawOrder.front));
            })).toEqual(["Big", "Margin", "Units", "Small"]);
        });

        it("should keep the automatic order within an axis", () => {
            const names: string[] = getDrawnNames((settings: Settings) => {
                setDrawOrder(settings.yAxis, YAxisDrawOrder.front);
            });

            expect(names.indexOf("Big")).toBeLessThan(names.indexOf("Small"));
        });

        it("should not change which series are drawn", () => {
            const names: string[] = getDrawnNames((settings: Settings) => {
                setDrawOrder(settings.yAxis, YAxisDrawOrder.back);
                setDrawOrder(settings.tertiaryYAxis, YAxisDrawOrder.front);
            });

            expect(names.slice().sort()).toEqual(["Big", "Margin", "Small", "Units"]);
        });
    });

    describe("formatting pane", () => {
        it("should offer the draw order on every Y axis card, right after the display mode", () => {
            new Settings().getYAxes().forEach((axis: YAxisDescriptor) => {
                expect(axis.slices.indexOf(axis.drawOrder)).toBe(axis.slices.indexOf(axis.displayMode) + 1);
                expect(axis.drawOrder.items.map((item) => item.value)).toEqual([
                    YAxisDrawOrder.back,
                    YAxisDrawOrder.auto,
                    YAxisDrawOrder.front,
                ]);
            });
        });
    });

    describe("rendering", () => {
        function normalizeColor(color: string): string {
            const element: HTMLElement = document.createElement("div");
            element.style.color = color;

            return element.style.color;
        }

        function getDrawnStrokes(root: HTMLElement): string[] {
            return Array.from(root.querySelectorAll(".powerKpi_multiShapeComponent .powerKpi_comboComponent"))
                .map((combo: Element) => (combo.querySelector("path") as SVGPathElement).style.stroke)
                .map(normalizeColor);
        }

        function render(objects: powerbi.DataViewObjects, callback: (root: HTMLElement, visualBuilder: VisualBuilder) => void, done: DoneFn): void {
            const visualBuilder: VisualBuilder = new VisualBuilder(800, 400);
            const dataView: powerbi.DataView = buildValueRolesDataView(columns);

            dataView.metadata.objects = objects;

            visualBuilder.updateRenderTimeout(dataView, () => {
                callback(visualBuilder.element, visualBuilder);
                done();
            });
        }

        function getSeriesColor(visualBuilder: VisualBuilder, name: string): string {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const dataRepresentation: IDataRepresentation = (visualBuilder.instance as any).dataRepresentation;

            return normalizeColor(dataRepresentation.series
                .filter((series: IDataRepresentationSeries) => series.name === name)[0]
                .color);
        }

        it("should paint the series of a Front axis last in the chart", (done: DoneFn) => {
            render({ yAxis: { drawOrder: YAxisDrawOrder.front } }, (root: HTMLElement, visualBuilder: VisualBuilder) => {
                const strokes: string[] = getDrawnStrokes(root);

                expect(strokes.length).toBe(4);
                expect(strokes.slice(-2)).toEqual([
                    getSeriesColor(visualBuilder, "Big"),
                    getSeriesColor(visualBuilder, "Small"),
                ]);
            }, done);
        });

        it("should paint the series of a Back axis first in the chart", (done: DoneFn) => {
            render({ tertiaryYAxis: { drawOrder: YAxisDrawOrder.back } }, (root: HTMLElement, visualBuilder: VisualBuilder) => {
                expect(getDrawnStrokes(root)[0]).toBe(getSeriesColor(visualBuilder, "Units"));
            }, done);
        });
    });
});
