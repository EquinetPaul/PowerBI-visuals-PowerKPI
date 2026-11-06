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
import { Settings } from "../src/settings/settings";

import {
    buildValueRolesDataView,
    convert,
    IValueColumn,
} from "./valueRolesDataBuilder";
import { VisualBuilder } from "./visualBuilder";

const threeRoles: IValueColumn[] = [
    { role: "Values", name: "Sales", values: [10, 20, 30, 25, 40] },
    { role: "SecondaryValues", name: "Margin", values: [0.1, 0.2, 0.15, 0.3, 0.25], format: "0%" },
    { role: "TertiaryValues", name: "Units", values: [1000, 4000, 2000, 8000, 5000] },
];

describe("Tertiary Y axis", () => {
    describe("data", () => {
        it("should put the Tertiary Values in a third group with its own domain", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(threeRoles));

            expect(dataRepresentation.groups.length).toBe(3);
            expect(dataRepresentation.groups[2].series.map((series: IDataRepresentationSeries) => series.name))
                .toEqual(["Units"]);
            expect(dataRepresentation.groups[2].y.min).toBe(1000);
            expect(dataRepresentation.groups[2].y.max).toBe(8000);

            // The other axes keep their own range
            expect(dataRepresentation.groups[0].y.max).toBe(40);
            expect(dataRepresentation.groups[1].y.max).toBe(0.3);
        });

        it("should give the third group its own scale", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(threeRoles));

            expect(dataRepresentation.groups[2].y.scale.getDomain()).toEqual([1000, 8000]);
            expect(dataRepresentation.groups[2].series[0].y).toBe(dataRepresentation.groups[2].y);
        });

        it("should list the tertiary series after the others", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(threeRoles));

            expect(dataRepresentation.series.map((series: IDataRepresentationSeries) => series.name))
                .toEqual(["Sales", "Margin", "Units"]);
        });

        it("should apply the Min and Max of the tertiary axis card to the third group only", () => {
            const dataRepresentation: IDataRepresentation = convert(
                buildValueRolesDataView(threeRoles),
                (settings: Settings) => {
                    settings.tertiaryYAxis.min.value = 0;
                    settings.tertiaryYAxis.max.value = 10000;
                },
            );

            expect(dataRepresentation.groups[2].y.scale.getDomain()).toEqual([0, 10000]);
            expect(dataRepresentation.groups[0].y.scale.getDomain()).toEqual([10, 40]);
        });

        it("should work with Tertiary Values alone", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView([threeRoles[2]]));

            expect(dataRepresentation.groups[0]).toBeUndefined();
            expect(dataRepresentation.groups[2].series.length).toBe(1);
            expect(dataRepresentation.series[0].current.y).toBe(5000);
        });
    });

    describe("formatting pane", () => {
        it("should hide the tertiary cards while no Tertiary Values are bound", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(threeRoles.slice(0, 2)));
            const settings: Settings = new Settings();

            expect(settings.shouldHideSettingsCard("tertiaryYAxis", dataRepresentation)).toBe(true);
            expect(settings.shouldHideSettingsCard("tertiaryReferenceLineOfYAxis", dataRepresentation)).toBe(true);
        });

        it("should show the tertiary cards once Tertiary Values are bound", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(threeRoles));
            const settings: Settings = new Settings();

            expect(settings.shouldHideSettingsCard("tertiaryYAxis", dataRepresentation)).toBe(false);
            expect(settings.shouldHideSettingsCard("tertiaryReferenceLineOfYAxis", dataRepresentation)).toBe(false);
        });

        it("should list the tertiary cards right after their secondary counterparts", () => {
            const names: string[] = new Settings().cards.map((card) => card.name);

            expect(names.indexOf("tertiaryYAxis")).toBe(names.indexOf("secondaryYAxis") + 1);
            expect(names.indexOf("tertiaryReferenceLineOfYAxis"))
                .toBe(names.indexOf("secondaryReferenceLineOfYAxis") + 1);
        });
    });

    describe("layout", () => {
        // Sub pixel rounding makes exact comparisons brittle
        const tolerance: number = 1;

        function render(
            columns: IValueColumn[],
            callback: (root: HTMLElement) => void,
            done: DoneFn,
            objects: powerbi.DataViewObjects = {},
            width: number = 1024,
            height: number = 768,
        ): void {
            const visualBuilder: VisualBuilder = new VisualBuilder(width, height);
            const dataView: powerbi.DataView = buildValueRolesDataView(columns);

            dataView.metadata.objects = objects;

            visualBuilder.updateRenderTimeout(dataView, () => {
                callback(visualBuilder.element);
                done();
            });
        }

        function getYAxisRects(root: HTMLElement): DOMRect[] {
            return Array.from(root.querySelectorAll(".powerKpi_plot svg.powerKpi_visualYAxis"))
                .map((element: Element) => element.getBoundingClientRect());
        }

        function getChartRect(root: HTMLElement): DOMRect {
            return root.querySelector(".powerKpi_plot svg.powerKpi_svgComponent").getBoundingClientRect();
        }

        function getTickLabels(axisElement: Element): string[] {
            return Array.from(axisElement.querySelectorAll(".tick text"))
                .map((element: Element) => element.textContent);
        }

        it("should render three Y axes: one on the left, two side by side on the right", (done: DoneFn) => {
            render(threeRoles, (root: HTMLElement) => {
                const [left, secondary, tertiary] = getYAxisRects(root);
                const chart: DOMRect = getChartRect(root);

                expect(left.width).toBeGreaterThan(0);
                expect(secondary.width).toBeGreaterThan(0);
                expect(tertiary.width).toBeGreaterThan(0);

                expect(left.right).toBeLessThanOrEqual(chart.left + tolerance);
                expect(secondary.left).toBeGreaterThanOrEqual(chart.right - tolerance);
                expect(tertiary.left).toBeGreaterThanOrEqual(secondary.right - tolerance);
            }, done);
        });

        it("should keep the chart and every axis inside the plot", (done: DoneFn) => {
            render(threeRoles, (root: HTMLElement) => {
                const plot: DOMRect = root.querySelector(".powerKpi_plot").getBoundingClientRect();
                const rects: DOMRect[] = [getChartRect(root), ...getYAxisRects(root)];

                rects.forEach((rect: DOMRect) => {
                    expect(rect.left).toBeGreaterThanOrEqual(plot.left - tolerance);
                    expect(rect.right).toBeLessThanOrEqual(plot.right + tolerance);
                });

                expect(getChartRect(root).width).toBeGreaterThan(0);
            }, done, {}, 400, 250);
        });

        it("should label the tertiary axis with the range of the Tertiary Values", (done: DoneFn) => {
            render(threeRoles, (root: HTMLElement) => {
                const axes: Element[] = Array.from(root.querySelectorAll(".powerKpi_plot svg.powerKpi_visualYAxis"));
                const tertiaryLabels: string[] = getTickLabels(axes[2]);

                expect(tertiaryLabels.length).toBeGreaterThan(1);
                expect(tertiaryLabels.some((label: string) => /K$/.test(label))).toBe(true);
            }, done);
        });

        it("should take no room for the tertiary axis while it is turned off", (done: DoneFn) => {
            render(threeRoles, (root: HTMLElement) => {
                const tertiary: DOMRect = getYAxisRects(root)[2];

                expect(tertiary.width).toBe(0);
            }, done, { tertiaryYAxis: { show: false } });
        });

        it("should take no room for the tertiary axis while no Tertiary Values are bound", (done: DoneFn) => {
            render(threeRoles.slice(0, 2), (root: HTMLElement) => {
                const tertiary: DOMRect = getYAxisRects(root)[2];

                expect(tertiary.width).toBe(0);
            }, done);
        });
    });
});
