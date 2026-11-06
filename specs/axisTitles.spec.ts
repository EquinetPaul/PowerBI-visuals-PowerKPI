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
import { AxisDescriptor, IAxisTitle } from "../src/settings/descriptors/axis/axisDescriptor";
import { Settings } from "../src/settings/settings";

import { DataBuilder } from "./dataBuilder";
import { buildValueRolesDataView, convert, IValueColumn } from "./valueRolesDataBuilder";
import { VisualBuilder } from "./visualBuilder";

const columns: IValueColumn[] = [
    { role: "Values", name: "Sales", values: [10, 20, 30, 25, 40] },
    { role: "Values", name: "Cost", values: [5, 8, 12, 10, 15] },
    { role: "SecondaryValues", name: "Margin", values: [0.1, 0.2, 0.15, 0.3, 0.25], format: "0%" },
    { role: "TertiaryValues", name: "Units", values: [1000, 4000, 2000, 8000, 5000] },
];

const allTitlesOn: powerbi.DataViewObjects = {
    secondaryYAxis: { showTitle: true },
    tertiaryYAxis: { showTitle: true },
    xAxis: { showTitle: true },
    yAxis: { showTitle: true },
};

describe("Axis titles", () => {
    describe("settings", () => {
        function getCards(settings: Settings): AxisDescriptor[] {
            return [settings.xAxis, ...settings.getYAxes()];
        }

        it("should be off by default on every axis, leaving existing reports unchanged", () => {
            getCards(new Settings()).forEach((card: AxisDescriptor) => {
                expect(card.showTitle.value).toBe(false);
                expect(card.getTitle("Sales")).toBeNull();
            });
        });

        it("should read the default text while the author left the text empty", () => {
            const settings: Settings = new Settings();
            settings.yAxis.showTitle.value = true;

            expect(settings.yAxis.getTitle("Sales, Cost").text).toBe("Sales, Cost");
        });

        it("should read the text the author typed instead of the default one", () => {
            const settings: Settings = new Settings();
            settings.xAxis.showTitle.value = true;
            settings.xAxis.titleText.value = "  Order date  ";

            expect(settings.xAxis.getTitle("Date").text).toBe("Order date");
        });

        it("should carry the font and colour of the title settings", () => {
            const settings: Settings = new Settings();
            const card: AxisDescriptor = settings.tertiaryYAxis;

            card.showTitle.value = true;
            card.titleFontSize.value = 12;
            card.titleFontFamily.value = "Arial";
            card.titleBold.value = true;
            card.titleItalic.value = true;
            card.titleColor.value = { value: "#5C2D91" };

            const title: IAxisTitle = card.getTitle("Units");

            expect(title.fontSizeInPx).toBe(16);
            expect(title.fontFamily).toBe("Arial");
            expect(title.isBold).toBe(true);
            expect(title.isItalic).toBe(true);
            expect(title.color).toBe("#5C2D91");
        });

        it("should draw no title when there is no text to show", () => {
            const settings: Settings = new Settings();
            settings.yAxis.showTitle.value = true;

            expect(settings.yAxis.getTitle("")).toBeNull();
            expect(settings.yAxis.getTitle(undefined)).toBeNull();
        });

        it("should draw no title while the axis itself is turned off", () => {
            const settings: Settings = new Settings();
            settings.yAxis.showTitle.value = true;
            settings.yAxis.show.value = false;

            expect(settings.yAxis.getTitle("Sales")).toBeNull();
        });

        it("should offer the title settings on the four axis cards", () => {
            getCards(new Settings()).forEach((card: AxisDescriptor) => {
                [card.showTitle, ...card.getTitleSlices()].forEach((slice) => {
                    expect(card.slices).toContain(slice);
                });
            });
        });
    });

    describe("default text", () => {
        it("should name each Y axis after its measures", () => {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(columns));

            expect(dataRepresentation.groups[0].y.name).toBe("Sales, Cost");
            expect(dataRepresentation.groups[1].y.name).toBe("Margin");
            expect(dataRepresentation.groups[2].y.name).toBe("Units");
            expect(dataRepresentation.x.name).toBe("Axis");
        });

        it("should list each measure once when a legend splits it into a series per group", () => {
            const dataBuilder: DataBuilder = new DataBuilder();
            const dataRepresentation: IDataRepresentation = convert(
                dataBuilder.getGroupedDataView(["GroupA", "GroupB", "GroupC"], ["Sales", "Cost"]),
            );

            expect(dataRepresentation.groups[0].series.length).toBe(6);
            expect(dataRepresentation.groups[0].y.name).toBe("Sales, Cost");
        });
    });

    describe("formatting pane", () => {
        function filter(configureSettings: (settings: Settings) => void): Settings {
            const dataRepresentation: IDataRepresentation = convert(buildValueRolesDataView(columns), configureSettings);
            const settings: Settings = dataRepresentation.settings;

            settings.filterFormattingProperties(
                dataRepresentation,
                dataRepresentation.x.axisType,
                { getDisplayName: (key: string) => key } as powerbi.extensibility.ILocalizationManager,
                false,
            );

            return settings;
        }

        it("should only show the title settings once the title is on", () => {
            const settings: Settings = filter((currentSettings: Settings) => {
                currentSettings.secondaryYAxis.showTitle.value = true;
            });

            settings.yAxis.getTitleSlices().forEach((slice) => expect(slice.visible).toBe(false));
            settings.secondaryYAxis.getTitleSlices().forEach((slice) => expect(slice.visible).toBe(true));
        });

        it("should hint at the default text in the text input", () => {
            const settings: Settings = filter(() => undefined);

            expect(settings.xAxis.titleText.placeholder).toBe("Axis");
            expect(settings.yAxis.titleText.placeholder).toBe("Sales, Cost");
            expect(settings.tertiaryYAxis.titleText.placeholder).toBe("Units");
        });
    });

    describe("layout", () => {
        // Sub pixel rounding makes exact comparisons brittle
        const tolerance: number = 1;

        function render(
            objects: powerbi.DataViewObjects,
            callback: (root: HTMLElement) => void,
            done: DoneFn,
            width: number = 1024,
            height: number = 600,
        ): void {
            const visualBuilder: VisualBuilder = new VisualBuilder(width, height);
            const dataView: powerbi.DataView = buildValueRolesDataView(columns);

            dataView.metadata.objects = objects;

            visualBuilder.updateRenderTimeout(dataView, () => {
                callback(visualBuilder.element);
                done();
            });
        }

        function getTitles(root: HTMLElement): SVGTextElement[] {
            return Array.from(root.querySelectorAll("text.powerKpi_axisTitle"));
        }

        function getYAxes(root: HTMLElement): Element[] {
            return Array.from(root.querySelectorAll(".powerKpi_plot svg.powerKpi_visualYAxis"));
        }

        function getTickTextsRect(axisElement: Element): { left: number; right: number; bottom: number } {
            const rects: DOMRect[] = Array.from(axisElement.querySelectorAll(".tick text"))
                .map((element: Element) => element.getBoundingClientRect());

            return {
                bottom: Math.max(...rects.map((rect: DOMRect) => rect.bottom)),
                left: Math.min(...rects.map((rect: DOMRect) => rect.left)),
                right: Math.max(...rects.map((rect: DOMRect) => rect.right)),
            };
        }

        function getChartRect(root: HTMLElement): DOMRect {
            return root.querySelector(".powerKpi_plot svg.powerKpi_svgComponent").getBoundingClientRect();
        }

        it("should draw no title while every title is off", (done: DoneFn) => {
            render({}, (root: HTMLElement) => {
                expect(getTitles(root).length).toBe(0);
            }, done);
        });

        it("should draw a title for each of the four axes", (done: DoneFn) => {
            render(allTitlesOn, (root: HTMLElement) => {
                expect(getTitles(root).map((title: SVGTextElement) => title.textContent).sort())
                    .toEqual(["Axis", "Margin", "Sales, Cost", "Units"]);
            }, done);
        });

        it("should put the title of the left axis on the left of its labels", (done: DoneFn) => {
            render(allTitlesOn, (root: HTMLElement) => {
                const axis: Element = getYAxes(root)[0];
                const title: DOMRect = axis.querySelector("text.powerKpi_axisTitle").getBoundingClientRect();

                expect(title.width).toBeGreaterThan(0);
                expect(title.height).toBeGreaterThan(title.width);
                expect(title.right).toBeLessThanOrEqual(getTickTextsRect(axis).left + tolerance);
                expect(title.left).toBeGreaterThanOrEqual(axis.getBoundingClientRect().left - tolerance);
            }, done);
        });

        it("should put the titles of the right axes on the right of their labels, away from the chart", (done: DoneFn) => {
            render(allTitlesOn, (root: HTMLElement) => {
                const [, secondary, tertiary] = getYAxes(root);

                [secondary, tertiary].forEach((axis: Element) => {
                    const title: DOMRect = axis.querySelector("text.powerKpi_axisTitle").getBoundingClientRect();

                    expect(title.left).toBeGreaterThanOrEqual(getTickTextsRect(axis).right - tolerance);
                    expect(title.right).toBeLessThanOrEqual(axis.getBoundingClientRect().right + tolerance);
                });

                // The secondary title stays clear of the tertiary axis next to it
                const secondaryTitle: DOMRect = secondary.querySelector("text.powerKpi_axisTitle").getBoundingClientRect();
                expect(secondaryTitle.right).toBeLessThanOrEqual(getTickTextsRect(tertiary).left + tolerance);
            }, done);
        });

        it("should put the title of the X axis underneath its labels", (done: DoneFn) => {
            render(allTitlesOn, (root: HTMLElement) => {
                const axis: Element = root.querySelector(".visualXAxis svg");
                const title: DOMRect = axis.querySelector("text.powerKpi_axisTitle").getBoundingClientRect();
                const chart: DOMRect = getChartRect(root);

                expect(title.top).toBeGreaterThanOrEqual(getTickTextsRect(axis).bottom - tolerance);
                expect(title.bottom).toBeLessThanOrEqual(axis.getBoundingClientRect().bottom + tolerance);
                // Centred under the drawing area
                expect(Math.abs((title.left + title.right) / 2 - (chart.left + chart.right) / 2)).toBeLessThan(20);
            }, done);
        });

        it("should take the room of the titles from the chart, leaving it inside the plot", async () => {
            const renderChart = (objects: powerbi.DataViewObjects): Promise<HTMLElement> => new Promise(
                (resolve: (root: HTMLElement) => void) => render(objects, resolve, (() => undefined) as unknown as DoneFn),
            );

            const withoutTitles: DOMRect = getChartRect(await renderChart({}));
            const root: HTMLElement = await renderChart(allTitlesOn);
            const withTitles: DOMRect = getChartRect(root);
            const plot: DOMRect = root.querySelector(".powerKpi_plot").getBoundingClientRect();

            expect(withTitles.width).toBeLessThan(withoutTitles.width);
            expect(withTitles.height).toBeLessThan(withoutTitles.height);
            expect(withTitles.right).toBeLessThanOrEqual(plot.right + tolerance);
        });

        it("should shorten a title longer than its axis", (done: DoneFn) => {
            render({
                yAxis: {
                    showTitle: true,
                    titleText: "A very long title that could never fit along the height of this small axis",
                },
            }, (root: HTMLElement) => {
                const axis: Element = getYAxes(root)[0];
                const title: SVGTextElement = axis.querySelector("text.powerKpi_axisTitle");

                expect(title.textContent).toContain("...");
                expect(title.getBoundingClientRect().height).toBeLessThanOrEqual(axis.getBoundingClientRect().height + tolerance);
            }, done, 400, 250);
        });

        it("should draw the title with its font and colour", (done: DoneFn) => {
            render({
                tertiaryYAxis: {
                    showTitle: true,
                    titleBold: true,
                    titleColor: { solid: { color: "#5C2D91" } },
                    titleFontSize: 12,
                },
            }, (root: HTMLElement) => {
                const title: SVGTextElement = getYAxes(root)[2].querySelector("text.powerKpi_axisTitle");

                expect(title.style.fontWeight).toBe("bold");
                expect(title.style.fontSize).toBe("16px");
                expect(title.style.fill).toBe("rgb(92, 45, 145)");
            }, done);
        });
    });
});
