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
import {
    createDefaultZoom,
    IDataRepresentationZoom,
    IDataRepresentationZoomRange,
} from "../src/dataRepresentation/dataRepresentationZoom";
import { PowerKPI } from "../src/visual";
import { ZoomSliderAxis } from "../src/visualComponent/zoomSliderComponent";

import { buildValueRolesDataView, IValueColumn } from "./valueRolesDataBuilder";
import { VisualBuilder } from "./visualBuilder";

const sales: IValueColumn = {
    name: "Sales",
    role: "Values",
    values: [10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
};

const sliderOnBothAxes: powerbi.DataViewObjects = {
    zoomSlider: { show: true, showForXAxis: true, showForYAxis: true },
};

/**
 * Reaches into the visual the way the zoom slider does: the slider dispatches its range
 * once a drag ends, and the visual keeps it and re-renders.
 */
interface IZoomableVisual {
    zoom: IDataRepresentationZoom;
    dataRepresentation: IDataRepresentation;
    applyZoom(axis: ZoomSliderAxis, range: IDataRepresentationZoomRange): void;
}

function getVisual(visualBuilder: VisualBuilder): IZoomableVisual {
    return visualBuilder.instance as unknown as IZoomableVisual;
}

function createDataView(
    objects: powerbi.DataViewObjects = sliderOnBothAxes,
    columns: IValueColumn[] = [sales],
): powerbi.DataView {
    const dataView: powerbi.DataView = buildValueRolesDataView(columns);

    dataView.metadata.objects = objects;

    return dataView;
}

function createZoomedVisual(): VisualBuilder {
    const visualBuilder: VisualBuilder = new VisualBuilder(800, 400);

    visualBuilder.update(createDataView());
    getVisual(visualBuilder).applyZoom(ZoomSliderAxis.x, { end: 1, start: 0.5 });
    getVisual(visualBuilder).applyZoom(ZoomSliderAxis.y, { end: 0.8, start: 0.2 });

    return visualBuilder;
}

function expectZoomKept(visualBuilder: VisualBuilder): void {
    expect(getVisual(visualBuilder).zoom).toEqual({
        x: { end: 1, start: 0.5 },
        y: { end: 0.8, start: 0.2 },
    });
}

describe("Zoom slider persistence", () => {
    it("should narrow the axes to the zoomed range", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();
        const dataRepresentation: IDataRepresentation = getVisual(visualBuilder).dataRepresentation;

        // The second half of the ten days of data
        expect((dataRepresentation.x.min as Date).getTime()).toBeGreaterThan(new Date(2016, 0, 5).getTime());
        expectZoomKept(visualBuilder);
    });

    it("should keep the zoom when the visual is resized", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();

        visualBuilder.viewport = { height: 300, width: 500 };
        visualBuilder.update(createDataView());

        expectZoomKept(visualBuilder);
        expect((getVisual(visualBuilder).dataRepresentation.x.min as Date).getTime())
            .toBeGreaterThan(new Date(2016, 0, 5).getTime());
    });

    it("should keep the zoom when a formatting setting changes", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();

        visualBuilder.update(createDataView({
            ...sliderOnBothAxes,
            yAxis: { fontSize: 14 },
        }));

        expectZoomKept(visualBuilder);
    });

    it("should keep the zoom when a filter changes the values of the same fields", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();

        visualBuilder.update(createDataView(sliderOnBothAxes, [{ ...sales, values: [5, 4, 3, 2, 1] }]));

        expectZoomKept(visualBuilder);
    });

    it("should reset the zoom when other fields are bound", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();

        visualBuilder.update(createDataView(sliderOnBothAxes, [
            sales,
            { name: "Cost", role: "SecondaryValues", values: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10] },
        ]));

        expect(getVisual(visualBuilder).zoom).toEqual(createDefaultZoom());
    });

    it("should drop the zoom of an axis whose slider is turned off", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();

        visualBuilder.update(createDataView({
            zoomSlider: { show: true, showForXAxis: true, showForYAxis: false },
        }));

        expect(getVisual(visualBuilder).zoom).toEqual({
            x: { end: 1, start: 0.5 },
            y: { end: 1, start: 0 },
        });
    });

    it("should drop the whole zoom when the zoom slider card is turned off", () => {
        const visualBuilder: VisualBuilder = createZoomedVisual();

        visualBuilder.update(createDataView({ zoomSlider: { show: false } }));

        expect(getVisual(visualBuilder).zoom).toEqual(createDefaultZoom());
        expect((getVisual(visualBuilder).dataRepresentation.x.min as Date).getTime())
            .toBe(new Date(2016, 0, 1).getTime());
    });

    describe("bound fields key", () => {
        it("should not depend on the values a filter leaves", () => {
            expect(PowerKPI.getBoundFieldsKey(createDataView(sliderOnBothAxes, [sales])))
                .toBe(PowerKPI.getBoundFieldsKey(createDataView(sliderOnBothAxes, [{ ...sales, values: [1] }])));
        });

        it("should change when a measure moves to another role", () => {
            expect(PowerKPI.getBoundFieldsKey(createDataView(sliderOnBothAxes, [sales])))
                .not.toBe(PowerKPI.getBoundFieldsKey(createDataView(sliderOnBothAxes, [{ ...sales, role: "TertiaryValues" }])));
        });

        it("should cope with an update that has no data", () => {
            expect(PowerKPI.getBoundFieldsKey(undefined)).toBe("");
        });
    });
});
