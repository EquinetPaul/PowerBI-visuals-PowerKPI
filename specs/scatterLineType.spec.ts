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
import { select as d3Select } from "d3-selection";
import { Selection } from "d3-selection";

import { testDom, createColorPalette } from "powerbi-visuals-utils-testutils";

import { ComboComponent, IComboComponentRenderOptions } from "../src/visualComponent/combo/comboComponent";
import { IDataRepresentationPoint, IDataRepresentationPointGradientColor } from "../src/dataRepresentation/dataRepresentationPoint";
import { IDataRepresentationSeries } from "../src/dataRepresentation/dataRepresentationSeries";
import { DataRepresentationScale } from "../src/dataRepresentation/dataRepresentationScale";
import { DataRepresentationTypeEnum } from "../src/dataRepresentation/dataRepresentationType";
import { LineInterpolation, LineStyle, LineType } from "../src/settings/descriptors/line/lineTypes";

const seriesColor: string = "#01B8AA";

function buildPoints(yValues: number[]): IDataRepresentationPoint[] {
    return yValues.map((y: number, index: number) => ({
        color: seriesColor,
        kpiIndex: NaN,
        x: index,
        y,
    }));
}

/**
 * Builds the minimal series shape the chart marks read: the points themselves, the
 * gradient colour segments the line/area paths are built from, and the Y scale.
 */
function buildSeries(points: IDataRepresentationPoint[]): IDataRepresentationSeries {
    const validYValues: number[] = points
        .map((point: IDataRepresentationPoint) => point.y)
        .filter((y: number) => y !== null && !isNaN(y));

    const gradientPoints: IDataRepresentationPointGradientColor[] = [{
        color: seriesColor,
        points,
    }];

    return {
        color: seriesColor,
        gradientPoints,
        hasSelection: false,
        points,
        selected: false,
        y: {
            format: null,
            max: Math.max(...validYValues),
            min: Math.min(...validYValues),
            scale: DataRepresentationScale
                .create()
                .domain([Math.min(...validYValues), Math.max(...validYValues)], DataRepresentationTypeEnum.NumberType),
        },
    } as IDataRepresentationSeries;
}

function renderCombo(
    lineType: LineType,
    yValues: number[] = [10, 20, 30, 40],
): Element {
    const viewport: powerbi.IViewport = { width: 400, height: 300 };

    const rootElement: HTMLElement = testDom(viewport.height.toString(), viewport.width.toString());
    const element: Selection<HTMLElement, any, any, any> = d3Select(rootElement);

    const points: IDataRepresentationPoint[] = buildPoints(yValues);
    const series: IDataRepresentationSeries = buildSeries(points);

    const options: IComboComponentRenderOptions = {
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
            .domain([0, points.length - 1], DataRepresentationTypeEnum.NumberType),
        y: series.y.scale,
    };

    const comboComponent: ComboComponent = new ComboComponent({ element });

    comboComponent.render(options);

    return rootElement;
}

function getCircles(rootElement: Element): SVGCircleElement[] {
    return Array.from(rootElement.querySelectorAll("circle"));
}

function getPaths(rootElement: Element): SVGPathElement[] {
    return Array.from(rootElement.querySelectorAll("path"));
}

describe("Scatter line type", () => {
    it("should render one circle per data point and no line path", () => {
        const yValues: number[] = [10, 20, 30, 40];

        const rootElement: Element = renderCombo(LineType.scatter, yValues);

        expect(getCircles(rootElement).length).toBe(yValues.length);
        expect(getPaths(rootElement).length).toBe(0);
    });

    it("should render a line path and no circles for the Line type", () => {
        const rootElement: Element = renderCombo(LineType.line);

        expect(getPaths(rootElement).length).toBe(1);
        expect(getCircles(rootElement).length).toBe(0);
    });

    it("should render an area path plus its line and no circles for the Area type", () => {
        const rootElement: Element = renderCombo(LineType.area);

        // AreaComponent draws the filled area and then delegates to LineComponent
        expect(getPaths(rootElement).length).toBe(2);
        expect(getCircles(rootElement).length).toBe(0);
    });

    it("should place every circle at finite coordinates", () => {
        const rootElement: Element = renderCombo(LineType.scatter);

        const circles: SVGCircleElement[] = getCircles(rootElement);

        expect(circles.length).toBeGreaterThan(0);

        circles.forEach((circle: SVGCircleElement) => {
            const cx: number = parseFloat(circle.getAttribute("cx"));
            const cy: number = parseFloat(circle.getAttribute("cy"));
            const r: number = parseFloat(circle.getAttribute("r"));

            expect(isFinite(cx)).toBe(true);
            expect(isFinite(cy)).toBe(true);
            expect(r).toBeGreaterThan(0);
        });
    });

    it("should skip points without a value instead of placing circles at NaN coordinates", () => {
        const yValues: number[] = [10, null, 30, NaN, 50];

        const rootElement: Element = renderCombo(LineType.scatter, yValues);

        // Only the three points that carry a value are drawn
        expect(getCircles(rootElement).length).toBe(3);

        getCircles(rootElement).forEach((circle: SVGCircleElement) => {
            expect(circle.getAttribute("cy")).not.toContain("NaN");
        });
    });

    it("should colour the circles with the series colour", () => {
        const rootElement: Element = renderCombo(LineType.scatter);

        getCircles(rootElement).forEach((circle: SVGCircleElement) => {
            expect(circle.style.fill).toBeTruthy();
        });
    });
});

describe("Line and markers line type", () => {
    it("should render the line path and one marker per data point", () => {
        const yValues: number[] = [10, 20, 30, 40];
        const rootElement: Element = renderCombo(LineType.lineMarkers, yValues);

        expect(getPaths(rootElement).length).toBe(1);
        expect(getCircles(rootElement).length).toBe(yValues.length);
    });

    it("should paint the markers over the line", () => {
        const rootElement: Element = renderCombo(LineType.lineMarkers);
        const comboElement: Element = rootElement.querySelector(".powerKpi_comboComponent");
        const children: Element[] = Array.from(comboElement.children);

        expect(children.length).toBe(2);
        expect(children[0].classList.contains("powerKpi_lineComponent")).toBe(true);
        expect(children[1].classList.contains("powerKpi_scatterComponent")).toBe(true);
    });

    it("should put every marker on a vertex of the line", () => {
        const yValues: number[] = [10, 20, 30, 40];
        const rootElement: Element = renderCombo(LineType.lineMarkers, yValues);

        // d3 rounds path coordinates to three decimals, so vertices are compared loosely
        const vertices: number[][] = (getPaths(rootElement)[0].getAttribute("d").match(/-?[\d.]+,-?[\d.]+/g) || [])
            .map((pair: string) => pair.split(",").map(parseFloat));

        expect(vertices.length).toBe(yValues.length);

        getCircles(rootElement).forEach((circle: SVGCircleElement, index: number) => {
            expect(parseFloat(circle.getAttribute("cx"))).toBeCloseTo(vertices[index][0], 2);
            expect(parseFloat(circle.getAttribute("cy"))).toBeCloseTo(vertices[index][1], 2);
        });
    });

    it("should skip the markers of points without a value", () => {
        const rootElement: Element = renderCombo(LineType.lineMarkers, [10, null, 30, NaN, 50]);

        expect(getCircles(rootElement).length).toBe(3);
    });

    it("should drop the markers when switching back to the Line type", () => {
        const viewport: powerbi.IViewport = { width: 400, height: 300 };
        const rootElement: HTMLElement = testDom(viewport.height.toString(), viewport.width.toString());
        const points: IDataRepresentationPoint[] = buildPoints([10, 20, 30]);
        const series: IDataRepresentationSeries = buildSeries(points);
        const comboComponent: ComboComponent = new ComboComponent({ element: d3Select(rootElement) });

        const options: IComboComponentRenderOptions = {
            areaOpacity: 0.5,
            colorPalette: createColorPalette(),
            gradientPoints: series.gradientPoints,
            interpolation: LineInterpolation.linear,
            lineStyle: LineStyle.solidLine,
            lineType: LineType.lineMarkers,
            opacity: 1,
            radiusFactor: 1.4,
            series,
            thickness: 2,
            viewport,
            x: DataRepresentationScale
                .create()
                .domain([0, points.length - 1], DataRepresentationTypeEnum.NumberType),
            y: series.y.scale,
        };

        comboComponent.render(options);
        expect(getCircles(rootElement).length).toBe(3);

        comboComponent.render({ ...options, lineType: LineType.line });
        expect(getCircles(rootElement).length).toBe(0);
        expect(getPaths(rootElement).length).toBe(1);
    });
});
