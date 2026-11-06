/**
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

import { DataRepresentationAxisValueType } from "./dataRepresentationAxisValueType";

export interface IDataRepresentationPoint {
    x: DataRepresentationAxisValueType;
    y: number;
    kpiIndex: number;
    color: string;
    // Set only when the axis of the series is stacked: the bottom and the top of the band
    // the point occupies, in axis units. y keeps the raw value, which is what the
    // tooltips, the data labels and the KPI keep reading.
    stackedY0?: number;
    stackedY1?: number;
}

/**
 * The value a point is drawn at: the top of its band on a stacked axis, its own value
 * otherwise.
 */
export function getPlottedY(point: IDataRepresentationPoint): number {
    return point && point.stackedY1 !== undefined
        ? point.stackedY1
        : point && point.y;
}

/**
 * The value the area under a stacked point starts from, or undefined when the point is
 * not stacked and its area runs down to the bottom of the chart.
 */
export function getPlottedBaseline(point: IDataRepresentationPoint): number | undefined {
    return point && point.stackedY0 !== undefined
        ? point.stackedY0
        : undefined;
}

export interface IDataRepresentationPointGradientColor {
    color: string;
    points: IDataRepresentationPoint[];
}
