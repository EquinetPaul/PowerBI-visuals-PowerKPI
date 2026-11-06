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
import { Selection } from "d3-selection";
import {
    interfaces,
    textMeasurementService,
} from "powerbi-visuals-utils-formattingutils";
import { pixelConverter } from "powerbi-visuals-utils-typeutils";

import { IAxisTitle } from "../../../settings/descriptors/axis/axisDescriptor";

import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;

// The room kept free on both sides of a title, between it and the labels of its axis
// on one side and the edge of the axis area on the other
export const axisTitlePadding: number = 4;

export function getAxisTitleTextProperties(title: IAxisTitle, text: string = title.text): interfaces.TextProperties {
    return {
        fontFamily: title.fontFamily,
        fontSize: pixelConverter.toString(title.fontSizeInPx),
        fontStyle: title.isItalic ? "italic" : "normal",
        fontWeight: title.isBold ? "bold" : "normal",
        text,
    };
}

/**
 * The room a title takes across its axis - its line height and the padding around it -
 * or 0 when there is no title.
 */
export function getAxisTitleThickness(title: IAxisTitle | null): number {
    if (!title) {
        return 0;
    }

    return textMeasurementService.estimateSvgTextHeight(getAxisTitleTextProperties(title))
        + 2 * axisTitlePadding;
}

export interface IAxisTitleRenderOptions {
    title: IAxisTitle | null;
    // The point the middle of the title sits on
    x: number;
    y: number;
    // In degrees; -90 reads the title bottom to top, along a vertical axis
    rotation: number;
    // The length the title may run along, beyond which it is shortened with an ellipsis
    maxLength: number;
    colorPalette: ISandboxExtendedColorPalette;
}

/**
 * Draws the title of an axis into the given element, or removes it when there is none.
 */
export function renderAxisTitle(
    element: Selection<any, any, any, any>,
    className: string,
    options: IAxisTitleRenderOptions,
): void {
    const { title, x, y, rotation, maxLength, colorPalette } = options;

    const titleSelection: Selection<any, IAxisTitle, any, any> = element
        .selectAll(`text.${className}`)
        .data(title && maxLength > 0 ? [title] : []);

    titleSelection
        .exit()
        .remove();

    const isHighContrast: boolean = !!(colorPalette && colorPalette.isHighContrast);

    titleSelection
        .enter()
        .append("text")
        .classed(className, true)
        .merge(titleSelection)
        .attr("transform", `translate(${x},${y}) rotate(${rotation})`)
        .attr("text-anchor", "middle")
        .attr("dominant-baseline", "central")
        .style("font-family", (currentTitle: IAxisTitle) => currentTitle.fontFamily)
        .style("font-size", (currentTitle: IAxisTitle) => pixelConverter.toString(currentTitle.fontSizeInPx))
        .style("font-weight", (currentTitle: IAxisTitle) => currentTitle.isBold ? "bold" : "normal")
        .style("font-style", (currentTitle: IAxisTitle) => currentTitle.isItalic ? "italic" : "normal")
        .style("fill", (currentTitle: IAxisTitle) => isHighContrast
            ? colorPalette.foreground.value
            : currentTitle.color)
        .text((currentTitle: IAxisTitle) => textMeasurementService.getTailoredTextOrDefault(
            getAxisTitleTextProperties(currentTitle),
            maxLength,
        ));
}
