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
import {
    createCategoricalDataViewBuilder,
    createColorPalette,
    createSelectionIdBuilder,
} from "powerbi-visuals-utils-testutils";
import { valueType } from "powerbi-visuals-utils-typeutils";

import { DataConverter } from "../src/converter/dataConverter";
import { IDataRepresentation } from "../src/dataRepresentation/dataRepresentation";
import { AxisType } from "../src/settings/descriptors/axis/axisDescriptor";
import { YAxisDescriptor, YAxisDisplayMode } from "../src/settings/descriptors/axis/yAxisDescriptor";
import { Settings } from "../src/settings/settings";

/**
 * One measure bound to one of the value roles of the visual.
 */
export interface IValueColumn {
    role: "Values" | "SecondaryValues" | "TertiaryValues";
    name: string;
    values: number[];
    format?: string;
}

/**
 * The first days of 2016, one per value of the measures.
 */
export function getDates(amount: number): Date[] {
    const dates: Date[] = [];

    for (let index: number = 0; index < amount; index++) {
        dates.push(new Date(2016, 0, index + 1));
    }

    return dates;
}

/**
 * Builds an ungrouped data view - a date axis and one column per measure - with values
 * chosen by the spec, so that stacked bands and axis domains can be asserted exactly.
 */
export function buildValueRolesDataView(columns: IValueColumn[]): powerbi.DataView {
    const amountOfPoints: number = columns.reduce(
        (amount: number, column: IValueColumn) => Math.max(amount, column.values.length),
        0,
    );

    return createCategoricalDataViewBuilder()
        .withCategory({
            source: {
                displayName: "Axis",
                format: "%M/%d/yyyy",
                roles: { Axis: true },
                type: valueType.ValueType.fromDescriptor({ dateTime: true }),
            },
            values: getDates(amountOfPoints),
            identityFrom: { fields: [] },
        })
        .withValues({
            columns: columns.map((column: IValueColumn) => ({
                source: {
                    displayName: column.name,
                    format: column.format,
                    queryName: `Sum(${column.name})`,
                    roles: { [column.role]: true },
                    type: valueType.ValueType.fromDescriptor({ numeric: true }),
                },
                values: column.values,
            })),
        })
        .build();
}

export function setDisplayMode(axisSettings: YAxisDescriptor, mode: YAxisDisplayMode): void {
    axisSettings.displayMode.value = axisSettings.getNewComplexValue(
        mode,
        axisSettings.displayMode.items,
    );
}

/**
 * Runs the data view through the converter the way the visual does, with the settings
 * the spec configured.
 */
export function convert(
    dataView: powerbi.DataView,
    configureSettings?: (settings: Settings) => void,
    viewport: powerbi.IViewport = { height: 300, width: 400 },
): IDataRepresentation {
    const settings: Settings = new Settings();

    if (configureSettings) {
        configureSettings(settings);
    }

    const converter: DataConverter = new DataConverter({
        colorPalette: createColorPalette(),
        createSelectionIdBuilder,
    });

    converter.getAxisType({
        dataView,
        xAxisType: AxisType.continuous,
    });

    return converter.convert({
        dataView,
        hasSelection: false,
        locale: "en-US",
        settings,
        viewport,
    });
}
